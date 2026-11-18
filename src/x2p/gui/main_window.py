"""Main window: sheets + settings on one side, the real PDF preview in the middle,
the verification report below. Arabic (RTL) and English (LTR)."""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

import pymupdf
from PySide6.QtCore import QLibraryInfo, QSettings, Qt, QTimer, QTranslator, QUrl
from PySide6.QtGui import QAction, QDesktopServices, QImage, QKeySequence, QPixmap
from PySide6.QtWidgets import (QApplication, QCheckBox, QComboBox, QDoubleSpinBox, QFileDialog,
                               QFormLayout, QFrame, QGroupBox, QHBoxLayout, QHeaderView, QLabel,
                               QMainWindow, QMessageBox, QProgressBar, QPushButton, QScrollArea,
                               QSplitter, QToolBar, QTreeWidget, QTreeWidgetItem, QVBoxLayout,
                               QWidget)

from .. import tempstore
from ..engines import engine_status
from ..i18n import Translator
from ..models import ConversionReport, FinalStatus, JobOptions, LayoutMode, Severity, WorkbookInfo
from ..pipeline import deliver
from ..layout import plan_sheet
from .manual_dialog import ManualDialog
from .worker import JobRunner

MODES = [LayoutMode.AUTO, LayoutMode.FIT_WIDTH, LayoutMode.ONE_PAGE, LayoutMode.KEEP_ORIGINAL,
         LayoutMode.CARDS]
SHEET_MODES = MODES + [LayoutMode.MANUAL]   # manual is per sheet only
SEV_MARK = {Severity.ERROR: "✖", Severity.WARNING: "⚠", Severity.INFO: "ℹ"}


class DropZone(QLabel):
    def __init__(self, on_file, parent=None):
        super().__init__(parent)
        self._on_file = on_file
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignCenter)
        self.setMinimumHeight(90)
        self.setWordWrap(True)
        self.setObjectName("dropZone")
        self.setStyleSheet("#dropZone{border:2px dashed palette(mid);border-radius:10px;padding:12px;}")

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        urls = [u.toLocalFile() for u in e.mimeData().urls() if u.isLocalFile()]
        if urls:
            self._on_file(urls[0])

    def mouseReleaseEvent(self, e):
        self._on_file(None)


class MainWindow(QMainWindow):
    def __init__(self, lang: str = "ar"):
        super().__init__()
        self.settings = QSettings("x2p", "ExcelToPdfProfessional")
        self.tr_ = Translator(self.settings.value("lang", lang))
        self.runner = JobRunner(self)
        self.runner.progress.connect(self._on_progress)
        self.runner.finished.connect(self._on_finished)
        self.runner.failed.connect(self._on_failed)
        self.path: str | None = None
        self.info: WorkbookInfo | None = None
        self.report: ConversionReport | None = None
        self.report_opts_key: str | None = None
        self.pdf: pymupdf.Document | None = None
        self.page = 0
        self.zoom = 1.0
        self._pending_export = False
        self._kept_dirs: list[Path] = []
        self.manual: dict[str, dict] = {}
        self._qt_translator = QTranslator(self)   # Qt's own strings (OK, Cancel, file dialogs)
        self._build()
        self._retranslate()
        self._set_state("idle")
        self.resize(1280, 820)
        if self.engine_combo.count() == 0:
            QTimer.singleShot(400, self.engine_help)

    # ------------------------------------------------------------------ engines
    LO_URL = "https://www.libreoffice.org/download/download-libreoffice/"
    WINGET = ("winget install --id TheDocumentFoundation.LibreOffice -e "
              "--accept-package-agreements --accept-source-agreements")

    def engine_help(self):
        """Re-detect engines; when none is installed offer to install LibreOffice (free)."""
        t = self.tr_
        self._retranslate()
        self._set_state("idle" if self.info is None else "inspected")
        if self.engine_combo.count():
            QMessageBox.information(self, t("app_title"), t("engine_ready", name=self.engine_combo.currentText()))
            return
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Warning)
        box.setWindowTitle(t("app_title"))
        box.setText(t("engine_missing_long"))
        b_install = box.addButton(t("install_lo"), QMessageBox.AcceptRole) if sys.platform == "win32" else None
        b_web = box.addButton(t("open_lo_site"), QMessageBox.ActionRole)
        box.addButton(t("later"), QMessageBox.RejectRole)
        box.exec()
        if b_install is not None and box.clickedButton() is b_install:
            # visible console: the user sees winget's progress and can stop it
            subprocess.Popen(["cmd", "/c", "start", "LibreOffice", "cmd", "/k", self.WINGET])
            QMessageBox.information(self, t("app_title"), t("install_started"))
        elif box.clickedButton() is b_web:
            QDesktopServices.openUrl(QUrl(self.LO_URL))

    # ------------------------------------------------------------------ layout
    def _build(self):
        tb = QToolBar()
        tb.setMovable(False)
        self.addToolBar(tb)
        self.act_open = QAction(self, shortcut=QKeySequence.Open, triggered=lambda: self.open_file(None))
        self.act_preview = QAction(self, shortcut="F5", triggered=self.preview)
        self.act_export = QAction(self, shortcut=QKeySequence.Save, triggered=self.export)
        self.act_cancel = QAction(self, shortcut="Esc", triggered=self.runner.cancel)
        self.act_lang = QAction(self, triggered=self.toggle_language)
        self.act_engines = QAction(self, triggered=self.engine_help)
        for a in (self.act_open, self.act_preview, self.act_export, self.act_cancel):
            tb.addAction(a)
        spacer = QWidget()
        spacer.setSizePolicy(spacer.sizePolicy().horizontalPolicy().Expanding,
                             spacer.sizePolicy().verticalPolicy().Preferred)
        tb.addWidget(spacer)
        tb.addAction(self.act_engines)
        tb.addAction(self.act_lang)

        # side panel
        side = QWidget()
        sv = QVBoxLayout(side)
        self.drop = DropZone(self.open_file)
        self.file_label = QLabel()
        self.file_label.setWordWrap(True)
        self.file_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.sheets_box = QGroupBox()
        sb = QVBoxLayout(self.sheets_box)
        self.sheet_tree = QTreeWidget()
        self.sheet_tree.setColumnCount(4)
        self.sheet_tree.setRootIsDecorated(False)
        self.sheet_tree.header().setSectionResizeMode(0, QHeaderView.Stretch)
        self.sheet_tree.itemChanged.connect(lambda *_: self._invalidate())
        sb.addWidget(self.sheet_tree)
        self.settings_box = QGroupBox()
        form = QFormLayout(self.settings_box)
        self.engine_combo = QComboBox()
        self.paper_combo = QComboBox()
        self.paper_combo.addItems(["A4", "Letter", "A3", "Legal"])
        self.mode_combo = QComboBox()
        self.min_font = QDoubleSpinBox(minimum=4, maximum=14, singleStep=0.5, value=8.0)
        self.hidden_chk = QCheckBox()
        self.a3_chk = QCheckBox()
        self.long_combo = QComboBox()
        self.lbl_long = QLabel()
        self.lbl_engine, self.lbl_paper, self.lbl_mode, self.lbl_min = QLabel(), QLabel(), QLabel(), QLabel()
        form.addRow(self.lbl_engine, self.engine_combo)
        form.addRow(self.lbl_paper, self.paper_combo)
        form.addRow(self.lbl_mode, self.mode_combo)
        form.addRow(self.lbl_min, self.min_font)
        form.addRow(self.lbl_long, self.long_combo)
        form.addRow(self.hidden_chk)
        form.addRow(self.a3_chk)
        for w in (self.engine_combo, self.paper_combo, self.mode_combo, self.long_combo):
            w.currentIndexChanged.connect(lambda *_: self._invalidate())
        self.min_font.valueChanged.connect(lambda *_: self._invalidate())
        self.hidden_chk.toggled.connect(self._on_hidden_toggled)
        self.a3_chk.toggled.connect(lambda *_: self._invalidate())
        sv.addWidget(self.drop)
        sv.addWidget(self.file_label)
        sv.addWidget(self.sheets_box, 1)
        sv.addWidget(self.settings_box)

        # preview
        center = QWidget()
        cv = QVBoxLayout(center)
        nav = QHBoxLayout()
        self.btn_prev, self.btn_next = QPushButton(), QPushButton()
        self.btn_zoom_out, self.btn_zoom_in, self.btn_fit = QPushButton(), QPushButton(), QPushButton()
        self.page_label = QLabel()
        self.page_label.setAlignment(Qt.AlignCenter)
        self.btn_prev.clicked.connect(lambda: self.goto(self.page - 1))
        self.btn_next.clicked.connect(lambda: self.goto(self.page + 1))
        self.btn_zoom_in.clicked.connect(lambda: self.set_zoom(self.zoom * 1.25))
        self.btn_zoom_out.clicked.connect(lambda: self.set_zoom(self.zoom / 1.25))
        self.btn_fit.clicked.connect(self.fit_width)
        for w in (self.btn_prev, self.btn_next):
            nav.addWidget(w)
        nav.addWidget(self.page_label, 1)
        for w in (self.btn_zoom_out, self.btn_zoom_in, self.btn_fit):
            nav.addWidget(w)
        self.scroll = QScrollArea()
        self.scroll.setAlignment(Qt.AlignCenter)
        self.scroll.setWidgetResizable(False)
        self.page_view = QLabel()
        self.page_view.setAlignment(Qt.AlignCenter)
        self.page_view.setFrameShape(QFrame.Box)
        self.scroll.setWidget(self.page_view)
        cv.addLayout(nav)
        cv.addWidget(self.scroll, 1)

        # report
        self.report_box = QGroupBox()
        rv = QVBoxLayout(self.report_box)
        self.summary = QLabel()
        self.summary.setWordWrap(True)
        self.summary.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.issue_tree = QTreeWidget()
        self.issue_tree.setColumnCount(3)
        self.issue_tree.setRootIsDecorated(False)
        self.issue_tree.header().setSectionResizeMode(2, QHeaderView.Stretch)
        self.issue_tree.itemDoubleClicked.connect(self._issue_clicked)
        actions = QHBoxLayout()
        self.btn_open_pdf, self.btn_open_dir = QPushButton(), QPushButton()
        self.btn_open_pdf.clicked.connect(lambda: self._open_path(self.last_output))
        self.btn_open_dir.clicked.connect(lambda: self._open_path(str(Path(self.last_output).parent)))
        actions.addStretch(1)
        actions.addWidget(self.btn_open_pdf)
        actions.addWidget(self.btn_open_dir)
        rv.addWidget(self.summary)
        rv.addWidget(self.issue_tree, 1)
        rv.addLayout(actions)
        self.last_output: str | None = None

        right = QSplitter(Qt.Vertical)
        right.addWidget(center)
        right.addWidget(self.report_box)
        right.setStretchFactor(0, 3)
        right.setStretchFactor(1, 1)
        main = QSplitter(Qt.Horizontal)
        main.addWidget(side)
        main.addWidget(right)
        main.setStretchFactor(1, 1)
        main.setSizes([360, 920])
        self.setCentralWidget(main)

        self.progress = QProgressBar(maximum=100)
        self.progress.setMaximumWidth(220)
        self.state_label = QLabel()
        self.engine_label = QLabel()
        self.statusBar().addWidget(self.state_label, 1)
        self.statusBar().addPermanentWidget(self.engine_label)
        self.statusBar().addPermanentWidget(self.progress)

    def _retranslate(self):
        t = self.tr_
        app = QApplication.instance()
        app.setLayoutDirection(Qt.RightToLeft if t.rtl else Qt.LeftToRight)
        app.removeTranslator(self._qt_translator)
        if t.lang != "en" and self._qt_translator.load(
                f"qtbase_{t.lang}", QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath)):
            app.installTranslator(self._qt_translator)
        self.setWindowTitle(t("app_title"))
        self.act_open.setText(t("open_file"))
        self.act_preview.setText(t("preview"))
        self.act_export.setText(t("export"))
        self.act_cancel.setText(t("cancel"))
        self.act_lang.setText(t("language"))
        self.act_engines.setText(t("engines"))
        self.drop.setText(t("drop_hint"))
        self.sheets_box.setTitle(t("sheets"))
        self.settings_box.setTitle(t("settings"))
        self.report_box.setTitle(t("report"))
        self.sheet_tree.setHeaderLabels([t("sheets"), "", t("mode"), t("est_pages")])
        self.lbl_long.setText(t("long_text"))
        cur_long = max(0, self.long_combo.currentIndex())
        self.long_combo.blockSignals(True)
        self.long_combo.clear()
        self.long_combo.addItem(t("long_appendix"), "appendix")
        self.long_combo.addItem(t("long_wrap"), "wrap")
        self.long_combo.setCurrentIndex(cur_long)
        self.long_combo.blockSignals(False)
        self.issue_tree.setHeaderLabels(["", t("sheets"), t("report")])
        self.lbl_engine.setText(t("engine"))
        self.lbl_paper.setText(t("paper"))
        self.lbl_mode.setText(t("mode"))
        self.lbl_min.setText(t("min_font"))
        self.hidden_chk.setText(t("include_hidden"))
        self.a3_chk.setText(t("allow_a3"))
        for b, k in ((self.btn_prev, "prev"), (self.btn_next, "next"), (self.btn_zoom_in, "zoom_in"),
                     (self.btn_zoom_out, "zoom_out"), (self.btn_fit, "fit"),
                     (self.btn_open_pdf, "open_pdf"), (self.btn_open_dir, "open_folder")):
            b.setText(t(k))
        cur = self.mode_combo.currentIndex()
        self.mode_combo.blockSignals(True)
        self.mode_combo.clear()
        self.mode_combo.addItems([t(m.value) for m in MODES])
        self.mode_combo.setCurrentIndex(max(cur, 0))
        self.mode_combo.blockSignals(False)
        status = engine_status()
        cur_engine = self.engine_combo.currentData()
        self.engine_combo.blockSignals(True)
        self.engine_combo.clear()
        names = {"libreoffice": "LibreOffice Calc", "excel": "Microsoft Excel"}
        for key, ok in status.items():
            if ok:
                self.engine_combo.addItem(names[key], key)
        if cur_engine and self.engine_combo.findData(cur_engine) >= 0:
            self.engine_combo.setCurrentIndex(self.engine_combo.findData(cur_engine))
        self.engine_combo.blockSignals(False)
        self.engine_label.setText(t("engine_ready", name=self.engine_combo.currentText())
                                  if self.engine_combo.count() else t("engine_missing"))
        if self.info:
            self._fill_sheets()
        if self.report:
            self._fill_report()
        self._update_page_label()

    def toggle_language(self):
        self.tr_.lang = "en" if self.tr_.lang == "ar" else "ar"
        self.settings.setValue("lang", self.tr_.lang)
        self._retranslate()

    # ------------------------------------------------------------------ state
    def _set_state(self, key: str):
        self.state_label.setText(self.tr_("state_" + key))
        busy = self.runner.busy
        has_engine = self.engine_combo.count() > 0
        self.act_open.setEnabled(not busy)
        self.act_preview.setEnabled(not busy and self.info is not None and has_engine)
        self.act_export.setEnabled(not busy and self.info is not None and has_engine)
        self.act_cancel.setEnabled(busy)
        self.progress.setVisible(busy and key not in ("done", "failed", "cancelled", "inspected"))
        done = bool(self.last_output)
        self.btn_open_pdf.setEnabled(done)
        self.btn_open_dir.setEnabled(done)

    def _invalidate(self):
        """Settings changed: the current preview no longer reflects them."""
        if self.report is not None and self._options_key() != self.report_opts_key:
            self.page_view.setEnabled(False)
        self._update_estimates()

    def _update_estimates(self):
        """Planner estimate per sheet (before any conversion): pages and horizontal parts."""
        if not self.info or self.sheet_tree.topLevelItemCount() != len(self.info.sheets):
            return
        opts = self._options()
        for i, s in enumerate(self.info.sheets):
            it = self.sheet_tree.topLevelItem(i)
            if s.is_empty or it.checkState(0) != Qt.Checked:
                it.setText(3, "")
                continue
            p = plan_sheet(s, opts)
            if p.keep_original:
                it.setText(3, "—")
                continue
            txt = f"~{p.expected_pages}"
            if p.expected_pages_across > 1:
                txt += f" ({p.expected_pages_across}↔)"
            if p.long_cells:
                txt += " +📎"
            it.setText(3, txt)
            it.setToolTip(3, " · ".join(self.tr_(n, count=len(p.long_cells), parts=p.expected_pages_across)
                                        for n in p.notes))
            if "cards_suggested" in p.notes:
                it.setText(3, txt + " 💡")

    # ------------------------------------------------------------------ open / inspect
    def open_file(self, path: str | None):
        if self.runner.busy:
            return
        if not path:
            start = self.settings.value("last_dir", str(Path.home()))
            path, _ = QFileDialog.getOpenFileName(self, self.tr_("open_file"), start,
                                                  "Excel (*.xlsx *.xlsm *.xls)")
            if not path:
                return
        self.settings.setValue("last_dir", str(Path(path).parent))
        self.path, self.info, self.last_output = path, None, None
        self.manual = {}
        self._clear_preview()
        p = Path(path)
        try:
            st = p.stat()
            import datetime as dt
            mod = dt.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M")
            self.file_label.setText(f"<b>{p.name}</b><br>{st.st_size / 1024:,.1f} KB — {mod}")
        except OSError:
            self.file_label.setText(p.name)
        self.runner.inspect(path)
        self._set_state("inspecting")

    def _fill_sheets(self):
        t = self.tr_
        self.sheet_tree.blockSignals(True)
        self.sheet_tree.clear()
        for s in self.info.sheets:
            dims = f"⁦{s.content.rows}×{s.content.cols}⁩" if s.content else "—"
            extras = []
            if s.merged_ranges:
                extras.append(f"⊞{s.merged_ranges}")
            if s.charts:
                extras.append(f"📊{s.charts}")
            if s.images:
                extras.append(f"🖼{s.images}")
            it = QTreeWidgetItem([s.name, f"{dims} {' '.join(extras)}  ({t(s.state)})"])
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
            checked = (s.state == "visible" or self.hidden_chk.isChecked()) and not s.is_empty
            it.setCheckState(0, Qt.Checked if checked else Qt.Unchecked)
            if s.is_empty:
                it.setDisabled(True)
                it.setToolTip(0, t("excluded_empty"))
            self.sheet_tree.addTopLevelItem(it)
            combo = QComboBox()
            combo.addItems([t(m.value) for m in SHEET_MODES])
            if s.name in self.manual:
                combo.setCurrentIndex(SHEET_MODES.index(LayoutMode.MANUAL))
            combo.setAccessibleName(f"{t('mode')} — {s.name}")
            combo.currentIndexChanged.connect(
                lambda idx, n=s.name, cb=combo: self._sheet_mode_changed(n, cb, idx))
            self.sheet_tree.setItemWidget(it, 2, combo)
        self.sheet_tree.blockSignals(False)
        self.sheet_tree.resizeColumnToContents(1)
        self._update_estimates()

    def _sheet_mode_changed(self, name: str, combo: QComboBox, idx: int):
        if SHEET_MODES[idx] == LayoutMode.MANUAL:
            si = next(s for s in self.info.sheets if s.name == name)
            opts = self._options()
            opts.per_sheet_mode[name] = LayoutMode.AUTO
            opts.manual.pop(name, None)
            dlg = ManualDialog(self.tr_, name, plan_sheet(si, opts), self.manual.get(name), self)
            if dlg.exec():
                self.manual[name] = dlg.values()
            elif name not in self.manual:
                combo.blockSignals(True)
                combo.setCurrentIndex(0)
                combo.blockSignals(False)
        else:
            self.manual.pop(name, None)
        self._invalidate()

    def _on_hidden_toggled(self, on):
        if self.info:
            for i, s in enumerate(self.info.sheets):
                if s.state != "visible" and not s.is_empty:
                    self.sheet_tree.topLevelItem(i).setCheckState(0, Qt.Checked if on else Qt.Unchecked)
        self._invalidate()

    # ------------------------------------------------------------------ convert
    def _options(self) -> JobOptions:
        selected, per_sheet = [], {}
        default_mode = MODES[max(0, self.mode_combo.currentIndex())]
        for i in range(self.sheet_tree.topLevelItemCount()):
            it = self.sheet_tree.topLevelItem(i)
            if it.checkState(0) == Qt.Checked:
                selected.append(it.text(0))
                combo = self.sheet_tree.itemWidget(it, 2)
                if combo and combo.currentIndex() > 0:
                    per_sheet[it.text(0)] = SHEET_MODES[combo.currentIndex()]
                else:
                    per_sheet[it.text(0)] = default_mode
        manual = {n: v for n, v in self.manual.items() if per_sheet.get(n) == LayoutMode.MANUAL}
        return JobOptions(sheets=selected, include_hidden=True, mode=default_mode, per_sheet_mode=per_sheet,
                          manual=manual, paper=self.paper_combo.currentText(),
                          allow_larger_paper=self.a3_chk.isChecked(), min_font_pt=self.min_font.value(),
                          engine=self.engine_combo.currentData() or "auto",
                          long_text=self.long_combo.currentData() or "appendix", lang=self.tr_.lang)

    def _options_key(self) -> str:
        o = self._options()
        return repr((o.sheets, sorted((k, v.value) for k, v in o.per_sheet_mode.items()), o.paper,
                     o.allow_larger_paper, o.min_font_pt, o.engine, o.long_text, sorted(o.manual.items())))

    def preview(self):
        if not self.path or self.runner.busy:
            return
        opts = self._options()
        if not opts.sheets:
            QMessageBox.information(self, self.tr_("app_title"), self.tr_("no_sheets_to_export"))
            return
        self.report_opts_key = self._options_key()
        self.runner.convert(self.path, opts)
        self._set_state("exporting")

    def export(self):
        if self.report and self.report.temp_pdf and self.report_opts_key == self._options_key():
            self._save_report_pdf()
        else:
            self._pending_export = True
            self.preview()

    def _save_report_pdf(self):
        t, rep = self.tr_, self.report
        if rep.status is FinalStatus.FAILED:
            if QMessageBox.warning(self, t("app_title"), t("failed_save_q"),
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
                return
        default = str(Path(self.path).with_suffix(".pdf"))
        dest, _ = QFileDialog.getSaveFileName(self, t("export"), default, "PDF (*.pdf)",
                                              options=QFileDialog.DontConfirmOverwrite)
        if not dest:
            return
        overwrite = False
        if Path(dest).exists():
            if QMessageBox.question(self, t("app_title"), t("overwrite_q")) != QMessageBox.Yes:
                return
            overwrite = True
        try:
            deliver(Path(rep.temp_pdf), Path(dest), overwrite=overwrite)
        except OSError as e:
            QMessageBox.critical(self, t("app_title"), str(e))
            return
        rep.output_pdf, rep.delivered = dest, True
        self.last_output = dest
        self._fill_report()
        self._set_state("done")

    # ------------------------------------------------------------------ signals
    def _on_progress(self, stage: str, pct: float):
        self.progress.setValue(int(pct * 100))
        self.state_label.setText(self.tr_("state_" + stage))

    def _on_finished(self, kind: str, result):
        if kind == "inspect":
            self.info = result
            self._fill_sheets()
            self._fill_issues(result.issues)
            self.summary.setText("")
            self._set_state("inspected")
            return
        self._release_preview()
        self.report = result
        if result.temp_pdf:
            self._kept_dirs.append(Path(result.temp_pdf).parent)
            self.pdf = pymupdf.open(result.temp_pdf)
            self.page = 0
            self.page_view.setEnabled(True)
            self.fit_width()
        self._fill_report()
        cancelled = any(i.code == "cancelled" for i in result.issues)
        self._set_state("cancelled" if cancelled else ("failed" if result.status is FinalStatus.FAILED else "done"))
        if self._pending_export and result.temp_pdf and not cancelled:
            self._pending_export = False
            self._save_report_pdf()
        self._pending_export = False

    def _on_failed(self, kind: str, code: str, detail: dict):
        self._pending_export = False
        self._set_state("failed")
        QMessageBox.warning(self, self.tr_("app_title"), self.tr_(code, **detail))

    # ------------------------------------------------------------------ report
    def _fill_issues(self, issues):
        self.issue_tree.clear()
        order = {Severity.ERROR: 0, Severity.WARNING: 1, Severity.INFO: 2}
        for i in sorted(issues, key=lambda x: order[x.severity]):
            it = QTreeWidgetItem([SEV_MARK[i.severity], i.sheet or "", self.tr_.issue(i)])
            it.setToolTip(0, i.severity.value)
            it.setData(0, Qt.UserRole, i.sheet)
            self.issue_tree.addTopLevelItem(it)
        self.issue_tree.resizeColumnToContents(0)

    def _fill_report(self):
        t, rep = self.tr_, self.report
        lines = [f"<b>{SEV_MARK.get(Severity.ERROR) if rep.status is FinalStatus.FAILED else '✔'} "
                 f"{t(rep.status.value)}</b>",
                 f"{rep.workbook} — {rep.total_pages} p — {rep.pdf_size_bytes / 1024:,.0f} KB — "
                 f"{rep.engine_version}"]
        if rep.min_font_pt is not None:
            lines.append(t("measured_min_font", pt=f"{rep.min_font_pt:.1f}"))
        cov = [f"{s.sheet}: {s.found_texts}/{s.expected_texts}" for s in rep.sheets if s.expected_texts]
        if cov:
            lines.append("✓ " + " · ".join(cov))
        if rep.excluded_sheets:
            lines.append(" · ".join(f"{n} ({t('excluded_' + r)})" for n, r in rep.excluded_sheets.items()))
        if rep.delivered and rep.output_pdf:
            lines.append(f"→ {rep.output_pdf}")
        self.summary.setText("<br>".join(lines))
        self._fill_issues(rep.issues)

    def _issue_clicked(self, item, _col):
        sheet = item.data(0, Qt.UserRole)
        if self.report and sheet in self.report.sheet_page_ranges:
            self.goto(self.report.sheet_page_ranges[sheet][0])

    # ------------------------------------------------------------------ preview
    def _sheet_for_page(self, p: int) -> str:
        if not self.report:
            return ""
        for name, (a, b) in self.report.sheet_page_ranges.items():
            if a <= p <= b:
                return name
        return ""

    def goto(self, p: int):
        if not self.pdf or not (0 <= p < self.pdf.page_count):
            return
        self.page = p
        self._render()

    def set_zoom(self, z: float):
        self.zoom = max(0.2, min(z, 5.0))
        self._render()

    def fit_width(self):
        if not self.pdf:
            return
        w = self.pdf[self.page].rect.width
        self.zoom = max(0.2, (self.scroll.viewport().width() - 24) / w)
        self._render()

    def _render(self):
        if not self.pdf:
            return
        dpr = self.devicePixelRatioF()
        pix = self.pdf[self.page].get_pixmap(matrix=pymupdf.Matrix(self.zoom * dpr, self.zoom * dpr), alpha=False)
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()
        qp = QPixmap.fromImage(img)
        qp.setDevicePixelRatio(dpr)
        self.page_view.setPixmap(qp)
        self.page_view.resize(qp.size() / dpr)
        self._update_page_label()

    def _update_page_label(self):
        if not self.pdf:
            self.page_label.setText("")
            return
        sheet = self._sheet_for_page(self.page)
        mark = ""
        if self.report:
            sevs = {i.severity for i in self.report.issues if i.sheet == sheet}
            mark = SEV_MARK[Severity.ERROR] + " " if Severity.ERROR in sevs else \
                (SEV_MARK[Severity.WARNING] + " " if Severity.WARNING in sevs else "")
        r = self.pdf[self.page].rect
        orient = "⇔" if r.width > r.height else "⇕"
        self.page_label.setText(mark + self.tr_("page_of", n=self.page + 1, total=self.pdf.page_count,
                                                  sheet=sheet) + f"  {orient}")

    def _release_preview(self):
        if self.pdf is not None:
            self.pdf.close()
            self.pdf = None

    def _clear_preview(self):
        self._release_preview()
        self.report = None
        self.page_view.clear()
        self.summary.setText("")
        self.issue_tree.clear()
        self.page_label.setText("")

    @staticmethod
    def _open_path(p: str | None):
        if p and os.path.exists(p):
            QDesktopServices.openUrl(QUrl.fromLocalFile(p))

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key_PageDown,):
            self.goto(self.page + 1)
        elif e.key() in (Qt.Key_PageUp,):
            self.goto(self.page - 1)
        else:
            super().keyPressEvent(e)

    def closeEvent(self, e):
        self.runner.cancel()
        self.runner.wait(8000)
        self._release_preview()
        for d in self._kept_dirs:
            tempstore.remove(d)
        super().closeEvent(e)
