"""Template section of the main window: for the selected sheet, choose the layout shape and which
rows/columns stay fixed on every page. Fields start with the program's suggestion (print titles,
frozen panes, detected header/key); only what the user changes is stored."""
from __future__ import annotations

import re

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (QCheckBox, QComboBox, QFormLayout, QGroupBox, QHBoxLayout, QLabel,
                               QLineEdit, QPushButton, QWidget)

from ..layout import plan_sheet, suggested_titles
from ..models import LayoutMode

ROWS_RE = re.compile(r"^\d+:\d+$")
COLS_RE = re.compile(r"^[A-Za-z]{1,3}:[A-Za-z]{1,3}$")


class TemplatePanel(QGroupBox):
    changed = Signal()                    # a template value changed: estimates must refresh
    shape_selected = Signal(str, int)     # sheet name, index in SHEET_MODES

    def __init__(self, window, sheet_modes):
        super().__init__()
        self.w = window
        self.modes = sheet_modes
        self.templates: dict[str, dict] = {}
        self.sheet: str | None = None
        form = QFormLayout(self)
        self.lbl_sheet = QLabel()
        self.lbl_sheet.setWordWrap(True)
        self.shape = QComboBox()
        self.rows = QLineEdit()
        self.cols = QLineEdit()
        self.rows.setPlaceholderText("1:1")
        self.cols.setPlaceholderText("A:A")
        self.order = QComboBox()
        self.rownum = QCheckBox()
        self.reset = QPushButton()
        self.summary = QLabel()
        self.summary.setWordWrap(True)
        self.error = QLabel()
        self.error.setStyleSheet("color: palette(link);")
        self.lbl_shape, self.lbl_rows, self.lbl_cols, self.lbl_order = QLabel(), QLabel(), QLabel(), QLabel()
        fixed = QWidget()
        h = QHBoxLayout(fixed)
        h.setContentsMargins(0, 0, 0, 0)
        form.addRow(self.lbl_sheet)
        form.addRow(self.lbl_shape, self.shape)
        form.addRow(self.lbl_rows, self.rows)
        form.addRow(self.lbl_cols, self.cols)
        form.addRow(self.lbl_order, self.order)
        form.addRow(self.rownum)
        form.addRow(self.reset)
        form.addRow(self.summary)
        form.addRow(self.error)
        self.shape.currentIndexChanged.connect(self._shape_changed)
        self.rows.editingFinished.connect(lambda: self._text_changed("title_rows", self.rows, ROWS_RE))
        self.cols.editingFinished.connect(lambda: self._text_changed("title_cols", self.cols, COLS_RE))
        self.order.currentIndexChanged.connect(self._order_changed)
        self.rownum.toggled.connect(self._rownum_changed)
        self.reset.clicked.connect(self._reset)
        self._loading = False
        self.setEnabled(False)

    # ---------------------------------------------------------------- texts
    def retranslate(self, t):
        self.setTitle(t("template"))
        self.lbl_shape.setText(t("template_shape"))
        self.lbl_rows.setText(t("fixed_rows"))
        self.lbl_cols.setText(t("fixed_cols"))
        self.lbl_order.setText(t("page_order"))
        self.rownum.setText(t("row_numbers"))
        self.reset.setText(t("use_suggestion"))
        self.rows.setToolTip(t("fixed_rows_tip"))
        self.cols.setToolTip(t("fixed_cols_tip"))
        self._loading = True
        cur = self.shape.currentIndex()
        self.shape.clear()
        self.shape.addItems([t(m.value) for m in self.modes])
        self.shape.setCurrentIndex(max(cur, 0))
        cur = self.order.currentIndex()
        self.order.clear()
        self.order.addItem(t("order_over"), "over")
        self.order.addItem(t("order_down"), "down")
        self.order.setCurrentIndex(max(cur, 0))
        self._loading = False
        self.refresh()

    # ---------------------------------------------------------------- state
    def clear(self):
        self.templates = {}
        self.sheet = None
        self.setEnabled(False)
        self.lbl_sheet.setText("")
        self.summary.setText("")

    def show_sheet(self, name: str | None, mode_index: int = 0):
        self.sheet = name
        self.setEnabled(name is not None)
        self._loading = True
        self.shape.setCurrentIndex(mode_index)
        self._loading = False
        self.refresh()

    def _sheet_info(self):
        if not self.sheet or not self.w.info:
            return None
        return next((s for s in self.w.info.sheets if s.name == self.sheet), None)

    def refresh(self):
        """Fill the fields with the user's values or the suggestion, and the summary line."""
        t = self.w.tr_
        si = self._sheet_info()
        self.error.setText("")
        if si is None:
            return
        self.lbl_sheet.setText(f"<b>{t('sheet')}:</b> {si.name}")
        opts = self.w._options()
        tpl = self.templates.get(si.name, {})
        s_rows, s_cols = suggested_titles(si, opts)
        self._loading = True
        self.rows.setText(s_rows or "")
        self.cols.setText(s_cols or "")
        self.order.setCurrentIndex(0 if tpl.get("page_order", "over") == "over" else 1)
        self.rownum.setChecked(tpl.get("row_numbers", True))
        self._loading = False
        mode = opts.per_sheet_mode.get(si.name, opts.mode)
        table = mode in (LayoutMode.AUTO, LayoutMode.FIT_WIDTH, LayoutMode.MANUAL)
        for w in (self.rows, self.cols, self.order, self.rownum):
            w.setEnabled(table)
        if si.is_empty or not si.col_widths_pt:
            self.summary.setText("")
            return
        p = plan_sheet(si, opts)
        if p.keep_original:
            self.summary.setText(t("tpl_keep"))
            return
        parts = [t("tpl_pages", pages=p.expected_pages)]
        if p.expected_pages_across > 1:
            parts.insert(0, t("tpl_parts", parts=p.expected_pages_across))
        fixed = []
        if p.title_rows:
            fixed.append(t("tpl_rows_fixed", rows=p.title_rows))
        if p.title_cols:
            fixed.append(t("tpl_cols_fixed", cols=p.title_cols))
        elif p.print_headings:
            fixed.append(t("tpl_row_numbers"))
        src = ""
        if not tpl and (si.frozen_rows or si.frozen_cols):
            src = " " + t("tpl_from_frozen")
        self.summary.setText(" · ".join(parts + fixed) + src)

    # ---------------------------------------------------------------- edits
    def _store(self, key, value):
        if self.sheet is None:
            return
        self.templates.setdefault(self.sheet, {})[key] = value
        self.changed.emit()
        self.refresh()

    def _text_changed(self, key, edit: QLineEdit, rx):
        if self._loading or self.sheet is None:
            return
        txt = edit.text().strip().upper() if key == "title_cols" else edit.text().strip()
        if txt and not rx.match(txt):
            self.error.setText("⚠ " + self.w.tr_("fixed_rows" if key == "title_rows" else "fixed_cols")
                               + (": 1:2" if key == "title_rows" else ": A:B"))
            return
        si = self._sheet_info()
        suggested = suggested_titles(si, self.w._options())[0 if key == "title_rows" else 1] if si else None
        if (txt or None) == suggested and key not in self.templates.get(self.sheet, {}):
            return
        self._store(key, txt)

    def _order_changed(self, idx):
        if not self._loading:
            self._store("page_order", self.order.itemData(idx))

    def _rownum_changed(self, on):
        if not self._loading:
            self._store("row_numbers", bool(on))

    def _shape_changed(self, idx):
        if not self._loading and self.sheet is not None:
            self.shape_selected.emit(self.sheet, idx)

    def _reset(self):
        if self.sheet in self.templates:
            del self.templates[self.sheet]
            self.changed.emit()
        self.refresh()
