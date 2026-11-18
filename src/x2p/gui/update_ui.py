"""Update button + daily background check. Network and file work run off the UI thread."""
from __future__ import annotations

import os
import threading
import time

from PySide6.QtCore import QObject, Qt, QTimer, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import QApplication, QMessageBox, QProgressDialog

from .. import __version__, updater

DAY = 24 * 3600


class _Bridge(QObject):
    checked = Signal(object, object)      # UpdateInfo | None, error code | None
    progress = Signal(float)
    staged = Signal(object, object)       # new app dir | None, error code | None


class UpdateController(QObject):
    def __init__(self, window):
        super().__init__(window)
        self.w = window
        self.b = _Bridge(self)
        self.b.checked.connect(self._on_checked)
        self.b.progress.connect(self._on_progress)
        self.b.staged.connect(self._on_staged)
        self._checking = False
        self._manual = False
        self._dlg: QProgressDialog | None = None
        self._cancel = threading.Event()
        self.info: updater.UpdateInfo | None = None

    # ---------------------------------------------------------------- check
    def schedule_startup_check(self):
        if os.environ.get("X2P_NO_UPDATE_CHECK"):
            return
        last = float(self.w.settings.value("update_last_check", 0) or 0)
        if time.time() - last >= DAY:
            QTimer.singleShot(3000, lambda: self.check(manual=False))

    def check(self, manual: bool = True):
        if self._checking:
            return
        self._checking, self._manual = True, manual
        if manual:
            self.w.statusBar().showMessage(self.w.tr_("update_checking"))
            QApplication.setOverrideCursor(Qt.BusyCursor)

        def run():
            try:
                self.b.checked.emit(updater.check_latest(), None)
            except updater.UpdateError as e:
                self.b.checked.emit(None, e.code)
        threading.Thread(target=run, daemon=True).start()

    def _on_checked(self, info, err):
        t = self.w.tr_
        self._checking = False
        if self._manual:
            QApplication.restoreOverrideCursor()
            self.w.statusBar().clearMessage()
        if err is None:
            self.w.settings.setValue("update_last_check", time.time())
        if err:
            if self._manual:
                QMessageBox.warning(self.w, t("app_title"), t(err))
            return
        if info is None:
            if self._manual:
                QMessageBox.information(self.w, t("app_title"), t("update_none", version=__version__))
            return
        self.info = info
        self.w.act_update.setText(t("update_available_short", version=info.version))
        if not self._manual and self.w.settings.value("update_skipped") == info.version:
            return
        self._offer(info)

    def _offer(self, info):
        t = self.w.tr_
        box = QMessageBox(self.w)
        box.setIcon(QMessageBox.Information)
        box.setWindowTitle(t("app_title"))
        box.setText(t("update_available", version=info.version, current=__version__))
        notes = info.notes.strip()
        if notes:
            box.setDetailedText(notes[:4000])
        self_update = updater.can_self_update()
        b_now = box.addButton(t("update_now") if self_update else t("update_open_page"), QMessageBox.AcceptRole)
        box.addButton(t("later"), QMessageBox.RejectRole)
        box.exec()
        if box.clickedButton() is not b_now:
            self.w.settings.setValue("update_skipped", info.version)
            return
        if not self_update:
            QDesktopServices.openUrl(QUrl(info.page_url))
            return
        if self.w.runner.busy:
            QMessageBox.information(self.w, t("app_title"), t("update_wait_job"))
            return
        self._download(info)

    # ---------------------------------------------------------------- download + install
    def _download(self, info):
        t = self.w.tr_
        self._cancel = threading.Event()
        self._dlg = QProgressDialog(t("update_downloading", version=info.version), t("cancel"), 0, 100, self.w)
        self._dlg.setWindowTitle(t("app_title"))
        self._dlg.setWindowModality(Qt.WindowModal)
        self._dlg.setMinimumDuration(0)
        self._dlg.canceled.connect(self._cancel.set)
        self._dlg.show()

        def run():
            try:
                work = updater.staging_dir().parent
                z = updater.download(info, work, progress=self.b.progress.emit, cancel=self._cancel)
                app = updater.stage(z, updater.staging_dir())
                z.unlink(missing_ok=True)
                self.b.staged.emit(app, None)
            except updater.UpdateError as e:
                self.b.staged.emit(None, e.code)
            except OSError:
                self.b.staged.emit(None, "update_download_failed")
        threading.Thread(target=run, daemon=True).start()

    def _on_progress(self, frac: float):
        if self._dlg is not None:
            self._dlg.setValue(int(frac * 100))

    def _on_staged(self, app, err):
        t = self.w.tr_
        if self._dlg is not None:
            self._dlg.canceled.disconnect()
            self._dlg.close()
            self._dlg = None
        if err:
            if err != "update_cancelled":
                QMessageBox.warning(self.w, t("app_title"), t(err))
            return
        if QMessageBox.question(self.w, t("app_title"), t("update_ready_restart")) != QMessageBox.Yes:
            return
        dest = updater.install_dir()
        if dest is None:
            return
        try:
            updater.launch_replacer(app, dest)
        except OSError:
            QMessageBox.warning(self.w, t("app_title"), t("update_install_failed"))
            return
        self.w.close()
        QApplication.quit()
