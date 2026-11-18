"""Runs core jobs off the UI thread. The UI never talks to an engine directly."""
from __future__ import annotations

import threading
import traceback

from PySide6.QtCore import QObject, QThread, Signal

from ..formats import UnsafeFileError
from ..inspector import inspect_workbook
from ..models import JobOptions
from ..pipeline import Converter


class _Job(QObject):
    progress = Signal(str, float)
    finished = Signal(object)          # result object
    failed = Signal(str, dict)         # code, detail

    def __init__(self, fn):
        super().__init__()
        self._fn = fn

    def run(self):
        try:
            self.finished.emit(self._fn(self))
        except UnsafeFileError as e:
            self.failed.emit(e.code, e.detail)
        except Exception:  # noqa: BLE001 - surface any crash as a translated error
            traceback.print_exc()
            self.failed.emit("engine_failed", {})


class JobRunner(QObject):
    """One job at a time per window; each job gets its own thread and cancel flag."""
    progress = Signal(str, float)
    finished = Signal(str, object)     # kind, result
    failed = Signal(str, str, dict)    # kind, code, detail

    def __init__(self, parent=None):
        super().__init__(parent)
        self._thread: QThread | None = None
        self._job: _Job | None = None
        self.cancel_event = threading.Event()

    @property
    def busy(self) -> bool:
        return self._thread is not None and self._thread.isRunning()

    def _start(self, kind: str, fn):
        if self.busy:
            return
        self.cancel_event = threading.Event()
        self._thread = QThread()
        self._job = _Job(fn)
        self._job.moveToThread(self._thread)
        self._thread.started.connect(self._job.run)
        self._job.progress.connect(self.progress)
        self._job.finished.connect(lambda r: self.finished.emit(kind, r))
        self._job.failed.connect(lambda c, d: self.failed.emit(kind, c, d))
        self._job.finished.connect(self._thread.quit)
        self._job.failed.connect(self._thread.quit)
        self._thread.start()

    def inspect(self, path: str):
        self._start("inspect", lambda job: inspect_workbook(path))

    def convert(self, path: str, opts: JobOptions):
        def fn(job):  # reads self.cancel_event at run time: _start() has replaced it by then
            return Converter().convert(path, None, opts, keep_temp=True, cancel=self.cancel_event,
                                       progress=lambda s, p: job.progress.emit(s, p))
        self._start("convert", fn)

    def cancel(self):
        self.cancel_event.set()

    def wait(self, ms: int = 5000):
        if self._thread is not None:
            self._thread.wait(ms)
