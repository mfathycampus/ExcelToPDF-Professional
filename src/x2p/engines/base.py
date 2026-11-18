"""Engine abstraction.

An engine receives an engine-neutral job (working copy path + per-sheet plans), exports
one PDF per sheet and returns the *displayed* text of the printed cells so verification
can compare them with the PDF independently of the engine.

Every engine runs its worker in a separate process: a crash, hang or modal dialog in
the office engine can never freeze the UI, and the process is killed on timeout/cancel.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import time
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any


class EngineError(Exception):
    def __init__(self, code: str, **detail):
        super().__init__(code)
        self.code = code
        self.detail = detail


class Cancelled(Exception):
    pass


class Engine(ABC):
    name: str = "base"
    display_name: str = "Base"

    @abstractmethod
    def is_available(self) -> bool: ...

    @abstractmethod
    def _command(self, job_path: Path, result_path: Path) -> list[str]: ...

    def prepare_job(self, job: dict[str, Any], workdir: Path) -> dict[str, Any]:
        return job

    def run(self, job: dict[str, Any], workdir: Path, *, timeout_s: int = 300,
            cancel: threading.Event | None = None) -> dict[str, Any]:
        if not self.is_available():
            raise EngineError("engine_not_available", engine=self.name)
        job = self.prepare_job(dict(job), workdir)
        job_path, result_path = workdir / "job.json", workdir / "result.json"
        job_path.write_text(json.dumps(job, ensure_ascii=False), encoding="utf-8")
        result_path.unlink(missing_ok=True)
        flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        log_path = workdir / f"{self.name}_worker.log"
        with log_path.open("wb") as log:
            proc = subprocess.Popen(self._command(job_path, result_path), stdout=log, stderr=log,
                                    stdin=subprocess.DEVNULL, creationflags=flags, cwd=str(workdir))
            start = time.monotonic()
            while proc.poll() is None:
                if cancel is not None and cancel.is_set():
                    kill_tree(proc.pid)
                    raise Cancelled()
                if time.monotonic() - start > timeout_s:
                    kill_tree(proc.pid)
                    raise EngineError("engine_timeout", seconds=timeout_s)
                time.sleep(0.15)
        if not result_path.exists():
            raise EngineError("engine_crashed", exit_code=proc.returncode, log=str(log_path))
        result = json.loads(result_path.read_text(encoding="utf-8"))
        if not result.get("ok"):
            raise EngineError(result.get("error") or "engine_failed", log=str(log_path))
        return result


def kill_tree(pid: int) -> None:
    if sys.platform == "win32":
        subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], capture_output=True,
                       creationflags=subprocess.CREATE_NO_WINDOW)
    else:  # pragma: no cover - non-Windows
        try:
            os.killpg(os.getpgid(pid), 9)
        except OSError:
            pass
