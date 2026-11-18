"""Optional Microsoft Excel engine (COM). Never required; used when the user picks it or
as the reference renderer in tests (spec §12.3)."""
from __future__ import annotations

import sys
from pathlib import Path

from .base import Engine

WORKER = Path(__file__).with_name("excel_worker.py")


def excel_installed() -> bool:
    if sys.platform != "win32":
        return False
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"):
            pass
    except OSError:
        return False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


class ExcelEngine(Engine):
    name = "excel"
    display_name = "Microsoft Excel"

    def is_available(self) -> bool:
        return excel_installed()

    def _command(self, job_path: Path, result_path: Path) -> list[str]:
        if getattr(sys, "frozen", False):   # packaged app: the executable itself hosts the worker
            return [sys.executable, "--x2p-worker", "excel", str(job_path), str(result_path)]
        return [sys.executable, str(WORKER), str(job_path), str(result_path)]
