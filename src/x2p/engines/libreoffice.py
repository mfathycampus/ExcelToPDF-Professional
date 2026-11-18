"""LibreOffice Calc engine (default, free). The worker runs under LibreOffice's bundled
Python, because the `uno` module cannot be imported from a regular Python on Windows."""
from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path
from typing import Any

from .base import Engine

WORKER = Path(__file__).with_name("lo_worker.py")


def find_soffice() -> Path | None:
    env = os.environ.get("X2P_SOFFICE")
    if env and Path(env).is_file():
        return Path(env)
    candidates: list[Path] = []
    if sys.platform == "win32":
        import winreg
        for hive, key in ((winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\LibreOffice\UNO\InstallPath"),
                          (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\LibreOffice\UNO\InstallPath"),
                          (winreg.HKEY_CURRENT_USER, r"SOFTWARE\LibreOffice\UNO\InstallPath")):
            try:
                with winreg.OpenKey(hive, key) as k:
                    candidates.append(Path(winreg.QueryValueEx(k, "")[0]) / "soffice.exe")
            except OSError:
                pass
        for base in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                     os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
            candidates.append(Path(base) / "LibreOffice" / "program" / "soffice.exe")
    else:  # pragma: no cover
        candidates += [Path("/usr/bin/soffice"), Path("/usr/lib/libreoffice/program/soffice"),
                       Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")]
    which = shutil.which("soffice")
    if which:
        candidates.append(Path(which))
    for c in candidates:
        if c.is_file():
            return c
    return None


def find_lo_python(soffice: Path) -> Path | None:
    prog = soffice.parent
    for name in ("python.exe", "python"):
        if (prog / name).is_file():
            return prog / name
    if sys.platform != "win32":  # pragma: no cover - distro packages use system python3 + python3-uno
        return Path(shutil.which("python3") or "python3")
    return None


class LibreOfficeEngine(Engine):
    name = "libreoffice"
    display_name = "LibreOffice Calc"

    def __init__(self) -> None:
        self.soffice = find_soffice()
        self.python = find_lo_python(self.soffice) if self.soffice else None

    def is_available(self) -> bool:
        return bool(self.soffice and self.python and WORKER.is_file())

    def prepare_job(self, job: dict[str, Any], workdir: Path) -> dict[str, Any]:
        job["soffice"] = str(self.soffice)
        job["profile_dir"] = str(workdir / "lo_profile")   # isolated, per job
        return job

    def _command(self, job_path: Path, result_path: Path) -> list[str]:
        return [str(self.python), str(WORKER), str(job_path), str(result_path)]
