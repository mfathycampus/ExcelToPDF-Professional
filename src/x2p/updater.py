"""Self-update from the project's GitHub releases.

Flow: check the latest release -> download the win64 zip -> verify it against the release's
.sha256 file (mandatory: no checksum, no install) -> extract to a staging folder with path
checks -> after the app exits, a small cmd script copies the new files over the install folder
and starts the new version. Only the packaged Windows build updates itself; elsewhere the user
is sent to the release page.

Nothing about the user or their files is sent: one anonymous GET to api.github.com and the
downloads from github.com.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Callable

from . import __version__

REPO = "mfathycampus/ExcelToPDF-Professional"
LATEST_API = f"https://api.github.com/repos/{REPO}/releases/latest"
RELEASES_PAGE = f"https://github.com/{REPO}/releases/latest"
APP_DIR_NAME = "ExcelToPDF"
EXE_NAME = "ExcelToPDF.exe"
USER_AGENT = f"ExcelToPDF-Professional/{__version__}"
MAX_ZIP = 300 * 1024 * 1024


class UpdateError(Exception):
    """code is an i18n key."""

    def __init__(self, code: str):
        super().__init__(code)
        self.code = code


@dataclass
class UpdateInfo:
    version: str
    notes: str
    page_url: str
    zip_url: str
    zip_size: int
    sha_url: str


def parse_version(v: str) -> tuple[int, ...]:
    nums = re.findall(r"\d+", v.split("-")[0].split("+")[0])
    return tuple(int(n) for n in nums[:4]) or (0,)


def is_newer(remote: str, local: str = __version__) -> bool:
    return parse_version(remote) > parse_version(local)


def _api_url() -> str:
    return os.environ.get("X2P_UPDATE_URL", LATEST_API)   # override for tests only


def _get(url: str, timeout: float = 15):
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT,
                                               "Accept": "application/vnd.github+json"})
    return urllib.request.urlopen(req, timeout=timeout)  # noqa: S310 - https URLs we build


def check_latest(timeout: float = 15) -> UpdateInfo | None:
    """The latest release if it is newer than this build, else None. Raises UpdateError."""
    try:
        with _get(_api_url(), timeout) as r:
            data = json.loads(r.read(2_000_000).decode("utf-8"))
    except (OSError, ValueError) as e:
        raise UpdateError("update_check_failed") from e
    version = str(data.get("tag_name", "")).lstrip("vV")
    if not version or not is_newer(version):
        return None
    zip_a = sha_a = None
    for a in data.get("assets", []):
        name = a.get("name", "")
        if name.endswith("-win64.zip"):
            zip_a = a
        elif name.endswith("-win64.zip.sha256"):
            sha_a = a
    if not zip_a or not sha_a:
        raise UpdateError("update_no_package")
    return UpdateInfo(version=version, notes=str(data.get("body") or ""),
                      page_url=str(data.get("html_url") or RELEASES_PAGE),
                      zip_url=zip_a["browser_download_url"], zip_size=int(zip_a.get("size") or 0),
                      sha_url=sha_a["browser_download_url"])


def download(info: UpdateInfo, dest_dir: Path, progress: Callable[[float], None] | None = None,
             cancel: threading.Event | None = None) -> Path:
    """Download and verify the package. Returns the zip path."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    try:
        with _get(info.sha_url) as r:
            expected = r.read(4096).decode("ascii", "replace").split()[0].lower()
    except (OSError, IndexError) as e:
        raise UpdateError("update_download_failed") from e
    if not re.fullmatch(r"[0-9a-f]{64}", expected):
        raise UpdateError("update_bad_checksum")
    out = dest_dir / f"ExcelToPDF-{info.version}.zip"
    h, done = hashlib.sha256(), 0
    try:
        with _get(info.zip_url, timeout=60) as r, open(out, "wb") as f:
            total = int(r.headers.get("Content-Length") or info.zip_size or 0)
            while chunk := r.read(256 * 1024):
                if cancel is not None and cancel.is_set():
                    raise UpdateError("update_cancelled")
                done += len(chunk)
                if done > MAX_ZIP:
                    raise UpdateError("update_bad_checksum")
                f.write(chunk)
                h.update(chunk)
                if progress and total:
                    progress(min(done / total, 1.0))
    except OSError as e:
        raise UpdateError("update_download_failed") from e
    if h.hexdigest() != expected:
        out.unlink(missing_ok=True)
        raise UpdateError("update_bad_checksum")
    return out


def stage(zip_path: Path, staging: Path) -> Path:
    """Extract the verified zip safely. Returns the new app folder (…/ExcelToPDF)."""
    if staging.exists():
        shutil.rmtree(staging, ignore_errors=True)
    staging.mkdir(parents=True)
    root = staging.resolve()
    try:
        with zipfile.ZipFile(zip_path) as z:
            for m in z.infolist():
                p = PurePosixPath(m.filename.replace("\\", "/"))
                if p.is_absolute() or ".." in p.parts or (p.parts and ":" in p.parts[0]):
                    raise UpdateError("update_bad_package")
                if not p.parts or p.parts[0] != APP_DIR_NAME:
                    raise UpdateError("update_bad_package")
                target = (staging / Path(*p.parts)).resolve()
                if root not in target.parents:
                    raise UpdateError("update_bad_package")
            z.extractall(staging)
    except zipfile.BadZipFile as e:
        raise UpdateError("update_bad_package") from e
    app = staging / APP_DIR_NAME
    if not (app / EXE_NAME).is_file() or not (app / "_internal").is_dir():
        raise UpdateError("update_bad_package")
    return app


def install_dir() -> Path | None:
    """Folder of the running packaged app, or None when not running from the package."""
    if not getattr(sys, "frozen", False) or sys.platform != "win32":
        return None
    d = Path(sys.executable).resolve().parent
    if (d / EXE_NAME).is_file() and (d / "_internal").is_dir():
        return d
    return None


def can_self_update() -> bool:
    d = install_dir()
    if d is None:
        return False
    try:
        probe = d / ".x2p-write-test"
        probe.write_bytes(b"")
        probe.unlink()
        return True
    except OSError:
        return False


# The replacer waits for the app to exit, mirrors _internal (removing stale libraries), copies the
# top-level files, then starts the new version. Paths arrive as arguments; the script is ASCII.
_REPLACER = r"""@echo off
setlocal
set "PID=%~1"
set "NEW=%~2"
set "DEST=%~3"
set "LOG=%~4"
set /a N=0
:wait
tasklist /FI "PID eq %PID%" /NH 2>nul | find "%PID%" >nul
if not errorlevel 1 (
  set /a N+=1
  if %N% GEQ 120 goto failed
  ping -n 2 127.0.0.1 >nul
  goto wait
)
robocopy "%NEW%\_internal" "%DEST%\_internal" /MIR /R:10 /W:1 /NFL /NDL /NJH /NJS /NP >>"%LOG%"
if errorlevel 8 goto failed
robocopy "%NEW%" "%DEST%" /E /XD "%NEW%\_internal" /R:10 /W:1 /NFL /NDL /NJH /NJS /NP >>"%LOG%"
if errorlevel 8 goto failed
echo ok>>"%LOG%"
start "" "%DEST%\ExcelToPDF.exe"
exit /b 0
:failed
echo failed>>"%LOG%"
start "" "%DEST%\ExcelToPDF.exe"
exit /b 1
"""


def launch_replacer(new_app: Path, dest: Path, pid: int | None = None) -> Path:
    """Start the detached replacer. The caller must exit right after. Returns the log path."""
    work = Path(tempfile.gettempdir()) / "x2p-update"
    work.mkdir(parents=True, exist_ok=True)
    script = work / "apply_update.cmd"
    script.write_text(_REPLACER, encoding="ascii", newline="\r\n")
    log = work / "update.log"
    log.write_text("", encoding="ascii")
    flags = 0
    if sys.platform == "win32":
        flags = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.CREATE_NO_WINDOW
    subprocess.Popen(["cmd", "/c", str(script), str(pid or os.getpid()), str(new_app), str(dest), str(log)],
                     creationflags=flags, close_fds=True, cwd=str(work))
    return log


def staging_dir() -> Path:
    return Path(tempfile.gettempdir()) / "x2p-update" / "staging"
