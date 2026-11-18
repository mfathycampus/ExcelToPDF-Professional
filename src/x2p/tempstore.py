"""Private temporary storage for jobs. Each job gets a random folder; stale folders from
crashed sessions are removed at start-up."""
from __future__ import annotations

import os
import shutil
import tempfile
import time
import uuid
from pathlib import Path

STALE_AFTER_S = 12 * 3600


def root() -> Path:
    # The per-user temp folder: private to the user and writable by every office engine
    # (some systems block Office from writing to arbitrary %LOCALAPPDATA% sub-folders).
    p = Path(os.environ.get("X2P_TEMP") or tempfile.gettempdir()) / "x2p-jobs"
    p.mkdir(parents=True, exist_ok=True)
    return p


def new_job_dir() -> Path:
    d = root() / uuid.uuid4().hex
    d.mkdir()  # no 0o700: on Windows that ACL blocks the Office engine from writing
    return d


def remove(path: Path | None) -> None:
    if path and path.exists():
        shutil.rmtree(path, ignore_errors=True)


def cleanup_stale(max_age_s: int = STALE_AFTER_S) -> int:
    n, now = 0, time.time()
    for d in root().iterdir():
        try:
            if d.is_dir() and now - d.stat().st_mtime > max_age_s:
                shutil.rmtree(d, ignore_errors=True)
                n += 1
        except OSError:
            continue
    return n
