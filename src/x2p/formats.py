"""Stage 1: safe file checks (signature, encryption, zip bombs, support policy)."""
from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path

ZIP_MAGIC = b"PK\x03\x04"
OLE_MAGIC = bytes.fromhex("D0CF11E0A1B11AE1")


class Support(str, Enum):
    SUPPORTED = "supported"
    EXPERIMENTAL = "experimental"   # only after explicit opt-in; not advertised
    UNSUPPORTED = "unsupported"


# A format is advertised only after passing its acceptance tests (spec §2.3).
SUPPORTED_FORMATS: dict[str, Support] = {
    ".xlsx": Support.SUPPORTED,
    ".xls": Support.SUPPORTED,
    ".xlsm": Support.EXPERIMENTAL,   # macros are never executed
    ".xlsb": Support.UNSUPPORTED,    # until acceptance tests pass
}

# Resource limits for OOXML containers.
MAX_ZIP_ENTRIES = 20_000
MAX_UNCOMPRESSED_BYTES = 2 * 1024**3
MAX_RATIO = 200          # per entry, only checked for entries larger than RATIO_MIN_BYTES
RATIO_MIN_BYTES = 20 * 1024**2


class UnsafeFileError(Exception):
    """Raised with a stable `code` the UI translates."""

    def __init__(self, code: str, **detail):
        super().__init__(code)
        self.code = code
        self.detail = detail


@dataclass
class FileCheck:
    path: Path
    extension: str
    container: str               # "ooxml" | "ole"
    size_bytes: int
    has_macros: bool = False
    external_links: int = 0
    notes: list[str] = field(default_factory=list)


def check_file(path: str | Path, *, max_mb: int = 200, allow_experimental: bool = True) -> FileCheck:
    p = Path(path)
    if not p.is_file():
        raise UnsafeFileError("file_not_found")
    ext = p.suffix.lower()
    support = SUPPORTED_FORMATS.get(ext)
    if support is None or support is Support.UNSUPPORTED:
        raise UnsafeFileError("format_unsupported", extension=ext)
    if support is Support.EXPERIMENTAL and not allow_experimental:
        raise UnsafeFileError("format_experimental", extension=ext)
    size = p.stat().st_size
    if size == 0:
        raise UnsafeFileError("file_empty")
    if size > max_mb * 1024**2:
        raise UnsafeFileError("file_too_large", size_mb=round(size / 1024**2, 1), limit_mb=max_mb)

    with p.open("rb") as fh:
        head = fh.read(8)

    if ext in (".xlsx", ".xlsm", ".xlsb"):
        if head.startswith(OLE_MAGIC):
            # Encrypted OOXML is wrapped in an OLE compound file.
            raise UnsafeFileError("password_protected")
        if not head.startswith(ZIP_MAGIC):
            raise UnsafeFileError("signature_mismatch", extension=ext)
        return _check_ooxml(p, ext, size)

    if ext == ".xls":
        if head.startswith(ZIP_MAGIC):
            raise UnsafeFileError("signature_mismatch", extension=ext, hint="looks_like_xlsx")
        if not head.startswith(OLE_MAGIC):
            raise UnsafeFileError("signature_mismatch", extension=ext)
        return FileCheck(p, ext, "ole", size)

    raise UnsafeFileError("format_unsupported", extension=ext)


def _check_ooxml(p: Path, ext: str, size: int) -> FileCheck:
    try:
        zf = zipfile.ZipFile(p)
    except zipfile.BadZipFile:
        raise UnsafeFileError("corrupt_file") from None
    with zf:
        infos = zf.infolist()
        if len(infos) > MAX_ZIP_ENTRIES:
            raise UnsafeFileError("zip_too_many_entries", entries=len(infos))
        total = 0
        names = set()
        for info in infos:
            total += info.file_size
            names.add(info.filename)
            if info.file_size > RATIO_MIN_BYTES and info.compress_size and \
                    info.file_size / info.compress_size > MAX_RATIO:
                raise UnsafeFileError("zip_bomb_suspected")
            if info.filename.startswith("/") or ".." in info.filename.split("/"):
                raise UnsafeFileError("zip_bad_path")
        if total > MAX_UNCOMPRESSED_BYTES:
            raise UnsafeFileError("zip_bomb_suspected")
        if "[Content_Types].xml" not in names:
            raise UnsafeFileError("corrupt_file")
        if ext == ".xlsb" or "xl/workbook.bin" in names:
            if ext != ".xlsb":
                raise UnsafeFileError("signature_mismatch", extension=ext, hint="looks_like_xlsb")
        elif "xl/workbook.xml" not in names:
            raise UnsafeFileError("corrupt_file")
        fc = FileCheck(p, ext, "ooxml", size)
        fc.has_macros = any(n.lower().endswith("vbaproject.bin") for n in names)
        fc.external_links = sum(1 for n in names if n.startswith("xl/externalLinks/externalLink"))
        return fc
