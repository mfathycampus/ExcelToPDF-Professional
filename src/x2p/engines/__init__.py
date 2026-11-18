from __future__ import annotations

from .base import Cancelled, Engine, EngineError
from .excel import ExcelEngine
from .libreoffice import LibreOfficeEngine

ENGINES = {"libreoffice": LibreOfficeEngine, "excel": ExcelEngine}


def get_engine(name: str = "auto") -> Engine:
    """'auto' prefers the free LibreOffice engine, then Excel if installed."""
    if name != "auto":
        if name not in ENGINES:
            raise EngineError("engine_unknown", engine=name)
        return ENGINES[name]()
    for cls in (LibreOfficeEngine, ExcelEngine):
        eng = cls()
        if eng.is_available():
            return eng
    raise EngineError("no_engine_available")


def engine_status() -> dict[str, bool]:
    return {name: cls().is_available() for name, cls in ENGINES.items()}


__all__ = ["Cancelled", "Engine", "EngineError", "get_engine", "engine_status", "ENGINES"]
