"""Command line interface: x2p inspect | convert | engines"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .engines import engine_status
from .formats import UnsafeFileError
from .i18n import Translator
from .inspector import inspect_workbook
from .models import FinalStatus, JobOptions, LayoutMode
from .pipeline import Converter


def _print_json(obj) -> None:
    sys.stdout.write(json.dumps(obj, ensure_ascii=False, indent=2, default=str) + "\n")


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(prog="x2p", description="Excel to PDF Professional")
    ap.add_argument("--lang", default="en", choices=["ar", "en"])
    sub = ap.add_subparsers(dest="cmd", required=True)

    sub.add_parser("engines", help="show available conversion engines")

    p_in = sub.add_parser("inspect", help="inspect a workbook (read-only)")
    p_in.add_argument("file")

    p_cv = sub.add_parser("convert", help="convert a workbook to PDF")
    p_cv.add_argument("file")
    p_cv.add_argument("-o", "--output", help="output PDF (default: next to the source)")
    p_cv.add_argument("--sheets", nargs="*", help="sheet names (default: all visible)")
    p_cv.add_argument("--include-hidden", action="store_true")
    p_cv.add_argument("--mode", default="auto", choices=[m.value for m in LayoutMode])
    p_cv.add_argument("--paper", default="A4", choices=["A4", "A3", "Letter", "Legal"])
    p_cv.add_argument("--allow-a3", action="store_true")
    p_cv.add_argument("--min-font", type=float, default=8.0)
    p_cv.add_argument("--engine", default="auto", choices=["auto", "libreoffice", "excel"])
    p_cv.add_argument("--timeout", type=int, default=300)
    p_cv.add_argument("--overwrite", action="store_true")
    p_cv.add_argument("--allow-failed", action="store_true",
                      help="write the PDF even if verification failed (marked as not verified)")
    p_cv.add_argument("--report", help="write the JSON report to this path")
    p_cv.add_argument("--long-text", default="appendix", choices=["appendix", "wrap"],
                      help="very long cell text: move to an appendix (default) or wrap in place")

    a = ap.parse_args(argv)
    tr = Translator(a.lang)

    if a.cmd == "engines":
        _print_json(engine_status())
        return 0

    if a.cmd == "inspect":
        try:
            _print_json(inspect_workbook(a.file).to_dict())
        except UnsafeFileError as e:
            sys.stderr.write(tr(e.code, **e.detail) + "\n")
            return 2
        return 0

    src = Path(a.file)
    out = Path(a.output) if a.output else src.with_suffix(".pdf")
    opts = JobOptions(sheets=a.sheets or None, include_hidden=a.include_hidden, mode=LayoutMode(a.mode),
                      paper=a.paper, allow_larger_paper=a.allow_a3, min_font_pt=a.min_font,
                      engine=a.engine, timeout_s=a.timeout, allow_failed_output=a.allow_failed,
                      lang=a.lang, long_text=a.long_text)
    try:
        rep = Converter().convert(src, out, opts, overwrite=a.overwrite,
                                  progress=lambda s, p: sys.stderr.write(f"[{int(p * 100):3d}%] {tr('state_' + s)}\n"))
    except UnsafeFileError as e:
        sys.stderr.write(tr(e.code, **e.detail) + "\n")
        return 2
    except FileExistsError:
        sys.stderr.write(tr("overwrite_q") + " (--overwrite)\n")
        return 3
    d = rep.to_dict()
    if a.report:
        Path(a.report).write_text(json.dumps(d, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    sys.stderr.write(f"\n{tr(rep.status.value)} — {rep.total_pages} pages, engine={rep.engine_version}\n")
    for i in rep.issues:
        sys.stderr.write(f"  [{i.severity.value}] {i.sheet or '-'}: {tr.issue(i)}\n")
    if rep.delivered:
        sys.stderr.write(f"→ {rep.output_pdf}\n")
    return 0 if rep.status is not FinalStatus.FAILED else 1


if __name__ == "__main__":
    raise SystemExit(main())
