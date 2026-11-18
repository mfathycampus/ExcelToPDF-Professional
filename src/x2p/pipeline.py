"""Orchestrates stages 1-7. No Qt here: the CLI, the GUI and the tests all use this."""
from __future__ import annotations

import hashlib
import shutil
import threading
import time
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

import pymupdf

from . import tempstore
from .appendix import APPENDIX_COL_WIDTHS, build_appendix
from .engines import Cancelled, Engine, EngineError, get_engine
from .inspector import inspect_workbook
from .layout import alternative_plan, plan_sheet
from .models import (ConversionReport, FinalStatus, Issue, JobOptions, LayoutMode, Severity,
                     SheetInfo, SheetPlan, SheetVerification, WorkbookInfo)
from .validation import chars_below, verify_sheet_pdf

Progress = Callable[[str, float], None]


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def select_sheets(info: WorkbookInfo, opts: JobOptions) -> tuple[list[SheetInfo], dict[str, str]]:
    chosen, excluded = [], {}
    wanted = set(opts.sheets) if opts.sheets is not None else None
    for s in info.sheets:
        if wanted is not None and s.name not in wanted:
            excluded[s.name] = "not_selected"
        elif s.state != "visible" and not opts.include_hidden and wanted is None:
            excluded[s.name] = "hidden"
        elif s.is_empty:
            excluded[s.name] = "empty"
        else:
            chosen.append(s)
    return chosen, excluded


@dataclass
class _SheetOutcome:
    info: SheetInfo
    plan: SheetPlan
    pdf: str | None = None
    verification: SheetVerification | None = None
    issues: list[Issue] = field(default_factory=list)
    retry_reason: str | None = None
    create: dict | None = None        # engine instruction to create this sheet (appendix)


class Converter:
    def __init__(self, engine: Engine | None = None):
        self._engine = engine

    def convert(self, source: str | Path, output_pdf: str | Path | None, opts: JobOptions | None = None,
                *, progress: Progress | None = None, cancel: threading.Event | None = None,
                overwrite: bool = False, keep_temp: bool = False) -> ConversionReport:
        opts = opts or JobOptions()
        progress = progress or (lambda stage, pct: None)
        src = Path(source)
        report = ConversionReport(workbook=src.name)
        t0 = time.monotonic()
        job_dir: Path | None = None
        try:
            progress("inspecting", 0.05)
            before = sha256(src)
            info = inspect_workbook(src, max_mb=opts.max_file_mb)
            report.issues.extend(info.issues)
            chosen, report.excluded_sheets = select_sheets(info, opts)
            if not chosen:
                report.issues.append(Issue("no_sheets_to_export", Severity.ERROR))
                return self._finish(report, t0, src, before)

            engine = self._engine or get_engine(opts.engine)
            report.engine = engine.name
            outcomes = [_SheetOutcome(s, plan_sheet(s, opts)) for s in chosen]
            app = build_appendix([(o.info, o.plan) for o in outcomes], [s.name for s in info.sheets],
                                 opts.lang, opts.paper)
            if app is not None:
                a_plan, a_rows, a_bold, a_rtl = app
                a_info = SheetInfo(index=len(info.sheets), name=a_plan.sheet, right_to_left=a_rtl)
                outcomes.append(_SheetOutcome(a_info, a_plan, create={
                    "rows": a_rows, "bold_rows": a_bold, "rtl": a_rtl, "widths": APPENDIX_COL_WIDTHS}))

            job_dir = tempstore.new_job_dir()
            work = job_dir / ("workbook" + src.suffix.lower())
            shutil.copy2(src, work)

            progress("exporting", 0.2)
            self._run_engine(engine, work, job_dir / "pass1", outcomes, opts, cancel, report)
            progress("verifying", 0.6)
            for o in outcomes:
                self._verify(o, opts)

            retry = []
            for o in outcomes:
                if o.retry_reason:
                    measured = o.verification.min_font_pt if o.verification else None
                    alt = alternative_plan(o.info, o.plan, opts, o.retry_reason, measured)
                    if alt is not None:
                        o.plan = alt
                        retry.append(o)
            if retry:
                progress("retrying", 0.7)
                self._run_engine(engine, work, job_dir / "pass2", retry, opts, cancel, report)
                for o in retry:
                    o.issues = [Issue("fallback_plan_used", Severity.INFO, o.info.name,
                                      {"reason": o.retry_reason})]
                    o.retry_reason = None
                    self._verify(o, opts, attempt=2)

            progress("merging", 0.85)
            merged = job_dir / "output.pdf"
            self._merge(outcomes, merged, src.stem, report)
            for o in outcomes:
                report.issues.extend(o.issues)
                if o.verification:
                    report.sheets.append(o.verification)
            report.temp_pdf = str(merged) if merged.exists() else None
            report.pdf_size_bytes = merged.stat().st_size if merged.exists() else 0

            report.status = self._status(report)
            if output_pdf and report.temp_pdf:
                if report.status is FinalStatus.FAILED and not opts.allow_failed_output:
                    report.issues.append(Issue("delivery_blocked_verification_failed", Severity.ERROR))
                else:
                    deliver(Path(report.temp_pdf), Path(output_pdf), overwrite=overwrite)
                    report.output_pdf = str(Path(output_pdf))
                    report.delivered = True
            progress("done", 1.0)
            return self._finish(report, t0, src, before)
        except Cancelled:
            report.issues.append(Issue("cancelled", Severity.ERROR))
            keep_temp = False
            return self._finish(report, t0, src, None)
        except EngineError as exc:
            report.issues.append(Issue(exc.code, Severity.ERROR, detail={k: v for k, v in exc.detail.items()
                                                                          if k != "log"}))
            return self._finish(report, t0, src, None)
        finally:
            if job_dir is not None and not keep_temp:
                tempstore.remove(job_dir)
                report.temp_pdf = None

    # ------------------------------------------------------------------ steps
    def _run_engine(self, engine: Engine, work: Path, out_dir: Path, outcomes: list[_SheetOutcome],
                    opts: JobOptions, cancel, report: ConversionReport) -> None:
        out_dir.mkdir(parents=True, exist_ok=True)
        job = {
            "input": str(work), "out_dir": str(out_dir), "collect_text": True,
            "max_text_cells": opts.max_text_cells,
            "sheets": [{"name": o.info.name, "plan": o.plan.to_dict(),
                        **({"create_appendix": o.create} if o.create else {})} for o in outcomes],
        }
        # big sheets take longer (wrapping, row heights, text collection): scale the time limit
        cells = sum(o.info.value_cells + o.info.formula_cells for o in outcomes)
        timeout = int(opts.timeout_s + cells / 50)
        result = engine.run(job, out_dir, timeout_s=timeout, cancel=cancel)
        report.engine_version = result.get("engine_version", "")
        by_name = {r["name"]: r for r in result.get("sheets", [])}
        for o in outcomes:
            r = by_name.get(o.info.name, {"error": "sheet_not_processed"})
            o.pdf = r.get("pdf")
            o._engine_sheet = r  # type: ignore[attr-defined]

    def _verify(self, o: _SheetOutcome, opts: JobOptions, attempt: int = 1) -> None:
        r = getattr(o, "_engine_sheet", {})
        name = o.info.name
        if not o.pdf:
            o.issues.append(Issue(r.get("error") or "sheet_export_failed", Severity.ERROR, name))
            o.verification = SheetVerification(sheet=name, plan=o.plan.to_dict(), attempt=attempt)
            return
        v, sizes = verify_sheet_pdf(o.pdf, r, o.info, o.plan, opts.min_font_pt)
        v.attempt = attempt
        o.verification = v
        iss = o.issues
        if v.pages == 0:
            iss.append(Issue("no_pages_exported", Severity.ERROR, name))
        if v.missing_cells:
            iss.append(Issue("text_missing", Severity.ERROR, name,
                             {"missing": v.expected_texts - v.found_texts - len(v.rtl_unverified),
                              "cells": v.missing_cells[:20]}))
        if v.rtl_unverified:
            iss.append(Issue("rtl_order_unverified", Severity.WARNING, name,
                             {"count": len(v.rtl_unverified), "cells": v.rtl_unverified[:20]}))
        if v.hash_cells:
            iss.append(Issue("cell_overflow_hash", Severity.ERROR, name, {"count": v.hash_cells}))
        if r.get("paper_mismatch"):
            w, h = r["paper_mismatch"]
            iss.append(Issue("paper_changed_by_printer", Severity.WARNING, name,
                             {"paper": o.plan.paper, "size": f"{round(w / 72 * 25.4)}x{round(h / 72 * 25.4)} mm"}))
        threshold = min(opts.min_font_pt, o.info.min_font_pt_source or opts.min_font_pt)
        below, total = chars_below(sizes, threshold)
        if below:
            ratio = below / max(total, 1)
            if o.plan.mode == LayoutMode.ONE_PAGE or o.plan.mode == LayoutMode.MANUAL:
                iss.append(Issue("font_below_min_user_choice", Severity.WARNING, name,
                                 {"min_font_pt": v.min_font_pt, "threshold": threshold}))
            elif ratio > 0.01:
                iss.append(Issue("font_below_min", Severity.ERROR, name,
                                 {"min_font_pt": v.min_font_pt, "threshold": threshold}))
                if attempt == 1:
                    o.retry_reason = "font_below_min"
            else:
                iss.append(Issue("small_text_present", Severity.WARNING, name,
                                 {"min_font_pt": v.min_font_pt, "ratio": round(ratio, 4)}))
        if v.blank_pages:
            iss.append(Issue("blank_pages", Severity.WARNING, name, {"pages": v.blank_pages}))
        if v.out_of_page_spans:
            iss.append(Issue("content_outside_page", Severity.WARNING, name, {"count": v.out_of_page_spans}))
        cards = o.plan.mode == LayoutMode.CARDS   # visuals are reported by the plan note instead
        if o.info.images and v.pdf_images == 0 and not cards:
            iss.append(Issue("images_not_found_in_pdf", Severity.WARNING, name, {"expected": o.info.images}))
        if o.info.charts and v.pdf_drawings == 0 and v.pdf_images == 0 and not cards:
            iss.append(Issue("charts_not_found_in_pdf", Severity.WARNING, name, {"expected": o.info.charts}))
        if o.info.shapes and not cards:
            iss.append(Issue("shapes_not_verified", Severity.WARNING, name, {"count": o.info.shapes}))
        if v.sampled:
            iss.append(Issue("coverage_sampled", Severity.INFO, name, {"checked": v.expected_texts}))
        if o.info.kind == "chartsheet":
            iss.append(Issue("chartsheet_visual_only", Severity.WARNING, name))
        for note in o.plan.notes:
            sev = Severity.WARNING if note.endswith("below_minimum") else Severity.INFO
            detail = {}
            if note == "cards_suggested":
                detail = {"parts": o.plan.expected_pages_across}
            elif note == "overflow_text_wrapped":
                detail = {"count": o.plan.wrapped_cells}
            elif note == "columns_widened":
                detail = {"count": len(o.plan.col_widths)}
            elif note == "long_text_to_appendix":
                detail = {"count": len(o.plan.long_cells)}
            iss.append(Issue(note, sev, name, detail))

    def _merge(self, outcomes: list[_SheetOutcome], merged: Path, title: str,
               report: ConversionReport) -> None:
        out = pymupdf.open()
        toc = []
        for o in outcomes:
            if not o.pdf:
                continue
            with pymupdf.open(o.pdf) as part:
                if part.page_count == 0:
                    continue
                start = out.page_count
                out.insert_pdf(part)
                report.sheet_page_ranges[o.info.name] = (start, out.page_count - 1)
                report.exported_sheets.append(o.info.name)
                toc.append([1, o.info.name, start + 1])
        if out.page_count == 0:
            out.close()
            return
        out.set_toc(toc)
        out.set_metadata({"title": title, "creator": "Excel to PDF Professional",
                          "producer": f"x2p + {report.engine_version or report.engine}"})
        out.save(merged, garbage=3, deflate=True)
        report.total_pages = out.page_count
        out.close()

    @staticmethod
    def _status(report: ConversionReport) -> FinalStatus:
        sev = Counter(i.severity for i in report.issues)
        if sev[Severity.ERROR] or not report.exported_sheets:
            return FinalStatus.FAILED
        if sev[Severity.WARNING]:
            return FinalStatus.PASSED_WITH_WARNINGS
        return FinalStatus.PASSED

    @staticmethod
    def _finish(report: ConversionReport, t0: float, src: Path, before: str | None) -> ConversionReport:
        if before is not None and src.exists():
            report.source_unchanged = sha256(src) == before
            if not report.source_unchanged:
                report.issues.append(Issue("source_modified", Severity.ERROR))
                report.status = FinalStatus.FAILED
        if any(i.severity is Severity.ERROR for i in report.issues):
            report.status = FinalStatus.FAILED
        report.elapsed_s = round(time.monotonic() - t0, 2)
        return report


def deliver(temp_pdf: Path, dest: Path, *, overwrite: bool = False) -> None:
    """Atomic-ish move into place; never overwrites without explicit consent."""
    if dest.exists() and not overwrite:
        raise FileExistsError(str(dest))
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_name(dest.name + ".x2p-partial")
    shutil.copy2(temp_pdf, tmp)
    tmp.replace(dest)
