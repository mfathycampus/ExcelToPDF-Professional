"""Stages 3-4: engine-neutral page planning.

The planner works on estimates (column widths/row heights from the file). The engine
does the real layout, and verification then *measures* the PDF; when the measurement
violates a hard rule the pipeline asks `alternative_plan` for exactly one retry.
"""
from __future__ import annotations

import math
from dataclasses import replace

from .models import JobOptions, LayoutMode, SheetInfo, SheetPlan

MM = 72 / 25.4
PAPER_PT = {
    "A4": (595.28, 841.89),
    "A3": (841.89, 1190.55),
    "Letter": (612.0, 792.0),
    "Legal": (612.0, 1008.0),
}
DEFAULT_MARGINS_MM = (19.0, 19.0, 17.8, 17.8)  # Excel "Normal": top, bottom, left, right


def col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def printable_area(paper: str, orientation: str, margins_mm=DEFAULT_MARGINS_MM) -> tuple[float, float]:
    w, h = PAPER_PT[paper]
    if orientation == "landscape":
        w, h = h, w
    top, bottom, left, right = margins_mm
    return w - (left + right) * MM, h - (top + bottom) * MM


def min_scale_for(si: SheetInfo, min_font_pt: float) -> float:
    """Smallest allowed scale so that the smallest source text stays >= min_font_pt.
    If the source itself is already below the minimum we never shrink (scale 1.0)."""
    smallest = si.min_font_pt_source or si.base_font_pt or 11.0
    if smallest <= min_font_pt:
        return 1.0
    return min(1.0, min_font_pt / smallest)


def _title_height(si: SheetInfo, rows: str | None) -> float:
    rect = si.layout_rect or si.content
    if not rows or not rect:
        return 0.0
    try:
        a, b = (int(x) for x in rows.split(":"))
    except ValueError:
        return 0.0
    i0, i1 = a - rect.min_row, b - rect.min_row
    return sum(si.row_heights_pt[max(0, i0): max(0, i1 + 1)])


def _evaluate(si: SheetInfo, paper: str, orient: str, min_scale: float,
              title_rows: str | None, key_cols: int) -> dict:
    pw, ph = printable_area(paper, orient)
    cw, ch = max(si.content_width_pt, 1.0), max(si.content_height_pt, 1.0)
    key_w = sum(si.col_widths_pt[:key_cols]) if key_cols else 0.0
    title_h = _title_height(si, title_rows)
    s_fit = min(1.0, pw / cw)
    if s_fit >= min_scale:
        scale, across = s_fit, 1
    else:
        # fewest horizontal parts at the smallest legal scale, then grow scale while
        # the number of parts stays the same.
        scale = min_scale
        usable = max(pw - key_w * scale, pw * 0.3)
        across = max(2, math.ceil((cw * scale - key_w * scale) / usable))
        grow = across * pw / (cw + (across - 1) * key_w)
        scale = max(min_scale, min(1.0, grow * 0.98))
    usable_h = max(ph - title_h * scale, ph * 0.3)
    down = max(1, math.ceil((ch * scale - title_h * scale) / usable_h)) if ch * scale > ph else 1
    return {"paper": paper, "orientation": orient, "scale": scale, "across": across,
            "down": down, "pages": across * down}


def plan_sheet(si: SheetInfo, opts: JobOptions) -> SheetPlan:
    plan = _plan_sheet(si, opts)
    if plan.mode == LayoutMode.CARDS:
        return plan
    if opts.wrap_overflow and si.overflow and not plan.keep_original:
        add_overflow_wrapping(si, plan, opts)
    if plan.expected_pages_across >= CARDS_SUGGEST_ACROSS and plan.mode == LayoutMode.AUTO:
        plan.notes.append("cards_suggested")
    return plan


CARDS_SUGGEST_ACROSS = 4          # suggest card mode when a sheet splits into this many parts
CARD_WIDTHS = [30, 60]            # field / value columns (Excel width units): A4 portrait at 100 %
CARD_CHUNK = 900                  # characters per value row (row-height limit)


def plan_cards(si: SheetInfo, opts: JobOptions) -> SheetPlan:
    """Card per record: every visible data row becomes a block of "field | value" rows in a new
    sheet of the working copy (values exactly as the engine displays them). Chosen explicitly by
    the user (spec §7.4 forbids rearranging data automatically)."""
    from .i18n import Translator
    t = Translator(opts.lang)
    rect = si.layout_rect or si.content
    header = si.header_row_guess if si.header_row_guess and si.header_confidence >= 0.6 else None
    first = (header + 1) if header else rect.min_row
    name = f"{si.name[:20]} - {t('cards_suffix')}"[:31]
    plan = SheetPlan(sheet=si.name, mode=LayoutMode.CARDS, paper=opts.paper, orientation="portrait",
                     scale=None, fit_width=1, fit_height=0)
    plan.cards = {
        "sheet_name": name, "header_row": header, "first_row": first, "last_row": rect.max_row,
        "first_col": rect.min_col, "last_col": rect.max_col,
        "key_col": si.key_col_guess if si.key_confidence >= 0.6 else None,
        "row_label": t("row"), "widths": CARD_WIDTHS, "chunk": CARD_CHUNK, "rtl": si.right_to_left,
    }
    # estimate: one line per non-empty field (+ wrapped lines), a title and a spacer per record
    records = max(1, len(si.row_heights_pt) - (1 if header else 0))
    cells = si.value_cells + si.formula_cells
    lines = cells + records * 2 + sum(len(v) // CARD_WIDTHS[1] for v in si.long_texts.values())
    _pw, ph = printable_area(opts.paper, "portrait")
    plan.expected_pages = max(1, math.ceil(lines * 11.5 / ph))
    plan.estimated_min_font_pt = 10.0
    plan.notes.append("cards_layout")
    if si.images or si.charts or si.shapes:
        plan.notes.append("cards_visuals_not_included")
    return plan


def visible_row_blocks(si: SheetInfo, lo: int, hi: int) -> list[str]:
    hidden = set(si.hidden_row_numbers)
    blocks, start = [], None
    for r in range(lo, hi + 2):
        if r <= hi and r not in hidden:
            start = r if start is None else start
        elif start is not None:
            blocks.append(f"{start}:{r - 1}")
            start = None
    return blocks


def add_overflow_wrapping(si: SheetInfo, plan: SheetPlan, opts: JobOptions) -> None:
    """Text that would be clipped (by its neighbour or a page edge) is wrapped in its cell and
    the row height recomputed. Text too long to wrap legibly (> opts.long_text_lines lines)
    goes to the long-text appendix instead and the cell shows a reference marker.
    Applied to the working copy only; both are reported."""
    to_appendix = opts.long_text == "appendix"
    long_cells, cols = [], {}
    for r, c, lines in si.overflow:
        if to_appendix and lines > opts.long_text_lines and (r, c) in si.long_texts:
            long_cells.append((r, c))
            continue
        lo, hi = cols.get(c, (r, r))
        cols[c] = (min(lo, r), max(hi, r))
    plan.long_cells = sorted(long_cells)
    if to_appendix and long_cells:
        # markers are short: no wrapping needed for those cells, rows can be autofit
        cols_long = {c for _, c in long_cells}
        for c in cols_long:
            rr = [r for r, cc in long_cells if cc == c]
            lo, hi = cols.get(c, (min(rr), max(rr)))
            cols.setdefault(c, (lo, hi))
    if cols:
        plan.wrap_ranges = [f"{col_letter(c)}{lo}:{col_letter(c)}{hi}" for c, (lo, hi) in sorted(cols.items())]
        lo = min(v[0] for v in cols.values())
        hi = max(v[1] for v in cols.values())
        plan.autofit_rows = visible_row_blocks(si, lo, hi)
        plan.notes.append("overflow_text_wrapped")
        _reestimate_after_wrapping(si, plan, set(long_cells), opts)
    if long_cells:
        plan.notes.append("long_text_to_appendix")


MAX_ROW_PT = 409.0   # Excel's row-height limit


def _reestimate_after_wrapping(si: SheetInfo, plan: SheetPlan, long_cells: set, opts: JobOptions) -> None:
    """Wrapped rows get taller: redo the vertical page count with estimated row heights
    (each wrapped line ~1.35 x font size; capped at the engine row limit)."""
    rect = si.layout_rect or si.content
    if not rect:
        return
    cap = opts.long_text_lines if opts.long_text == "appendix" else 10_000
    row_lines: dict[int, float] = {}
    for r, c, lines in si.overflow:
        if (r, c) not in long_cells:
            row_lines[r] = max(row_lines.get(r, 1.0), min(float(math.ceil(lines)), cap))
    hidden = set(si.hidden_row_numbers)
    visible = [r for r in range(rect.min_row, rect.max_row + 1) if r not in hidden]
    line_h = (si.base_font_pt or 11.0) * 1.35
    total = sum(min(MAX_ROW_PT, max(h, row_lines.get(r, 1.0) * line_h + 2))
                for r, h in zip(visible, si.row_heights_pt))
    smallest = si.min_font_pt_source or si.base_font_pt or 11.0
    if plan.scale:
        scale = plan.scale / 100
    elif plan.estimated_min_font_pt:
        scale = min(1.0, plan.estimated_min_font_pt / smallest)
    else:
        scale = 1.0
    _pw, ph = printable_area(plan.paper, plan.orientation)
    usable = max(ph - _title_height(si, plan.title_rows) * scale, ph * 0.3)
    down = max(1, math.ceil(total * scale / usable))
    plan.expected_pages = max(1, plan.expected_pages_across) * down


def _plan_sheet(si: SheetInfo, opts: JobOptions) -> SheetPlan:
    mode = opts.per_sheet_mode.get(si.name, opts.mode)
    plan = SheetPlan(sheet=si.name, mode=mode, paper=opts.paper)
    ps = si.page_setup
    min_scale = min_scale_for(si, opts.min_font_pt)
    smallest = si.min_font_pt_source or si.base_font_pt

    if si.kind == "chartsheet" or not si.col_widths_pt:
        plan.keep_original = True
        plan.notes.append("no_cell_grid_keep_original")
        return plan

    if mode == LayoutMode.KEEP_ORIGINAL:
        plan.keep_original = True
        return plan

    if mode == LayoutMode.CARDS:
        return plan_cards(si, opts)

    papers = [opts.paper] + (["A3"] if opts.allow_larger_paper and opts.paper != "A3" else [])

    if mode == LayoutMode.ONE_PAGE:
        best = None
        for paper in papers:
            for orient in ("portrait", "landscape"):
                pw, ph = printable_area(paper, orient)
                s = min(1.0, pw / si.content_width_pt, ph / max(si.content_height_pt, 1))
                if best is None or s > best[2]:
                    best = (paper, orient, s)
        plan.paper, plan.orientation, s = best
        plan.scale, plan.fit_width, plan.fit_height = None, 1, 1
        plan.estimated_min_font_pt = round(smallest * s, 2)
        if s < min_scale:
            plan.notes.append("one_page_text_below_minimum")
        return _apply_manual(plan, opts)

    title_rows = ps.title_rows
    if not title_rows and si.header_row_guess and si.header_confidence >= 0.7:
        title_rows = f"{si.header_row_guess}:{si.header_row_guess}"
    key_cols = 0
    title_cols = ps.title_cols
    if title_cols:
        key_cols = 1
    elif si.key_col_guess and si.key_confidence >= 0.7:
        key_cols = 1
        title_cols = f"{col_letter(si.key_col_guess)}:{col_letter(si.key_col_guess)}"

    if mode == LayoutMode.FIT_WIDTH:
        best = max((_evaluate(si, p, o, 0.0, title_rows, 0) for p in papers
                    for o in ("portrait", "landscape")), key=lambda e: (e["scale"], -e["pages"]))
        plan.paper, plan.orientation = best["paper"], best["orientation"]
        plan.scale, plan.fit_width, plan.fit_height = None, 1, 0
        plan.title_rows = title_rows if best["down"] > 1 else None
        plan.expected_pages = best["pages"]
        plan.estimated_min_font_pt = round(smallest * best["scale"], 2)
        if best["scale"] < min_scale:
            plan.notes.append("fit_width_text_below_minimum")
        return _apply_manual(plan, opts)

    # AUTO (and MANUAL starts from AUTO)
    if mode == LayoutMode.AUTO and ps.fit_to_page and ps.fit_width == 1:
        e = _evaluate(si, ps.paper or opts.paper, ps.orientation or "portrait", 0.0, title_rows, 0)
        if e["scale"] >= min_scale and e["across"] == 1:
            plan.keep_original = True
            plan.estimated_min_font_pt = round(smallest * e["scale"], 2)
            plan.notes.append("original_setup_is_readable")
            return plan

    evals = [_evaluate(si, p, o, min_scale, title_rows, key_cols)
             for p in papers for o in ("portrait", "landscape")]
    best = min(evals, key=lambda e: (e["across"] > 1, e["across"], e["pages"],
                                      e["paper"] != opts.paper, -round(e["scale"], 2),
                                      e["orientation"] != "portrait"))
    plan.paper, plan.orientation = best["paper"], best["orientation"]
    plan.expected_pages_across, plan.expected_pages = best["across"], best["pages"]
    plan.estimated_min_font_pt = round(smallest * best["scale"], 2)
    if best["across"] == 1:
        if best["scale"] >= 0.995:
            plan.scale = 100
        else:  # let the engine fit exactly one page wide; verification measures fonts
            plan.scale, plan.fit_width, plan.fit_height = None, 1, 0
    else:
        # round UP: rounding down would put the text just under the minimum
        plan.scale = min(100, max(10, int(math.ceil(best["scale"] * 100 - 1e-9))))
        plan.title_cols = title_cols
        plan.notes.append("horizontal_split")
        if not title_cols:
            plan.notes.append("no_key_column_detected")
    if best["down"] > 1 or best["across"] > 1:
        plan.title_rows = title_rows
    return _apply_manual(plan, opts)


def _apply_manual(plan: SheetPlan, opts: JobOptions) -> SheetPlan:
    over = opts.manual.get(plan.sheet)
    if not over:
        return plan
    allowed = {"paper", "orientation", "scale", "fit_width", "fit_height", "margins_mm",
               "title_rows", "title_cols", "center_horizontally"}
    plan = replace(plan, **{k: v for k, v in over.items() if k in allowed})
    plan.mode = LayoutMode.MANUAL
    plan.keep_original = False
    if over.get("scale"):
        plan.fit_width = plan.fit_height = None
    elif "scale" in over and over.get("fit_width"):
        plan.scale = None
    return plan


def alternative_plan(si: SheetInfo, plan: SheetPlan, opts: JobOptions, reason: str,
                     measured_min_pt: float | None = None) -> SheetPlan | None:
    """The single fallback allowed by the spec (§6 stage 7). Returns None if layout can't help.

    When the first attempt used an explicit scale and the PDF was measured, the engine's real
    rendering factor is known (Excel, for one, renders fonts a few percent smaller than the
    requested zoom), so the new scale is derived from the measurement, not the theory."""
    if plan.mode in (LayoutMode.MANUAL, LayoutMode.ONE_PAGE):
        return None  # user asked for this explicitly; report instead of overriding
    if reason == "font_below_min" and si.col_widths_pt:
        min_scale = min_scale_for(si, opts.min_font_pt)
        smallest = si.min_font_pt_source or si.base_font_pt
        if measured_min_pt and plan.scale and smallest:
            engine_factor = measured_min_pt / (smallest * plan.scale / 100)
            if 0.5 < engine_factor < 1.0:
                min_scale = min(1.0, min_scale / engine_factor)
        title_cols = plan.title_cols
        if not title_cols and si.key_col_guess and si.key_confidence >= 0.7:
            title_cols = f"{col_letter(si.key_col_guess)}:{col_letter(si.key_col_guess)}"
        alt = replace(plan, keep_original=False, fit_width=None, fit_height=None,
                      scale=min(100, max(10, int(math.ceil(min_scale * 100 - 1e-9)) + 1)),
                      title_cols=title_cols, notes=plan.notes + ["fallback_explicit_min_scale"])
        if not alt.title_rows and si.header_row_guess and si.header_confidence >= 0.7:
            alt.title_rows = f"{si.header_row_guess}:{si.header_row_guess}"
        return alt
    return None
