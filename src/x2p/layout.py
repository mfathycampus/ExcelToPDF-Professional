"""Stages 3-4: engine-neutral page planning.

The planner works on estimates (column widths/row heights from the file). The engine
does the real layout, and verification then *measures* the PDF; when the measurement
violates a hard rule the pipeline asks `alternative_plan` for exactly one retry.
"""
from __future__ import annotations

import math
import re
from dataclasses import replace

from .metrics import CELL_PAD_PT, pt_to_col_chars, text_width_pt, wrapped_lines
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


# Column widening (working copy). Exports from other systems often leave every column at the
# default width, so text is wrapped into tall narrow columns or pushed to the appendix and dates
# show ###. Numbers/dates always get the width they need (### is data loss). Text columns may
# grow up to a cap; the planner tries several caps and keeps the one with the fewest pages.
WIDTH_CAPS_PT = (0.0, 100.0, 140.0, 190.0, 250.0, 330.0, 450.0, 600.0)
APPENDIX_CHARS_PER_PAGE = 2600
TEXT_SLACK = 1.08   # engines differ in glyph widths (LibreOffice substitutes fonts): leave room


def column_widths_for(si: SheetInfo, cap_pt: float) -> dict[int, float]:
    """New widths (points) per column; only columns that grow are listed."""
    current = dict(zip(si.visible_cols, si.col_widths_pt))
    new: dict[int, float] = {}
    for c, need in si.value_need_pt.items():
        if c in current and need > current[c] + 0.5:
            new[c] = math.ceil(need)
    if cap_pt > 0:
        header = si.header_row_guess if si.header_confidence >= 0.6 else None
        need_text: dict[int, float] = {}
        for r, c, em, fsz, token in si.text_cells:
            if c not in current:
                continue
            # header: its longest word must fit (the header itself may wrap)
            w = text_width_pt(token if r == header else em, fsz) * TEXT_SLACK + CELL_PAD_PT
            if w > need_text.get(c, 0.0):
                need_text[c] = w
        for c, w in need_text.items():
            base = max(current[c], new.get(c, 0.0))
            if w > base + 0.5:
                new[c] = float(math.ceil(max(base, min(w, cap_pt))))
    return new


def with_widths(si: SheetInfo, widths: dict[int, float]) -> SheetInfo:
    """A copy of the sheet description as if the columns had these widths."""
    if not widths:
        return si
    s2 = replace(si)
    width_of = dict(zip(si.visible_cols, si.col_widths_pt))
    width_of.update(widths)
    s2.col_widths_pt = [width_of[c] for c in si.visible_cols]
    s2.overflow = []
    for r, c, em, fsz, token in si.text_cells:
        lines = wrapped_lines(em, token, fsz, width_of.get(c, 1.0))
        if lines > 1:
            s2.overflow.append((r, c, round(lines, 1)))
    s2.overflow_cells = len(s2.overflow)
    s2.widened_cols = dict(widths)
    return s2


def _appendix_pages(si: SheetInfo, plan: SheetPlan) -> float:
    if not plan.long_cells:
        return 0.0
    chars = sum(len(str(si.long_texts.get(k, ""))) for k in plan.long_cells)
    return chars / APPENDIX_CHARS_PER_PAGE + len(plan.long_cells) * 0.04


def plan_sheet(si: SheetInfo, opts: JobOptions) -> SheetPlan:
    mode = opts.per_sheet_mode.get(si.name, opts.mode)
    manual = opts.manual.get(si.name) or {}
    widen = opts.widen_columns and manual.get("widen_columns") is not False
    if not widen or not si.visible_cols:
        caps = (None,)                      # original widths exactly
    elif mode in (LayoutMode.AUTO, LayoutMode.MANUAL):
        caps = WIDTH_CAPS_PT
    else:
        caps = (0.0,)                       # fit modes: only numbers/dates grow (no ###)
    best = None
    for cap in caps:
        s2 = si if cap is None else with_widths(si, column_widths_for(si, cap))
        plan = _plan_with_wrapping(s2, opts)
        if plan.mode == LayoutMode.CARDS or plan.keep_original:
            return plan
        score = (math.ceil(plan.expected_pages + _appendix_pages(s2, plan)), len(plan.long_cells),
                 plan.expected_pages_across, plan.wrapped_cells)
        if best is None or score < best[0]:
            best = (score, plan, s2)
    _score, plan, s2 = best
    rect = si.layout_rect or si.content
    if (not si.has_borders and si.kind == "worksheet" and rect is not None
            and len(si.col_widths_pt) >= 2 and len(si.row_heights_pt) >= 3 and opts.gridlines):
        plan.print_gridlines = True
        plan.notes.append("gridlines_added")
    if s2.widened_cols:
        _wrap_text_columns(s2, plan)
        for c, pt in sorted(s2.widened_cols.items()):
            letter = col_letter(c)
            plan.col_widths[letter] = pt_to_col_chars(pt, si.wb_font_pt)
            plan.col_widths_pt[letter] = pt
        plan.notes.insert(0, "columns_widened")
    if plan.expected_pages_across >= CARDS_SUGGEST_ACROSS and plan.mode == LayoutMode.AUTO:
        plan.notes.append("cards_suggested")
    if manual:
        _apply_manual_extras(si, plan, manual)
    _apply_template(si, plan, opts)
    return plan


def _apply_manual_extras(si: SheetInfo, plan: SheetPlan, manual: dict) -> None:
    """Gridlines on/off and a page break every N visible data rows (header row repeats)."""
    if "print_gridlines" in manual:
        plan.print_gridlines = bool(manual["print_gridlines"])
        if not plan.print_gridlines and "gridlines_added" in plan.notes:
            plan.notes.remove("gridlines_added")
    n = int(manual.get("break_every") or 0)
    rect = si.layout_rect or si.content
    if n > 0 and rect is not None:
        header = si.header_row_guess if si.header_confidence >= 0.6 else None
        first = (header + 1) if header else rect.min_row
        hidden = set(si.hidden_row_numbers)
        visible = [r for r in range(first, rect.max_row + 1) if r not in hidden]
        plan.row_breaks = visible[n::n][:1000]          # Excel allows 1026 manual breaks per sheet
        if plan.row_breaks:
            plan.fit_height = 0 if plan.fit_width else plan.fit_height
            plan.notes.append("manual_page_breaks")


def _wrap_text_columns(si: SheetInfo, plan: SheetPlan) -> None:
    """Safety net once widths were changed from estimates: wrap every text column, so a text the
    estimate under-measured goes to a second line instead of being clipped (wrapping is lossless;
    rows get their optimal height)."""
    cols = {}
    for r, c, *_ in si.text_cells:
        lo, hi = cols.get(c, (r, r))
        cols[c] = (min(lo, r), max(hi, r))
    for rng in plan.wrap_ranges:
        m = re.fullmatch(r"([A-Z]+)(\d+):[A-Z]+(\d+)", rng)
        if m:
            c = col_index(m.group(1))
            lo, hi = cols.get(c, (int(m.group(2)), int(m.group(3))))
            cols[c] = (min(lo, int(m.group(2))), max(hi, int(m.group(3))))
    if not cols:
        return
    plan.wrap_ranges = [f"{col_letter(c)}{lo}:{col_letter(c)}{hi}" for c, (lo, hi) in sorted(cols.items())]
    lo = min(v[0] for v in cols.values())
    hi = max(v[1] for v in cols.values())
    plan.autofit_rows = visible_row_blocks(si, lo, hi)


def col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def _plan_with_wrapping(si: SheetInfo, opts: JobOptions) -> SheetPlan:
    plan = _plan_sheet(si, opts)
    if plan.mode == LayoutMode.CARDS:
        return plan
    if opts.wrap_overflow and si.overflow and not plan.keep_original:
        add_overflow_wrapping(si, plan, opts)
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
    header = si.header_row_guess if si.header_confidence >= 0.6 else None
    long_cells, cols = [], {}
    for r, c, lines in si.overflow:
        if to_appendix and r != header and lines > opts.long_text_lines and (r, c) in si.long_texts:
            long_cells.append((r, c))
            continue
        lo, hi = cols.get(c, (r, r))
        cols[c] = (min(lo, r), max(hi, r))
    plan.long_cells = sorted(long_cells)
    plan.wrapped_cells = len(si.overflow) - len(long_cells)
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

    title_rows, title_cols = suggested_titles(si, opts)
    key_cols = _range_cols(title_cols)

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


def suggested_titles(si: SheetInfo, opts: JobOptions | None = None) -> tuple[str | None, str | None]:
    """Rows/columns repeated on every page. Precedence: the user's template, the file's print
    titles, the file's frozen panes, then the detected header row / key column."""
    tpl = (opts.templates.get(si.name) if opts else None) or {}
    ps = si.page_setup
    rows = ps.title_rows or si.frozen_rows
    if not rows and si.header_row_guess and si.header_confidence >= 0.7:
        rows = f"{si.header_row_guess}:{si.header_row_guess}"
    cols = ps.title_cols or si.frozen_cols
    if not cols and si.key_col_guess and si.key_confidence >= 0.7:
        cols = f"{col_letter(si.key_col_guess)}:{col_letter(si.key_col_guess)}"
    if "title_rows" in tpl:
        rows = tpl["title_rows"] or None
    if "title_cols" in tpl:
        cols = tpl["title_cols"] or None
    return rows, cols


def _range_cols(cols: str | None) -> int:
    if not cols:
        return 0
    a, _, b = cols.partition(":")
    try:
        return max(1, col_index(b or a) - col_index(a) + 1)
    except (TypeError, ValueError):
        return 1


def _apply_template(si: SheetInfo, plan: SheetPlan, opts: JobOptions) -> None:
    """Page order and row numbers for split tables (defaults; the template may override)."""
    tpl = opts.templates.get(si.name) or {}
    if plan.mode in (LayoutMode.CARDS, LayoutMode.ONE_PAGE) or plan.keep_original:
        return
    if plan.expected_pages_across > 1:
        plan.page_order = tpl.get("page_order") or "over"
        if not plan.title_cols and tpl.get("row_numbers", True):
            plan.print_headings = True
            if "no_key_column_detected" in plan.notes:
                plan.notes.remove("no_key_column_detected")
            plan.notes.append("row_numbers_added")
    elif tpl.get("page_order"):
        plan.page_order = tpl["page_order"]


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
