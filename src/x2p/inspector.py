"""Stage 2: read-only workbook inspection.

The original file is opened for reading only and never saved. openpyxl is used for
OOXML structure, xlrd for legacy XLS, and a small zip/XML reader for drawings
(images, charts, shapes) because openpyxl drops some of them.
"""
from __future__ import annotations

import posixpath
import warnings
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

from .formats import FileCheck, UnsafeFileError, check_file
from .models import CellRect, Issue, PageSetupInfo, Severity, SheetInfo, WorkbookInfo

PAPER_CODES = {1: "Letter", 5: "Legal", 8: "A3", 9: "A4", 11: "A5"}

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}
_ANCHORS = ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor")
GRID_ROWS = 60  # rows sampled for header/key heuristics
# Average glyph advance as a fraction of the font size. Deliberately generous: wrapping a
# cell that would have fitted is harmless, missing one that overflows loses data.
CHAR_WIDTH_EM = 0.55
LONG_TEXT_KEEP_LINES = 2  # keep the text of cells above this estimate (appendix candidates)
UNBREAKABLE = 999.0       # wrapped-line estimate used for text that cannot wrap


def inspect_workbook(path: str | Path, *, max_mb: int = 200) -> WorkbookInfo:
    fc = check_file(path, max_mb=max_mb)
    if fc.extension == ".xls":
        info = _inspect_xls(fc)
    else:
        info = _inspect_ooxml(fc)
    info.has_macros = fc.has_macros
    info.external_links = fc.external_links
    if fc.has_macros:
        info.issues.append(Issue("macros_present_not_executed", Severity.INFO))
    if fc.external_links:
        info.issues.append(Issue("external_links_not_updated", Severity.INFO, detail={"count": fc.external_links}))
    for s in info.sheets:
        if s.shapes:
            info.issues.append(Issue("shapes_need_review", Severity.INFO, s.name, {"count": s.shapes}))
        if s.comments:
            info.issues.append(Issue("comments_not_printed", Severity.INFO, s.name, {"count": s.comments}))
        if s.state != "visible":
            info.issues.append(Issue("sheet_hidden", Severity.INFO, s.name, {"state": s.state}))
    return info


# --------------------------------------------------------------------------- helpers

def col_width_to_pt(width_chars: float, base_font_pt: float = 11.0) -> float:
    """Excel column width (characters of max digit width) to points."""
    mdw = 7.0 * base_font_pt / 11.0          # max digit width in px at 96 dpi
    px = int(width_chars * mdw + 5) if width_chars > 0 else 0
    return px * 0.75


def _strip_abs(ref: str | None) -> str | None:
    if not ref:
        return None
    ref = ref.split("!")[-1] if "!" in ref else ref
    return ref.replace("$", "").strip() or None


def _guess_header_and_key(grid: dict[tuple[int, int], object], rect: CellRect,
                          bold_rows: set[int]) -> tuple[int | None, float, int | None, float]:
    """Suggest a header row and a key column. Pure heuristics with a confidence value."""
    header, h_conf = None, 0.0
    for r in range(rect.min_row, min(rect.min_row + 10, rect.max_row) + 1):
        vals = [grid.get((r, c)) for c in range(rect.min_col, rect.max_col + 1)]
        filled = [v for v in vals if v not in (None, "")]
        if len(filled) < 2:
            continue
        str_ratio = sum(isinstance(v, str) for v in filled) / len(filled)
        fill_ratio = len(filled) / max(1, rect.cols)
        if str_ratio >= 0.8 and fill_ratio >= 0.5:
            nxt = [grid.get((r + 1, c)) for c in range(rect.min_col, rect.max_col + 1)]
            nxt_non_str = sum(1 for v in nxt if v not in (None, "") and not isinstance(v, str))
            conf = 0.5 + 0.2 * (r in bold_rows) + 0.2 * (nxt_non_str > 0) + 0.1 * (str_ratio == 1.0)
            header, h_conf = r, round(min(conf, 0.95), 2)
            break
    key, k_conf = None, 0.0
    start = (header or rect.min_row - 1) + 1
    if rect.cols >= 4 and start <= rect.max_row:
        last = min(rect.max_row, GRID_ROWS)  # only the first GRID_ROWS rows are sampled
        col_vals = [grid.get((r, rect.min_col)) for r in range(start, last + 1)]
        filled = [v for v in col_vals if v not in (None, "")]
        if col_vals and len(filled) / len(col_vals) >= 0.9 and len(filled) >= 3:
            uniq = len({str(v) for v in filled}) / len(filled)
            if uniq >= 0.9:
                key, k_conf = rect.min_col, round(0.5 + 0.4 * uniq * (len(filled) / len(col_vals)), 2)
    return header, h_conf, key, k_conf


# --------------------------------------------------------------------------- OOXML

def _rels(zf: zipfile.ZipFile, part: str) -> dict[str, str]:
    d, f = posixpath.split(part)
    rel_path = posixpath.join(d, "_rels", f + ".rels")
    try:
        root = ET.fromstring(zf.read(rel_path))
    except KeyError:
        return {}
    out = {}
    for rel in root.findall("rel:Relationship", NS):
        if rel.get("TargetMode") == "External":
            continue
        tgt = rel.get("Target", "")
        tgt = tgt.lstrip("/") if tgt.startswith("/") else posixpath.normpath(posixpath.join(d, tgt))
        out[rel.get("Id")] = tgt
    return out


def _drawing_stats(zf: zipfile.ZipFile) -> dict[str, dict]:
    """sheet name -> {images, charts, shapes, anchor_rect}"""
    stats: dict[str, dict] = {}
    try:
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
    except KeyError:
        return stats
    wb_rels = _rels(zf, "xl/workbook.xml")
    rid_attr = f"{{{NS['r']}}}id"
    for sh in wb.findall("m:sheets/m:sheet", NS):
        part = wb_rels.get(sh.get(rid_attr, ""))
        if not part:
            continue
        st = {"images": 0, "charts": 0, "shapes": 0, "rect": None}
        for tgt in _rels(zf, part).values():
            if "/drawings/" not in tgt or not tgt.endswith(".xml"):
                continue
            try:
                droot = ET.fromstring(zf.read(tgt))
            except (KeyError, ET.ParseError):
                continue
            for tag in _ANCHORS:
                for anc in droot.findall(f"xdr:{tag}", NS):
                    st["images"] += len(anc.findall(".//xdr:pic", NS))
                    st["charts"] += len(anc.findall(".//c:chart", NS))
                    st["shapes"] += len(anc.findall(".//xdr:sp", NS)) + len(anc.findall(".//xdr:cxnSp", NS))
                    fr, to = anc.find("xdr:from", NS), anc.find("xdr:to", NS)
                    if fr is not None:
                        r0 = int(fr.findtext("xdr:row", "0", NS)) + 1
                        c0 = int(fr.findtext("xdr:col", "0", NS)) + 1
                        r1, c1 = r0, c0
                        if to is not None:
                            r1 = int(to.findtext("xdr:row", "0", NS)) + 1
                            c1 = int(to.findtext("xdr:col", "0", NS)) + 1
                        rect = CellRect(r0, c0, r1, c1)
                        st["rect"] = rect.union(st["rect"])
        stats[sh.get("name", "")] = st
    return stats


def _inspect_ooxml(fc: FileCheck) -> WorkbookInfo:
    import openpyxl
    from openpyxl.utils.cell import range_boundaries

    info = WorkbookInfo(str(fc.path), fc.extension.lstrip("."), fc.size_bytes)
    with zipfile.ZipFile(fc.path) as zf:
        drawings = _drawing_stats(zf)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            wb = openpyxl.load_workbook(fc.path, read_only=False, data_only=False,
                                        keep_vba=False, keep_links=False, rich_text=False)
        except Exception as exc:  # noqa: BLE001 - openpyxl raises many types
            raise UnsafeFileError("corrupt_file", reason=type(exc).__name__) from None
    try:
        base_font = float(getattr(wb._fonts[0], "sz", None) or 11.0)  # noqa: SLF001
    except Exception:  # noqa: BLE001
        base_font = 11.0

    for idx, ws in enumerate(wb.worksheets + list(wb.chartsheets)):
        si = SheetInfo(index=idx, name=ws.title, base_font_pt=base_font)
        si.state = ws.sheet_state
        si.index = wb.sheetnames.index(ws.title)
        if ws in wb.chartsheets:
            si.kind = "chartsheet"
            si.charts = max(1, len(getattr(ws, "_charts", [])))
            info.sheets.append(si)
            continue
        grid: dict[tuple[int, int], object] = {}
        bold_rows: set[int] = set()
        rect: CellRect | None = None
        sizes: Counter = Counter()
        text_cells: list[tuple[int, int, int, float, int]] = []  # r, c, chars, font pt, longest token
        for (r, c), cell in ws._cells.items():  # noqa: SLF001 - fastest way to skip empty cells
            v = cell.value
            if v is None or v == "":
                continue
            sz = cell.font.sz if cell.font is not None else None
            fsz = float(sz) if sz else base_font
            sizes[fsz] += 1
            if isinstance(v, str) and cell.data_type != "f":
                al = cell.alignment
                if al is None or not (al.wrap_text or al.shrink_to_fit):
                    longest = len(v)  # wrap is off: line breaks do not start new lines
                    token = max((len(t) for t in v.split()), default=0)
                    text_cells.append((r, c, longest, fsz, token))
            if cell.data_type == "f":
                si.formula_cells += 1
            else:
                si.value_cells += 1
            if cell.comment is not None:
                si.comments += 1
            if r <= GRID_ROWS:
                grid[(r, c)] = v
                if cell.font is not None and cell.font.b:
                    bold_rows.add(r)
            rect = CellRect(r, c, r, c).union(rect)
        dstat = drawings.get(ws.title, {})
        si.images, si.charts, si.shapes = dstat.get("images", 0), dstat.get("charts", 0), dstat.get("shapes", 0)
        if dstat.get("rect") is not None:
            rect = dstat["rect"].union(rect)
        si.content = rect
        if sizes:
            si.base_font_pt = sizes.most_common(1)[0][0]
            si.min_font_pt_source = min(sizes)
        si.merged_ranges = len(ws.merged_cells.ranges)
        si.tables = len(ws.tables)
        si.conditional_formats = len(ws.conditional_formatting)
        si.right_to_left = bool(ws.sheet_view.rightToLeft)

        # dimensions
        fmt = ws.sheet_format
        default_w = fmt.defaultColWidth or ((fmt.baseColWidth or 8) + 0.43 if fmt.baseColWidth else 8.43)
        default_h = fmt.defaultRowHeight or 15.0
        col_w: dict[int, float] = {}
        col_hidden: set[int] = set()
        for cd in ws.column_dimensions.values():
            lo, hi = cd.min or 0, cd.max or 0
            for c in range(lo, hi + 1):
                if cd.width:
                    col_w[c] = cd.width
                if cd.hidden:
                    col_hidden.add(c)
        ps = ws.page_setup
        pinfo = PageSetupInfo()
        pinfo.orientation = ps.orientation or None
        try:
            pinfo.paper = PAPER_CODES.get(int(ps.paperSize)) if ps.paperSize else None
        except (TypeError, ValueError):
            pinfo.paper = None
        pinfo.scale = int(ps.scale) if ps.scale else None
        pinfo.fit_to_page = bool(ws.sheet_properties.pageSetUpPr and ws.sheet_properties.pageSetUpPr.fitToPage)
        if pinfo.fit_to_page:
            pinfo.fit_width = int(ps.fitToWidth) if ps.fitToWidth is not None else 1
            pinfo.fit_height = int(ps.fitToHeight) if ps.fitToHeight is not None else 1
        pa = ws.print_area
        if isinstance(pa, (list, tuple)):
            pa = ",".join(pa)
        pinfo.print_area = ",".join(_strip_abs(p) or "" for p in pa.split(",")) if pa else None
        pinfo.title_rows = _strip_abs(ws.print_title_rows)
        pinfo.title_cols = _strip_abs(ws.print_title_cols)
        pinfo.row_breaks = len(ws.row_breaks.brk)
        pinfo.col_breaks = len(ws.col_breaks.brk)
        hf = ws.HeaderFooter
        pinfo.has_header_footer = any(
            getattr(getattr(part, pos), "text", None)
            for part in (hf.oddHeader, hf.oddFooter) for pos in ("left", "center", "right"))
        si.page_setup = pinfo
        # Layout dimensions: the print area when defined (single range), otherwise content.
        lrect = rect
        if pinfo.print_area and "," not in pinfo.print_area:
            try:
                c0, r0, c1, r1 = range_boundaries(pinfo.print_area)
                if None not in (c0, r0, c1, r1):
                    lrect = CellRect(r0, c0, r1, c1)
            except ValueError:
                pass
        if lrect:
            for c in range(lrect.min_col, lrect.max_col + 1):
                if c in col_hidden:
                    si.hidden_cols += 1
                    continue
                si.col_widths_pt.append(col_width_to_pt(col_w.get(c, default_w), base_font))
            for r in range(lrect.min_row, lrect.max_row + 1):
                rd = ws.row_dimensions.get(r) if r in ws.row_dimensions else None
                if rd is not None and rd.hidden:
                    si.hidden_rows += 1
                    continue
                si.row_heights_pt.append(float(rd.ht) if rd is not None and rd.ht else float(default_h))
            si.layout_rect = lrect
            hidden_rows = {r for r in range(lrect.min_row, lrect.max_row + 1)
                           if r in ws.row_dimensions and ws.row_dimensions[r].hidden}
            si.hidden_row_numbers = sorted(hidden_rows)
            merged = set()
            for mr in ws.merged_cells.ranges:
                if (mr.max_row - mr.min_row + 1) * (mr.max_col - mr.min_col + 1) <= 100_000:
                    merged.update((rr, cc) for rr in range(mr.min_row, mr.max_row + 1)
                                  for cc in range(mr.min_col, mr.max_col + 1))
            for r, c, chars, fsz, token in text_cells:
                if (r, c) in merged or r in hidden_rows or c in col_hidden:
                    continue
                if not (lrect.min_row <= r <= lrect.max_row and lrect.min_col <= c <= lrect.max_col):
                    continue
                width = col_width_to_pt(col_w.get(c, default_w), base_font)
                usable = max(width * 0.95, 1.0)
                lines = chars * fsz * CHAR_WIDTH_EM / usable
                if token * fsz * CHAR_WIDTH_EM > usable:
                    lines = max(lines, UNBREAKABLE)  # a word/URL wider than the column cannot wrap
                if lines > 1:
                    si.overflow.append((r, c, round(lines, 1)))
                    si.overflow_cells += 1
                    if lines > LONG_TEXT_KEEP_LINES:
                        si.long_texts[(r, c)] = ws._cells[(r, c)].value  # noqa: SLF001
        if rect:
            (si.header_row_guess, si.header_confidence,
             si.key_col_guess, si.key_confidence) = _guess_header_and_key(grid, rect, bold_rows)
        if si.long_texts:  # context for the appendix: column captions and record keys
            hr = si.header_row_guess or (rect.min_row if rect else 1)
            for c in {c for _, c in si.long_texts}:
                h = ws._cells.get((hr, c))  # noqa: SLF001
                si.header_texts[c] = str(h.value) if h is not None and h.value is not None else ""
            kc = si.key_col_guess or (rect.min_col if rect else 1)
            for r in {r for r, _ in si.long_texts}:
                k = ws._cells.get((r, kc))  # noqa: SLF001
                si.key_texts[r] = str(k.value)[:80] if k is not None and k.value is not None else ""
        info.sheets.append(si)

    info.sheets.sort(key=lambda s: s.index)
    wb.close()
    if any(s.formula_cells for s in info.sheets):
        _check_cached_results(fc, info)
    return info


ERROR_VALUES = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA",
                "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#UNKNOWN!", "#BUSY!", "#CONNECT!"}


def _check_cached_results(fc: FileCheck, info: WorkbookInfo) -> None:
    """Formulas without a saved result (files written by programs, not by Excel) are computed
    by the engine at conversion time; error results (#N/A ...) are data and printed as such.
    Both are reported so the reader knows."""
    import openpyxl
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            wb = openpyxl.load_workbook(fc.path, read_only=False, data_only=True, keep_links=False)
        except Exception:  # noqa: BLE001
            return
    by_name = {s.name: s for s in info.sheets}
    for ws in wb.worksheets:
        si = by_name.get(ws.title)
        if si is None or not si.formula_cells:
            continue
        uncached = errors = 0
        for cell in ws._cells.values():  # noqa: SLF001
            v = cell.value
            if isinstance(v, str) and v in ERROR_VALUES:
                errors += 1
            elif v is None and cell.data_type == "n":
                uncached += 1
        si.uncached_formulas = min(uncached, si.formula_cells)
        si.error_values = errors
        if si.uncached_formulas:
            info.issues.append(Issue("formulas_computed_by_engine", Severity.INFO, si.name,
                                     {"count": si.uncached_formulas}))
        if errors:
            info.issues.append(Issue("error_values_present", Severity.INFO, si.name, {"count": errors}))
    wb.close()


# --------------------------------------------------------------------------- XLS

def _inspect_xls(fc: FileCheck) -> WorkbookInfo:
    import xlrd

    info = WorkbookInfo(str(fc.path), "xls", fc.size_bytes)
    try:
        book = xlrd.open_workbook(str(fc.path), formatting_info=True, on_demand=True)
    except xlrd.biffh.XLRDError as exc:
        if "encrypt" in str(exc).lower():
            raise UnsafeFileError("password_protected") from None
        raise UnsafeFileError("corrupt_file", reason="xlrd") from None
    except Exception:  # noqa: BLE001
        raise UnsafeFileError("corrupt_file", reason="xlrd") from None
    try:
        base_font = book.font_list[0].height / 20.0 if book.font_list else 10.0
    except Exception:  # noqa: BLE001
        base_font = 10.0
    states = {0: "visible", 1: "hidden", 2: "veryHidden"}
    for idx in range(book.nsheets):
        sh = book.sheet_by_index(idx)
        si = SheetInfo(index=idx, name=sh.name, base_font_pt=base_font)
        si.state = states.get(sh.visibility, "visible")
        grid: dict[tuple[int, int], object] = {}
        rect: CellRect | None = None
        for r in range(sh.nrows):
            for c, cell in enumerate(sh.row(r)):
                if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    continue
                if cell.ctype == xlrd.XL_CELL_TEXT and cell.value == "":
                    continue
                si.value_cells += 1
                if r < GRID_ROWS:
                    grid[(r + 1, c + 1)] = cell.value
                rect = CellRect(r + 1, c + 1, r + 1, c + 1).union(rect)
        si.content = si.layout_rect = rect
        si.merged_ranges = len(sh.merged_cells)
        si.right_to_left = bool(getattr(sh, "rtl", False) or getattr(sh, "show_in_rtl", False))
        if rect:
            default_w = (sh.defcolwidth or 8) + 0.43 if sh.standardwidth is None else sh.standardwidth / 256
            for c in range(rect.min_col - 1, rect.max_col):
                ci = sh.colinfo_map.get(c)
                if ci is not None and ci.hidden:
                    si.hidden_cols += 1
                    continue
                w = ci.width / 256 if ci is not None else default_w
                si.col_widths_pt.append(col_width_to_pt(w, base_font))
            default_h = (sh.default_row_height or 255) / 20
            for r in range(rect.min_row - 1, rect.max_row):
                ri = sh.rowinfo_map.get(r)
                if ri is not None and ri.hidden:
                    si.hidden_rows += 1
                    continue
                si.row_heights_pt.append(ri.height / 20 if ri is not None else default_h)
            (si.header_row_guess, si.header_confidence,
             si.key_col_guess, si.key_confidence) = _guess_header_and_key(grid, rect, set())
        si.page_setup.row_breaks = len(getattr(sh, "horizontal_page_breaks", []) or [])
        si.page_setup.col_breaks = len(getattr(sh, "vertical_page_breaks", []) or [])
        info.sheets.append(si)
        book.unload_sheet(idx)
    info.issues.append(Issue("xls_objects_not_inspected", Severity.INFO))
    book.release_resources()
    return info


def merged_count(info: WorkbookInfo) -> int:
    return sum(s.merged_ranges for s in info.sheets)


def summarize(info: WorkbookInfo) -> Counter:
    c: Counter = Counter()
    for s in info.sheets:
        c["sheets"] += 1
        c["hidden_sheets"] += s.state != "visible"
        for k in ("value_cells", "formula_cells", "merged_ranges", "images", "charts", "shapes", "tables"):
            c[k] += getattr(s, k)
    return c

