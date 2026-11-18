"""Microsoft Excel worker (optional engine, Windows only). Standalone; uses pywin32.

Usage: python excel_worker.py job.json result.json

Safety: a private Excel instance (DispatchEx), macros force-disabled, links not updated,
events off, the workbook is a temporary working copy closed without saving.
"""
import json
import math
import sys
import traceback
from pathlib import Path

PAPER = {"Letter": 1, "Legal": 5, "A3": 8, "A4": 9, "A5": 11}
XL_CONSTANTS, XL_FORMULAS, XL_VISIBLE = 2, -4123, 12
MM_TO_PT = 72 / 25.4


def apply_plan(xl, ws, plan):
    if plan.get("keep_original"):
        return
    for addr, marker in plan.get("endnotes") or []:
        ws.Range(addr).Value = marker
    for rng in plan.get("wrap_ranges") or []:
        ws.Range(rng).WrapText = True
    for block in plan.get("autofit_rows") or []:
        ws.Range(block).Rows.AutoFit()
    ps = ws.PageSetup
    try:
        xl.PrintCommunication = False
    except Exception:
        pass
    try:
        try:
            ps.PaperSize = PAPER.get(plan.get("paper") or "A4", 9)
        except Exception:
            pass  # printer driver may not offer this size; verification reports the result
        ps.Orientation = 2 if plan.get("orientation") == "landscape" else 1
        if plan.get("mode") == "manual" and plan.get("margins_mm"):
            top, bottom, left, right = plan["margins_mm"]
            ps.TopMargin, ps.BottomMargin = top * MM_TO_PT, bottom * MM_TO_PT
            ps.LeftMargin, ps.RightMargin = left * MM_TO_PT, right * MM_TO_PT
        if plan.get("scale"):
            ps.Zoom = int(plan["scale"])
        elif plan.get("fit_width") or plan.get("fit_height"):
            ps.Zoom = False
            ps.FitToPagesWide = int(plan["fit_width"]) if plan.get("fit_width") else False
            ps.FitToPagesTall = int(plan["fit_height"]) if plan.get("fit_height") else False
        ps.CenterHorizontally = bool(plan.get("center_horizontally"))
        if plan.get("title_rows"):
            a, b = plan["title_rows"].split(":")
            ps.PrintTitleRows = "$%s:$%s" % (a, b)
        if plan.get("title_cols"):
            a, b = plan["title_cols"].split(":")
            ps.PrintTitleColumns = "$%s:$%s" % (a, b)
    finally:
        try:
            xl.PrintCommunication = True
        except Exception:
            pass


def split_chunks(text, size):
    out, i = [], 0
    while i < len(text):
        j = min(len(text), i + size)
        if j < len(text):
            k = max(text.rfind(" ", i + size // 2, j), text.rfind("\n", i + size // 2, j))
            if k > i:
                j = k + 1
        out.append(text[i:j])
        i = j
    return out or [""]


def assemble_cards(headers, records, key_idx, spec):
    """records: [(source_row, [displayed strings])] -> (rows, bold_row_indexes)"""
    rows, bold = [], []
    for n, (src_row, vals) in enumerate(records, 1):
        title = "[%d]" % n
        if key_idx is not None and vals[key_idx].strip():
            title += "  %s: %s" % (headers[key_idx], vals[key_idx])
        title += "  \u00b7  %s %d" % (spec["row_label"], src_row)
        bold.append(len(rows))
        rows.append([title, ""])
        for h, v in zip(headers, vals):
            if not v.strip():
                continue
            chunks = split_chunks(v, spec["chunk"])
            rows.append([h, chunks[0]])
            rows.extend(["", c] for c in chunks[1:])
        rows.append(["", ""])
    return rows, bold


def unique_name(wb, name):
    names = {wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)}
    base, k = name, 2
    while name in names:
        name = "%s %d" % (base[:28], k)
        k += 1
    return name


def build_cards(wb, source, spec):
    src = wb.Sheets(source)
    r0, r1, c0, c1 = spec["first_row"], spec["last_row"], spec["first_col"], spec["last_col"]
    cols = [c for c in range(c0, c1 + 1) if not src.Columns(c).Hidden]
    # working copy only: wide columns so .Text never returns "####" for numbers
    for c in cols:
        src.Columns(c).ColumnWidth = 100
    hr = spec.get("header_row")
    headers = [(str(src.Cells(hr, c).Text) if hr else "") or col_letter(c) for c in cols]
    data = src.Range(src.Cells(r0, c0), src.Cells(r1, c1)).Value2
    if not isinstance(data, tuple):
        data = ((data,),)
    records = []
    for i, row in enumerate(data):
        r = r0 + i
        if src.Rows(r).Hidden:
            continue
        vals = []
        for c in cols:
            v = row[c - c0]
            if v is None:
                vals.append("")
            elif isinstance(v, str):
                vals.append(v)
            else:  # number/date: the displayed (formatted) text
                vals.append(str(src.Cells(r, c).Text))
        if any(x.strip() for x in vals):
            records.append((r, vals))
    key_idx = cols.index(spec["key_col"]) if spec.get("key_col") in cols else None
    rows, bold = assemble_cards(headers, records, key_idx, spec)
    ws = wb.Worksheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws.Name = unique_name(wb, spec["sheet_name"])
    rng = ws.Range(ws.Cells(1, 1), ws.Cells(len(rows), 2))
    rng.NumberFormat = "@"
    rng.Value = tuple(tuple(r) for r in rows)
    for i, w in enumerate(spec["widths"], 1):
        ws.Columns(i).ColumnWidth = w
    rng.WrapText = True
    rng.VerticalAlignment = -4160
    rng.Font.Size = 10
    ws.Range(ws.Cells(1, 1), ws.Cells(len(rows), 1)).Font.Bold = True
    for b in bold:
        ws.Range(ws.Cells(b + 1, 1), ws.Cells(b + 1, 2)).Interior.Color = 0xF7EBDD  # BGR
    if spec.get("rtl"):
        ws.DisplayRightToLeft = True
    rng.Rows.AutoFit()
    return ws.Name


def create_appendix(wb, name, spec):
    """Extra sheet holding the long texts. Cells are formatted as text ("@") before
    writing, so no value can be interpreted as a formula or a number."""
    ws = wb.Worksheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws.Name = name
    rows = spec["rows"]
    ncols = len(rows[0])
    rng = ws.Range(ws.Cells(1, 1), ws.Cells(len(rows), ncols))
    rng.NumberFormat = "@"
    rng.Value = tuple(tuple(str(v) for v in r) for r in rows)
    for i, w in enumerate(spec["widths"], 1):
        ws.Columns(i).ColumnWidth = w
    rng.WrapText = True
    rng.VerticalAlignment = -4160  # top
    rng.Font.Size = 10
    for r in spec.get("bold_rows") or []:
        ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + 1, ncols)).Font.Bold = True
    if spec.get("rtl"):
        ws.DisplayRightToLeft = True
    rng.Rows.AutoFit()
    return ws


def collect_texts(ws, max_cells, max_slow=1500):
    try:
        rng = ws.Range(ws.PageSetup.PrintArea) if ws.PageSetup.PrintArea else ws.UsedRange
    except Exception:
        rng = ws.UsedRange
    parts = []
    for typ in (XL_CONSTANTS, XL_FORMULAS):
        try:
            sc = rng.SpecialCells(typ)
            sc = sc.SpecialCells(XL_VISIBLE)
            parts.append(sc)
        except Exception:
            continue  # "No cells were found"
    # One bulk Value2 read per area (fast). Strings display as stored; only non-string
    # cells need the per-cell `.Text` call (slow COM round trip), which is budgeted.
    texts, numeric = [], []
    for p in parts:
        for area in p.Areas:
            r0, c0 = area.Row, area.Column
            vals = area.Value2
            if not isinstance(vals, tuple):
                vals = ((vals,),)
            for i, row in enumerate(vals):
                for j, v in enumerate(row):
                    if v is None:
                        continue
                    if isinstance(v, str):
                        if v.strip():
                            texts.append([col_letter(c0 + j) + str(r0 + i), v])
                    else:
                        numeric.append((r0 + i, c0 + j))
    budget = max(1, min(max_slow, max_cells or max_slow))  # each .Text is a COM round trip
    stride = max(1, int(math.ceil(len(numeric) / float(budget))))
    sampled = stride > 1
    if max_cells and len(texts) > max_cells:
        s2 = int(math.ceil(len(texts) / float(max_cells)))
        texts, sampled = texts[::s2], True
    hashes = []
    for r, c in numeric[::stride]:
        t = ws.Cells(r, c).Text
        if not t or not str(t).strip():
            continue
        addr = col_letter(c) + str(r)
        if set(str(t).strip()) == {"#"}:   # a number/date wider than its column
            hashes.append(addr)
            continue
        texts.append([addr, str(t)])
    return texts, sampled, hashes


def col_letter(n):
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def main(job_path, result_path):
    job = json.loads(Path(job_path).read_text(encoding="utf-8"))
    result = {"ok": False, "sheets": [], "engine_version": ""}
    xl = wb = None
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.ScreenUpdating = False
        xl.EnableEvents = False
        xl.AskToUpdateLinks = False
        xl.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        result["engine_version"] = "Microsoft Excel %s" % xl.Version
        wb = xl.Workbooks.Open(str(Path(job["input"]).resolve()), UpdateLinks=0, ReadOnly=False,
                               IgnoreReadOnlyRecommended=True, Notify=False, AddToMru=False)
        out_dir = Path(job["out_dir"])
        for i, spec in enumerate(job["sheets"]):
            r = {"name": spec["name"], "pdf": None, "error": None, "texts": [], "sampled": False,
                 "hash_cells": []}
            try:
                target = spec["name"]
                if spec.get("create_appendix"):
                    create_appendix(wb, spec["name"], spec["create_appendix"])
                if (spec["plan"] or {}).get("cards"):
                    target = build_cards(wb, spec["name"], spec["plan"]["cards"])
                ws = wb.Sheets(target)
                if ws.Visible != -1:
                    ws.Visible = -1
                apply_plan(xl, ws, spec["plan"])
                out = out_dir / ("sheet_%03d.pdf" % i)
                try:
                    ws.ExportAsFixedFormat(Type=0, Filename=str(out), Quality=0,
                                           IncludeDocProperties=True, IgnorePrintAreas=False,
                                           OpenAfterPublish=False)
                except Exception as exc:
                    sys.stderr.write("export %d: %s\n" % (i, exc))
                    r["error"] = "nothing_to_print"
                if out.exists() and out.stat().st_size > 0:
                    r["pdf"] = str(out)
                    r["error"] = None
                is_grid = getattr(ws, "UsedRange", None) is not None and ws.Type == -4167
                if job.get("collect_text", True) and is_grid:
                    r["texts"], r["sampled"], r["hash_cells"] = collect_texts(
                        ws, job.get("max_text_cells", 30000), job.get("max_slow_cells", 1500))
            except Exception as exc:
                r["error"] = r["error"] or "sheet_export_failed"
                sys.stderr.write("sheet %d: %s\n" % (i, type(exc).__name__))
            result["sheets"].append(r)
        result["ok"] = True
    except Exception:
        result["error"] = "engine_failed"
        traceback.print_exc()
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass
        Path(result_path).write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
