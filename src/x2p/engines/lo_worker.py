"""LibreOffice worker. Runs under LibreOffice's own Python. Standalone: no x2p imports.

Usage: python lo_worker.py job.json result.json

Safety: isolated profile, macros never executed, links never updated, no recalculation
on load, the document is a temporary working copy and is never stored.
"""
import json
import math
import os
import subprocess
import sys
import time
import traceback
import uuid
from pathlib import Path

import uno  # noqa: E402  (provided by LibreOffice)
from com.sun.star.beans import PropertyValue  # noqa: E402
from com.sun.star.table import CellRangeAddress  # noqa: E402

PAPER_100MM = {"A4": (21000, 29700), "A3": (29700, 42000), "Letter": (21590, 27940),
               "Legal": (21590, 35560), "A5": (14800, 21000)}

REGISTRY = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="DisableMacrosExecution" oor:op="fuse"><value>true</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>1</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>1</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>1</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Misc"><prop oor:name="FirstRun" oor:op="fuse"><value>false</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Misc"><prop oor:name="ShowTipOfTheDay" oor:op="fuse"><value>false</value></prop></item>
</oor:items>
"""


def pv(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def file_url(path):
    return uno.systemPathToFileUrl(os.path.abspath(path))


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(letters):
    n = 0
    for ch in letters.strip().upper():
        n = n * 26 + ord(ch) - 64
    return n


def connect(job):
    profile = Path(job["profile_dir"])
    (profile / "user").mkdir(parents=True, exist_ok=True)
    (profile / "user" / "registrymodifications.xcu").write_text(REGISTRY, encoding="utf-8")
    pipe = "x2p_" + uuid.uuid4().hex
    flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
    proc = subprocess.Popen([job["soffice"], "-env:UserInstallation=" + file_url(profile),
                             "--headless", "--invisible", "--nologo", "--nodefault", "--norestore",
                             "--nolockcheck", "--nofirststartwizard",
                             "--accept=pipe,name=%s;urp;StarOffice.ComponentContext" % pipe],
                            creationflags=flags)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    deadline = time.time() + 90
    while True:
        try:
            ctx = resolver.resolve("uno:pipe,name=%s;urp;StarOffice.ComponentContext" % pipe)
            break
        except Exception:  # NoConnectException until soffice is up
            if time.time() > deadline or proc.poll() is not None:
                raise RuntimeError("engine_start_failed")
            time.sleep(0.4)
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return proc, ctx, desktop


def engine_version(ctx):
    try:
        cp = ctx.ServiceManager.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)
        node = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess",
                                              (pv("nodepath", "/org.openoffice.Setup/Product"),))
        return "%s %s" % (node.getByName("ooName"), node.getByName("ooSetupVersionAboutBox"))
    except Exception:
        return "LibreOffice"


def apply_plan(doc, sheet, idx, plan):
    notes = []
    if plan.get("keep_original"):
        return notes
    for addr, marker in plan.get("endnotes") or []:
        sheet.getCellRangeByName(addr).setString(marker)
    for rng in plan.get("wrap_ranges") or []:
        sheet.getCellRangeByName(rng).IsTextWrapped = True
    for block in plan.get("autofit_rows") or []:
        a, b = block.split(":")
        sheet.getCellRangeByName("A%s:A%s" % (a, b)).Rows.OptimalHeight = True
    style = doc.StyleFamilies.getByName("PageStyles").getByName(sheet.PageStyle)
    w, h = PAPER_100MM.get(plan.get("paper") or "A4", PAPER_100MM["A4"])
    if plan.get("orientation") == "landscape":
        style.IsLandscape = True
        style.Width, style.Height = h, w
    else:
        style.IsLandscape = False
        style.Width, style.Height = w, h
    if plan.get("mode") == "manual" and plan.get("margins_mm"):
        top, bottom, left, right = plan["margins_mm"]
        style.TopMargin, style.BottomMargin = int(top * 100), int(bottom * 100)
        style.LeftMargin, style.RightMargin = int(left * 100), int(right * 100)
    if plan.get("scale"):
        style.ScaleToPages = 0
        style.ScaleToPagesX = 0
        style.ScaleToPagesY = 0
        style.PageScale = int(plan["scale"])
    elif plan.get("fit_width") or plan.get("fit_height"):
        style.ScaleToPages = 0
        style.ScaleToPagesX = int(plan.get("fit_width") or 0)
        style.ScaleToPagesY = int(plan.get("fit_height") or 0)
    style.CenterHorizontally = bool(plan.get("center_horizontally"))
    if plan.get("title_rows"):
        a, b = (int(x) for x in plan["title_rows"].split(":"))
        sheet.setTitleRows(CellRangeAddress(idx, 0, a - 1, 0, b - 1))
        sheet.setPrintTitleRows(True)
    if plan.get("title_cols"):
        a, b = plan["title_cols"].split(":")
        sheet.setTitleColumns(CellRangeAddress(idx, col_index(a) - 1, 0, col_index(b) - 1, 0))
        sheet.setPrintTitleColumns(True)
    return notes


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


def unique_name(sheets, name):
    base, k = name, 2
    while sheets.hasByName(name):
        name = "%s %d" % (base[:28], k)
        k += 1
    return name


def build_cards(doc, source, spec):
    sheets = doc.Sheets
    src = sheets.getByName(source)
    r0, r1, c0, c1 = spec["first_row"], spec["last_row"], spec["first_col"], spec["last_col"]
    cols = [c for c in range(c0, c1 + 1) if src.Columns.getByIndex(c - 1).IsVisible]
    def text(c, r):
        return src.getCellByPosition(c - 1, r - 1).getString()
    hr = spec.get("header_row")
    headers = [(text(c, hr) if hr else "") or col_letter(c) for c in cols]
    data = src.getCellRangeByPosition(c0 - 1, r0 - 1, c1 - 1, r1 - 1).getDataArray()
    records = []
    for i, row in enumerate(data):
        r = r0 + i
        if not src.Rows.getByIndex(r - 1).IsVisible:
            continue
        vals = []
        for c in cols:
            v = row[c - c0]
            if isinstance(v, str):
                vals.append(v)
            else:  # number/date: the displayed (formatted) text
                vals.append(text(c, r) if v != "" else "")
        if any(x.strip() for x in vals):
            records.append((r, vals))
    key_idx = cols.index(spec["key_col"]) if spec.get("key_col") in cols else None
    rows, bold = assemble_cards(headers, records, key_idx, spec)
    name = unique_name(sheets, spec["sheet_name"])
    sheets.insertNewByName(name, sheets.Count)
    sh = sheets.getByName(name)
    rng = sh.getCellRangeByPosition(0, 0, 1, len(rows) - 1)
    rng.setDataArray(tuple(tuple(r) for r in rows))
    for i, w in enumerate(spec["widths"]):
        sh.Columns.getByIndex(i).Width = width_100mm(w)
    rng.IsTextWrapped = True
    rng.VertJustify = 1
    rng.CharHeight = 10
    rng.CharHeightComplex = 10
    field = sh.getCellRangeByPosition(0, 0, 0, len(rows) - 1)
    field.CharWeight = 150.0
    field.CharWeightComplex = 150.0
    for b in bold:
        t = sh.getCellRangeByPosition(0, b, 1, b)
        t.CellBackColor = 0xDDEBF7
    if spec.get("rtl"):
        sh.TableLayout = 1
    rng.Rows.OptimalHeight = True
    return name


def width_100mm(chars):
    px = int(chars * 7 + 5)
    return int(px * 0.75 * 2540 / 72)


def create_appendix(doc, name, spec):
    """Extra sheet holding the long texts; written as plain strings (never formulas)."""
    sheets = doc.Sheets
    sheets.insertNewByName(name, sheets.Count)
    sh = sheets.getByName(name)
    rows = spec["rows"]
    ncols = len(rows[0])
    rng = sh.getCellRangeByPosition(0, 0, ncols - 1, len(rows) - 1)
    rng.setDataArray(tuple(tuple(str(v) for v in r) for r in rows))
    for i, w in enumerate(spec["widths"]):
        sh.Columns.getByIndex(i).Width = width_100mm(w)
    rng.IsTextWrapped = True
    rng.VertJustify = 1  # TOP
    rng.CharHeight = 10
    rng.CharHeightComplex = 10
    for r in spec.get("bold_rows") or []:
        row = sh.getCellRangeByPosition(0, r, ncols - 1, r)
        row.CharWeight = 150.0
        row.CharWeightComplex = 150.0
    if spec.get("rtl"):
        sh.TableLayout = 1  # RL_TB
    rng.Rows.OptimalHeight = True
    return sh


def in_areas(col, row, areas):
    if not areas:
        return True
    for a in areas:
        if a.StartColumn <= col <= a.EndColumn and a.StartRow <= row <= a.EndRow:
            return True
    return False


def collect_texts(sheet, max_cells):
    flags = 1 | 2 | 4 | 16  # VALUE | DATETIME | STRING | FORMULA
    try:
        ranges = sheet.queryVisibleCells().queryContentCells(flags)
    except Exception:
        ranges = sheet.queryContentCells(flags)
    addrs = ranges.getRangeAddresses()
    areas = list(sheet.getPrintAreas() or [])
    total = sum((a.EndRow - a.StartRow + 1) * (a.EndColumn - a.StartColumn + 1) for a in addrs)
    stride = max(1, int(math.ceil(total / float(max_cells)))) if max_cells else 1
    texts, k = [], 0
    for a in addrs:
        cols = a.EndColumn - a.StartColumn + 1
        size = (a.EndRow - a.StartRow + 1) * cols
        start = (-k) % stride
        for i in range(start, size, stride):
            r, c = a.StartRow + i // cols, a.StartColumn + i % cols
            if not in_areas(c, r, areas):
                continue
            s = sheet.getCellByPosition(c, r).getString()
            if s.strip():
                texts.append([col_letter(c + 1) + str(r + 1), s])
        k += size
    return texts, stride > 1


def export_pdf(doc, out_path):
    fd = uno.Any("[]com.sun.star.beans.PropertyValue", tuple([
        pv("Quality", 92), pv("ReduceImageResolution", False), pv("UseLosslessCompression", False),
        pv("ExportBookmarks", False), pv("ExportNotes", False), pv("EmbedStandardFonts", True),
        pv("IsSkipEmptyPages", True), pv("UseTaggedPDF", True), pv("SelectPdfVersion", 0),
    ]))
    uno.invoke(doc, "storeToURL", (file_url(out_path),
                                   (pv("FilterName", "calc_pdf_Export"), pv("FilterData", fd))))


def main(job_path, result_path):
    job = json.loads(Path(job_path).read_text(encoding="utf-8"))
    result = {"ok": False, "sheets": [], "engine_version": ""}
    proc = doc = desktop = None
    try:
        proc, ctx, desktop = connect(job)
        result["engine_version"] = engine_version(ctx)
        doc = desktop.loadComponentFromURL(file_url(job["input"]), "_blank", 0, (
            pv("Hidden", True), pv("MacroExecutionMode", 0), pv("UpdateDocMode", 0),
            pv("ReadOnly", False)))
        if doc is None:
            raise RuntimeError("engine_open_failed")
        sheets = doc.Sheets
        names = list(sheets.ElementNames)
        out_dir = Path(job["out_dir"])
        for i, spec in enumerate(job["sheets"]):
            r = {"name": spec["name"], "pdf": None, "error": None, "texts": [], "sampled": False,
                 "hash_cells": []}
            try:
                target = spec["name"]
                if spec.get("create_appendix"):
                    create_appendix(doc, spec["name"], spec["create_appendix"])
                    names = list(sheets.ElementNames)
                if (spec["plan"] or {}).get("cards"):
                    target = build_cards(doc, spec["name"], spec["plan"]["cards"])
                    names = list(sheets.ElementNames)
                sheet = sheets.getByName(target)
                idx = names.index(target)
                sheet.IsVisible = True
                for n in names:
                    if n != target:
                        sheets.getByName(n).IsVisible = False
                apply_plan(doc, sheet, idx, spec["plan"])
                out = out_dir / ("sheet_%03d.pdf" % i)
                export_pdf(doc, out)
                if out.exists() and out.stat().st_size > 0:
                    r["pdf"] = str(out)
                else:
                    r["error"] = "nothing_to_print"
                if job.get("collect_text", True):
                    r["texts"], r["sampled"] = collect_texts(sheet, job.get("max_text_cells", 30000))
            except Exception as exc:  # per-sheet failure must not stop other sheets
                r["error"] = "sheet_export_failed"
                sys.stderr.write("sheet %d: %s\n" % (i, type(exc).__name__))
            result["sheets"].append(r)
        result["ok"] = True
    except Exception as exc:
        result["error"] = str(exc) if str(exc).startswith("engine_") else "engine_failed"
        traceback.print_exc()
    finally:
        try:
            if doc is not None:
                doc.close(True)
        except Exception:
            pass
        try:
            if desktop is not None:
                desktop.terminate()
        except Exception:
            pass
        if proc is not None:
            try:
                proc.wait(15)
            except Exception:
                proc.kill()
        Path(result_path).write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
