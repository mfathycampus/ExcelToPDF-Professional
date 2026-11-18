"""Engine-free tests: file safety, inspection, planning, normalisation, translations."""
import re
import zipfile
from pathlib import Path

import pytest

from x2p.formats import OLE_MAGIC, UnsafeFileError, check_file
from x2p.i18n import STRINGS, Translator
from x2p.inspector import inspect_workbook
from x2p.layout import min_scale_for, plan_sheet
from x2p.models import JobOptions, LayoutMode
from x2p.pipeline import select_sheets
from x2p.validation import _found, normalize

SRC = Path(__file__).resolve().parents[2] / "src" / "x2p"


# ---------------------------------------------------------------- stage 1: file safety
def test_unsupported_and_mismatched_files(tmp_path, fixtures):
    p = tmp_path / "a.xlsb"
    p.write_bytes(b"PK\x03\x04" + b"0" * 20)
    with pytest.raises(UnsafeFileError) as e:
        check_file(p)
    assert e.value.code == "format_unsupported"

    fake = tmp_path / "fake.xls"
    fake.write_bytes(fixtures["01_simple"].read_bytes())
    with pytest.raises(UnsafeFileError) as e:
        check_file(fake)
    assert e.value.code == "signature_mismatch"

    enc = tmp_path / "enc.xlsx"
    enc.write_bytes(OLE_MAGIC + b"\0" * 100)
    with pytest.raises(UnsafeFileError) as e:
        check_file(enc)
    assert e.value.code == "password_protected"

    empty = tmp_path / "e.xlsx"
    empty.write_bytes(b"")
    with pytest.raises(UnsafeFileError) as e:
        check_file(empty)
    assert e.value.code == "file_empty"


def test_zip_without_workbook_is_corrupt(tmp_path):
    p = tmp_path / "x.xlsx"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
    with pytest.raises(UnsafeFileError) as e:
        check_file(p)
    assert e.value.code == "corrupt_file"


# ---------------------------------------------------------------- stage 2: inspection
def test_inspect_multi_sheet(fixtures):
    info = inspect_workbook(fixtures["03_multi_sheet"])
    assert [s.name for s in info.sheets] == ["Summary", "Hidden data", "Print area", "Empty"]
    assert info.sheets[1].state == "hidden"
    pa = info.sheets[2]
    assert pa.page_setup.print_area == "A1:D15"
    assert (pa.layout_rect.max_row, pa.layout_rect.max_col) == (15, 4)
    assert (pa.content.max_row, pa.content.max_col) == (30, 8)
    assert info.sheets[3].is_empty
    chosen, excluded = select_sheets(info, JobOptions())
    assert [s.name for s in chosen] == ["Summary", "Print area"]
    assert excluded == {"Hidden data": "hidden", "Empty": "empty"}


def test_inspect_objects_and_hidden(fixtures):
    assert inspect_workbook(fixtures["08_chart"]).sheets[0].charts == 1
    assert inspect_workbook(fixtures["07_merged"]).sheets[0].merged_ranges == 4
    h = inspect_workbook(fixtures["12_hidden_rows_cols"]).sheets[0]
    assert (h.hidden_rows, h.hidden_cols) == (1, 1)
    ar = inspect_workbook(fixtures["05_arabic_rtl"]).sheets[0]
    assert ar.right_to_left and ar.name == "بيانات"
    fi = inspect_workbook(fixtures["09_formulas_cf"])
    f = fi.sheets[0]
    assert f.formula_cells == 40 and f.conditional_formats == 1
    codes = {i.code for i in fi.issues}
    assert {"formulas_computed_by_engine", "error_values_present", "comments_not_printed"} <= codes
    assert f.uncached_formulas == 40 and f.error_values == 1 and f.comments == 1


def test_inspect_xls(fixtures):
    info = inspect_workbook(fixtures["02_legacy"])
    assert info.file_format == "xls"
    assert [s.name for s in info.sheets] == ["Legacy", "عربي"]
    assert info.sheets[0].content.max_row == 21


def test_inspection_never_modifies_source(fixtures):
    p = fixtures["10_numbers"]
    before = (p.read_bytes(), p.stat().st_mtime_ns)
    inspect_workbook(p)
    assert (p.read_bytes(), p.stat().st_mtime_ns) == before


def test_header_and_key_detection(fixtures):
    w = inspect_workbook(fixtures["06_wide_long"]).sheets[0]
    assert w.header_row_guess == 1 and w.header_confidence >= 0.7
    assert w.key_col_guess == 1 and w.key_confidence >= 0.7


# ---------------------------------------------------------------- stages 3-4: planning
def test_small_sheet_keeps_100_percent(fixtures):
    si = inspect_workbook(fixtures["01_simple"]).sheets[0]
    plan = plan_sheet(si, JobOptions())
    assert plan.scale == 100 and plan.expected_pages_across == 1


def test_wide_sheet_split_never_below_min_font(fixtures):
    si = inspect_workbook(fixtures["06_wide_long"]).sheets[0]
    opts = JobOptions(min_font_pt=8)
    plan = plan_sheet(si, opts)
    assert plan.expected_pages_across > 1
    assert plan.scale is not None and plan.scale / 100 * si.base_font_pt >= 8 - 1e-6
    assert plan.title_cols == "A:A" and plan.title_rows == "1:1"
    assert "horizontal_split" in plan.notes


def test_one_page_warns_when_unreadable(fixtures):
    si = inspect_workbook(fixtures["06_wide_long"]).sheets[0]
    plan = plan_sheet(si, JobOptions(mode=LayoutMode.ONE_PAGE))
    assert (plan.fit_width, plan.fit_height) == (1, 1)
    assert "one_page_text_below_minimum" in plan.notes


def test_min_scale_does_not_shrink_small_sources(fixtures):
    si = inspect_workbook(fixtures["01_simple"]).sheets[0]
    si.min_font_pt_source = 7
    assert min_scale_for(si, 8) == 1.0


def test_manual_overrides(fixtures):
    si = inspect_workbook(fixtures["01_simple"]).sheets[0]
    plan = plan_sheet(si, JobOptions(manual={"Items": {"orientation": "landscape", "scale": 90}}))
    assert plan.mode == LayoutMode.MANUAL and plan.orientation == "landscape" and plan.scale == 90


# ---------------------------------------------------------------- stage 7 helpers
def test_arabic_normalisation():
    assert normalize("الاسم") == normalize("ااسم".replace("ااسم", "اال" + "سم"))
    assert _found(normalize("ملاحظات"), normalize("مالحظات"), "ملاحظات")
    assert _found(normalize("مرحبا بكم"), normalize("مكب ابحرم"), "مرحبا بكم")  # visual order
    assert normalize("a‏b  c") == "abc"


# ---------------------------------------------------------------- i18n
def test_every_issue_code_is_translated():
    codes = set()
    for f in SRC.rglob("*.py"):
        text = f.read_text(encoding="utf-8")
        codes |= set(re.findall(r'Issue\(\s*"([a-z_]+)"', text))
        codes |= set(re.findall(r'(?:UnsafeFileError|EngineError)\(\s*"([a-z_]+)"', text))
        codes |= set(re.findall(r'notes(?:\.append\(| \+ \[)"([a-z_]+)"', text))
    missing = sorted(c for c in codes if c not in STRINGS or not {"ar", "en"} <= STRINGS[c].keys())
    assert not missing, missing
    assert Translator("ar")("passed") == "اجتاز التحقق"


def test_rtl_line_index_handles_scrambled_extraction():
    from x2p.validation import LineIndex
    # LibreOffice output: "عميل تجريبي 1" extracted as " ب يعميل تجري1"
    idx = LineIndex([" ب يعميل تجري1\n3,001.00 ر.س"])
    assert _found(normalize("عميل تجريبي 1"), normalize(" ب يعميل تجري1"), "عميل تجريبي 1", idx)
    # but a character that is really missing from that line is still reported
    assert not idx.contains(normalize("عميل تجريبي 19"))
    assert not idx.contains(normalize("عميل تجريبية 1"))


def test_rtl_word_letter_order_and_split_words():
    from x2p.validation import LineIndex
    pdf = "مقهى عرصي جديد | كوف ي هود"          # letters swapped / word split by the text layer
    idx = LineIndex([pdf])
    hay = normalize(pdf)
    assert _found(normalize("مقهى عصري جديد"), hay, "مقهى عصري جديد", idx)
    assert _found(normalize("كوفي هود"), hay, "كوفي هود", idx)
    assert not _found(normalize("مقهى عصري قديم"), hay, "مقهى عصري قديم", idx)  # real difference


def test_underscore_is_ignored_by_matching():
    assert normalize("a_b c") == normalize("a b c") == "abc"


def test_overflow_planning_and_lossless_appendix(fixtures):
    from x2p.appendix import build_appendix, split_text
    si = inspect_workbook(fixtures["13_long_text_appendix"]).sheets[0]
    plan = plan_sheet(si, JobOptions())
    assert plan.long_cells and plan.wrap_ranges and "long_text_to_appendix" in plan.notes
    # short URL/JSON: both engines break long words, so they stay in their (wider) cells
    assert (2, 3) not in plan.long_cells and (2, 4) not in plan.long_cells
    assert "columns_widened" in plan.notes and "C" in plan.col_widths
    built = build_appendix([(si, plan)], [si.name], "ar", "A4")
    assert built is not None
    a_plan, rows, bold, rtl = built
    assert rtl and a_plan.fit_width == 1 and a_plan.sheet != si.name
    assert [m for _, m in plan.endnotes] == [f"[{i}]" for i in range(1, len(plan.long_cells) + 1)]
    # every long text is reproduced exactly by concatenating its chunks
    for k, (r, c) in enumerate(plan.long_cells):
        start = bold[k + 1] + 1
        end = bold[k + 2] if k + 2 < len(bold) else len(rows)
        assert "".join(x[0] for x in rows[start:end]) == si.long_texts[(r, c)]
    text = "x" * 2000 + " " + "y" * 50
    assert "".join(split_text(text)) == text


def test_wrap_overflow_can_be_disabled(fixtures):
    si = inspect_workbook(fixtures["13_long_text_appendix"]).sheets[0]
    plan = plan_sheet(si, JobOptions(wrap_overflow=False))
    assert not plan.long_cells and not plan.wrap_ranges


def test_card_mode_plan(fixtures):
    si = inspect_workbook(fixtures["06_wide_long"]).sheets[0]
    assert "cards_suggested" not in plan_sheet(si, JobOptions()).notes   # 3 parts: fine as a table
    import copy
    wider = copy.deepcopy(si)
    wider.col_widths_pt = [w * 3 for w in si.col_widths_pt]
    auto = plan_sheet(wider, JobOptions())
    assert auto.expected_pages_across >= 4 and "cards_suggested" in auto.notes
    plan = plan_sheet(si, JobOptions(mode=LayoutMode.CARDS))
    c = plan.cards
    assert plan.mode == LayoutMode.CARDS and (plan.fit_width, plan.orientation) == (1, "portrait")
    assert (c["header_row"], c["first_row"], c["key_col"]) == (1, 2, 1)
    assert (c["first_col"], c["last_col"], c["last_row"]) == (1, 40, 301)
    assert len(c["sheet_name"]) <= 31 and not plan.long_cells and not plan.wrap_ranges


def test_column_widening_prevents_hash_and_appendix_for_short_text(tmp_path):
    """System exports keep every column at the default width: dates must not show ###, short
    texts must not go to the appendix, header words must fit, and the fewest pages win."""
    import datetime as dt
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.append(["ID", "Percentage", "Details", "End Date", "Logo"])
    for i in range(60):
        ws.append([i, 50, "تفاصيل العرض " * 4, dt.datetime(2026, 9, 27, 23, 59),
                   f"https://cdn.example.com/storage/places/logos/{i:04d}_abcdefghijklmnop.png"])
        ws.cell(ws.max_row, 4).number_format = "m/d/yyyy h:mm"
    p = tmp_path / "dump.xlsx"
    wb.save(p)
    si = inspect_workbook(p).sheets[0]
    assert all(w == si.col_widths_pt[0] for w in si.col_widths_pt) and not si.has_borders
    plan = plan_sheet(si, JobOptions())
    assert not plan.long_cells                         # nothing short goes to the appendix
    assert "D" in plan.col_widths                      # the date column grows: no ###
    assert plan.print_gridlines and "gridlines_added" in plan.notes
    # the chosen widths are never worse (in pages) than keeping the original widths
    plain = plan_sheet(si, JobOptions(widen_columns=False))
    assert plan.expected_pages <= plain.expected_pages + max(1, len(plain.long_cells) // 30)
    assert plan_sheet(si, JobOptions(mode=LayoutMode.KEEP_ORIGINAL)).col_widths == {}


def test_manual_page_breaks_gridlines_and_widening(tmp_path):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "data"
    ws.append(["ID", "Name", "Notes"])
    for i in range(1, 51):
        ws.append([i, f"name {i}", "a fairly long note that does not fit the default column " * 2])
    ws.row_dimensions[10].hidden = True
    p = tmp_path / "m.xlsx"
    wb.save(p)
    si = inspect_workbook(p).sheets[0]
    auto = plan_sheet(si, JobOptions())
    assert auto.print_gridlines and auto.col_widths and not auto.row_breaks
    manual = {"data": {"break_every": 20, "print_gridlines": False, "widen_columns": False}}
    plan = plan_sheet(si, JobOptions(mode=LayoutMode.MANUAL, manual=manual))
    assert not plan.print_gridlines and "gridlines_added" not in plan.notes
    assert plan.col_widths == {}
    # 20 visible data rows per page: 2-9 and 11-22 (row 10 is hidden) -> break before 23
    assert plan.row_breaks == [23, 43] and "manual_page_breaks" in plan.notes


def _wide_book(tmp_path, freeze=None, key=True):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "wide"
    ws.append((["ID"] if key else []) + [f"Field {i}" for i in range(1, 31)])
    for r in range(1, 41):
        ws.append(([r] if key else []) + [f"value {r}-{i} text" for i in range(1, 31)])
    if freeze:
        ws.freeze_panes = freeze
    p = tmp_path / f"w_{freeze}_{key}.xlsx"
    wb.save(p)
    return inspect_workbook(p).sheets[0]


def test_frozen_panes_define_repeated_titles(tmp_path):
    from x2p.layout import suggested_titles
    si = _wide_book(tmp_path, freeze="C3")
    assert (si.frozen_rows, si.frozen_cols) == ("1:2", "A:B")
    assert suggested_titles(si, JobOptions()) == ("1:2", "A:B")
    plan = plan_sheet(si, JobOptions())
    assert plan.expected_pages_across > 1
    assert plan.title_rows == "1:2" and plan.title_cols == "A:B"
    assert plan.page_order == "over"                     # parts of the same rows come together
    assert not plan.print_headings                        # a column repeats: no row numbers needed


def test_template_overrides_and_row_numbers(tmp_path):
    si = _wide_book(tmp_path, key=False)
    plan = plan_sheet(si, JobOptions(templates={"wide": {"title_cols": "", "page_order": "down"}}))
    assert plan.expected_pages_across > 1 and plan.title_cols is None
    assert plan.page_order == "down"
    assert plan.print_headings and "row_numbers_added" in plan.notes   # ties parts together
    off = plan_sheet(si, JobOptions(templates={"wide": {"title_cols": "", "row_numbers": False}}))
    assert not off.print_headings
    keyed = plan_sheet(si, JobOptions(templates={"wide": {"title_cols": "A:A", "title_rows": "1:1"}}))
    assert keyed.title_cols == "A:A" and keyed.title_rows == "1:1" and not keyed.print_headings
