"""Generate the synthetic reference workbooks of spec §12.1 (no personal data).

Usage: python tools/make_fixtures.py [out_dir]   (default: tests/fixtures/generated)
"""
from __future__ import annotations

import datetime as dt
import random
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BOLD = Font(bold=True)
FILL = PatternFill("solid", fgColor="DDEBF7")
rnd = random.Random(42)


def header(ws, row, names):
    for c, n in enumerate(names, 1):
        cell = ws.cell(row, c, n)
        cell.font, cell.fill = BOLD, FILL


def simple(p):
    wb = Workbook()
    ws = wb.active
    ws.title = "Items"
    header(ws, 1, ["Code", "Item", "Qty", "Price", "Total"])
    for r in range(2, 22):
        ws.append([f"IT-{r:04d}", f"Item number {r}", r * 3, round(r * 1.75, 2), f"=C{r}*D{r}"])
        ws.cell(r, 5).value = (r * 3) * round(r * 1.75, 2)  # literal value (no recalculation needed)
    for col, w in zip("ABCDE", (12, 24, 8, 10, 12)):
        ws.column_dimensions[col].width = w
    wb.save(p)


def multi_sheet(p):
    wb = Workbook()
    a = wb.active
    a.title = "Summary"
    a["A1"], a["B1"] = "Region", "Sales"
    for i, reg in enumerate(["North", "South", "East", "West"], 2):
        a.cell(i, 1, reg)
        a.cell(i, 2, 1000 * i)
    b = wb.create_sheet("Hidden data")
    b["A1"] = "secret-free hidden sheet"
    b.sheet_state = "hidden"
    c = wb.create_sheet("Print area")
    for r in range(1, 31):
        for col in range(1, 9):
            c.cell(r, col, f"R{r}C{col}")
    c.print_area = "A1:D15"
    wb.create_sheet("Empty")
    wb.save(p)


def arabic(p):
    wb = Workbook()
    ws = wb.active
    ws.title = "بيانات"
    ws.sheet_view.rightToLeft = True
    header(ws, 1, ["الرقم", "الاسم", "المدينة", "المبلغ", "ملاحظات"])
    cities = ["الرياض", "جدة", "الدمام", "مكة المكرمة", "المدينة المنورة"]
    for r in range(2, 32):
        ws.append([r - 1, f"عميل تجريبي {r - 1}", cities[r % 5], 1500.5 * r,
                   "Invoice INV-%03d مدفوعة" % r])
        ws.cell(r, 4).number_format = '#,##0.00 "ر.س"'
    for col, w in zip("ABCDE", (8, 22, 18, 16, 30)):
        ws.column_dimensions[col].width = w
    wb.save(p)


def wide_long(p):
    wb = Workbook()
    ws = wb.active
    ws.title = "Wide"
    cols = 40
    header(ws, 1, ["ID"] + [f"Metric {i}" for i in range(1, cols)])
    for r in range(2, 302):
        ws.append([f"ID{r:05d}"] + [rnd.randint(100, 99999) for _ in range(cols - 1)])
    for c in range(1, cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 11
    wb.save(p)


def merged(p):
    wb = Workbook()
    ws = wb.active
    ws.title = "Merged"
    ws["A1"] = "Quarterly report"
    ws.merge_cells("A1:G1")
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center")
    ws["A2"], ws["B2"], ws["E2"] = "Branch", "H1", "H2"
    ws.merge_cells("A2:A3")
    ws.merge_cells("B2:D2")
    ws.merge_cells("E2:G2")
    for c, q in enumerate(["Q1", "Q2", "Sum", "Q3", "Q4", "Sum"], 2):
        ws.cell(3, c, q).font = BOLD
    for r in range(4, 14):
        ws.cell(r, 1, f"Branch {r - 3}")
        for c in range(2, 8):
            ws.cell(r, c, rnd.randint(10, 500))
    wb.save(p)


def chart(p):
    wb = Workbook()
    ws = wb.active
    ws.title = "Chart"
    header(ws, 1, ["Month", "Value"])
    for i, m in enumerate(["Jan", "Feb", "Mar", "Apr", "May", "Jun"], 2):
        ws.cell(i, 1, m)
        ws.cell(i, 2, i * 10)
    ch = BarChart()
    ch.title = "Monthly values"
    ch.add_data(Reference(ws, min_col=2, min_row=1, max_row=7), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=7))
    ws.add_chart(ch, "D2")
    wb.save(p)


def formulas_cf(p):
    wb = Workbook()
    ws = wb.active
    ws.title = "Formulas"
    header(ws, 1, ["A", "B", "Sum", "Flag"])
    for r in range(2, 22):
        ws.cell(r, 1, r)
        ws.cell(r, 2, r * 2)
        ws.cell(r, 3, f"=A{r}+B{r}")
        ws.cell(r, 4, f'=IF(C{r}>30,"HIGH","low")')
    ws.conditional_formatting.add("C2:C21", CellIsRule(operator="greaterThan", formula=["30"],
                                                       fill=PatternFill("solid", fgColor="FFC7CE")))
    ws["E1"] = "Lookup"
    ws["E2"] = "#N/A"                      # literal error value: data, printed as is
    ws["A1"].comment = Comment("Synthetic note", "x2p")
    wb.save(p)


def numbers(p):
    wb = Workbook()
    ws = wb.active
    ws.title = "Numbers"
    header(ws, 1, ["National-like ID (text)", "Zip (00000)", "Date", "Currency", "Percent", "Big"])
    for r in range(2, 12):
        ws.cell(r, 1, f"{1000000000 + r * 7919}")          # long id stored as text
        z = ws.cell(r, 2, r * 7)
        z.number_format = "00000"                          # leading zeros must survive
        d = ws.cell(r, 3, dt.date(2026, 1, r))
        d.number_format = "yyyy-mm-dd"
        m = ws.cell(r, 4, r * 1234.5)
        m.number_format = '"$"#,##0.00'
        q = ws.cell(r, 5, r / 100)
        q.number_format = "0.0%"
        ws.cell(r, 6, 123456789012 + r)
        ws.cell(r, 6).number_format = "0"
    # widths leave headroom: LibreOffice's metric-compatible substitute fonts are a bit wider
    for col, w in zip("ABCDEF", (24, 13, 12, 14, 9, 16)):
        ws.column_dimensions[col].width = w
    wb.save(p)


def long_text(p):
    wb = Workbook()
    ws = wb.active
    ws.title = "Notes"
    header(ws, 1, ["#", "Note"])
    ws.column_dimensions["B"].width = 50
    for r in range(2, 12):
        c = ws.cell(r, 2, ("This is a long multi line note number %d. " % r) * 4)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 1, r - 1)
        ws.row_dimensions[r].height = 60
    wb.save(p)


def hidden_rows_cols(p):
    wb = Workbook()
    ws = wb.active
    ws.title = "Hidden"
    header(ws, 1, ["Visible A", "Hidden B", "Visible C"])
    for r in range(2, 12):
        ws.append([f"va{r}", f"hb{r}", f"vc{r}"])
    ws.column_dimensions["B"].hidden = True
    ws.row_dimensions[5].hidden = True
    wb.save(p)


def narrow_hash(p):
    """Negative fixture: a number too wide for its column must be reported as ###."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Narrow"
    ws["A1"] = 123456789.123
    ws["A1"].number_format = "#,##0.000"
    ws.column_dimensions["A"].width = 4
    wb.save(p)


def clipped_text(p):
    """Negative fixture: text cut off by a non-empty neighbour cell must be reported."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Clipped"
    ws.column_dimensions["A"].width = 6
    for r in range(1, 6):
        ws.cell(r, 1, f"A long label that cannot fit {r}")
        ws.cell(r, 2, r)
    wb.save(p)


def long_text_appendix(p):
    """Very long text, a URL and a JSON blob in narrow columns next to filled cells: must be
    wrapped or moved to the appendix, never clipped."""
    wb = Workbook()
    ws = wb.active
    ws.title = "طويل"
    ws.sheet_view.rightToLeft = True
    header(ws, 1, ["id", "الوصف", "الرابط", "بيانات", "التقييم"])
    for col, w in zip("ABCDE", (8, 14, 14, 14, 10)):
        ws.column_dimensions[col].width = w
    para = "مقهى مختص يقدم القهوة العربية والحلويات مع جلسات داخلية وخارجية هادئة ومناسبة للعائلات. "
    for r in range(2, 8):
        ws.cell(r, 1, r - 1)
        ws.cell(r, 2, para * (r * 3) + f"نهاية {r}")                    # long → appendix
        ws.cell(r, 3, f"https://example.com/places/{r:04d}/menu_images/item_{r}_large.jpg")  # unbreakable
        ws.cell(r, 4, '{"open":"08:00","close":"23:00","days":[1,2,3,4,5]}')  # unbreakable JSON
        ws.cell(r, 5, 4.5)
    ws.cell(8, 1, 7)
    ws.cell(8, 2, "نص متوسط يحتاج إلى سطرين")                              # short → wrapped in place
    ws.cell(8, 5, 4.0)
    wb.save(p)


def legacy_xls(p):
    import xlwt
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("Legacy")
    ws.col(1).width = 256 * 20
    for c, n in enumerate(["Code", "Name", "Value"]):
        ws.write(0, c, n)
    for r in range(1, 21):
        ws.write(r, 0, f"L{r:03d}")
        ws.write(r, 1, f"Legacy row {r}")
        ws.write(r, 2, r * 11)
    ws2 = wb.add_sheet("عربي")
    ws2.write(0, 0, "مرحبا")
    ws2.write(0, 1, 42)
    wb.save(str(p))


FIXTURES = {
    "01_simple.xlsx": simple, "02_legacy.xls": legacy_xls, "03_multi_sheet.xlsx": multi_sheet,
    "05_arabic_rtl.xlsx": arabic, "06_wide_long.xlsx": wide_long, "07_merged.xlsx": merged,
    "08_chart.xlsx": chart, "09_formulas_cf.xlsx": formulas_cf, "10_numbers.xlsx": numbers,
    "11_long_text.xlsx": long_text, "12_hidden_rows_cols.xlsx": hidden_rows_cols,
    "13_long_text_appendix.xlsx": long_text_appendix,
    "90_narrow_hash.xlsx": narrow_hash, "91_clipped_text.xlsx": clipped_text,
}


def build(out: Path) -> list[Path]:
    out.mkdir(parents=True, exist_ok=True)
    paths = []
    for name, fn in FIXTURES.items():
        p = out / name
        fn(p)
        paths.append(p)
    return paths


if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "generated"
    for p in build(target):
        print(p)
