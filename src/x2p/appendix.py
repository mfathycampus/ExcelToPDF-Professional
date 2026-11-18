"""Long-text appendix.

Cells whose text is far too long to wrap legibly in their column (or that contain a word/URL
wider than the column, which cannot wrap at all) are printed in an extra sheet that the
*engine* renders: same fonts, real Unicode text (searchable/copyable), verified like every
other sheet. In the working copy the original cell shows a reference marker "[n]".
Nothing is lost: the appendix rows concatenate back to the exact original text.

Layout: a single wide column. Each entry is a bold reference line
"[n] · sheet · cell · column caption · key" followed by the text in chunks. One column
means no neighbouring cell can interleave with the text during PDF text extraction.
"""
from __future__ import annotations

from .i18n import Translator
from .layout import col_letter
from .models import LayoutMode, SheetInfo, SheetPlan

CHUNK_CHARS = 900               # keeps every appendix row well under the 409 pt row-height limit
APPENDIX_COL_WIDTHS = [125]     # Excel width units (fits A4 landscape at 100 %)


def split_text(text: str, size: int = CHUNK_CHARS) -> list[str]:
    """Split at whitespace near `size`; never drops or alters characters."""
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


def build_appendix(items: list[tuple[SheetInfo, SheetPlan]], existing: list[str], lang: str,
                   paper: str) -> tuple[SheetPlan, list[list[str]], list[int], bool] | None:
    """Numbers the long cells across the workbook (sets plan.endnotes) and returns the
    appendix plan, its rows, the 0-based rows to print bold, and whether it is right-to-left."""
    t = Translator(lang)
    rows: list[list[str]] = [[t("appendix_title")]]
    bold = [0]
    n, rtl = 0, False
    for si, plan in items:
        plan.endnotes = []
        rtl |= si.right_to_left
        for r, c in plan.long_cells:
            n += 1
            marker = f"[{n}]"
            addr = f"{col_letter(c)}{r}"
            plan.endnotes.append((addr, marker))
            meta = [marker, si.name, addr, si.header_texts.get(c, ""), si.key_texts.get(r, "")]
            bold.append(len(rows))
            rows.append(["  ·  ".join(x for x in meta if x)])
            # engines write these cells as text, never as formulas or numbers
            rows.extend([chunk] for chunk in split_text(si.long_texts[(r, c)]))
    if n == 0:
        return None
    name = t("appendix_title")[:31]
    base, k = name, 2
    while name in existing:
        name = f"{base[:28]} {k}"
        k += 1
    plan = SheetPlan(sheet=name, mode=LayoutMode.FIT_WIDTH, paper=paper, orientation="landscape",
                     scale=None, fit_width=1, fit_height=0)
    return plan, rows, bold, rtl
