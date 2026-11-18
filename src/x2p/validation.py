"""Stage 7: verify each exported sheet PDF by *measuring* it.

- text coverage: every displayed cell text reported by the engine must be found in the
  PDF text (whitespace-insensitive; Arabic is also matched reversed, because PDF text
  extraction of RTL runs may come out in visual order).
- smallest font size among body text (header/footer margin bands are excluded).
- '###' overflow markers, blank pages, text outside the page box, missing images/charts.
"""
from __future__ import annotations

import re
import unicodedata
from collections import Counter

import pymupdf

from .layout import MM
from .models import SheetInfo, SheetPlan, SheetVerification

_BIDI = re.compile("[‎‏‪-‮⁦-⁩؜﻿]")
# Whitespace, plus "_": Excel's PDF export draws underscores but writes them to the text
# layer as spaces, so they are not comparable characters.
_WS = re.compile(r"[\s_]+")
_TATWEEL = "ـ"
HASH_RE = re.compile(r"#{3,}")
PAPER_NAMES = {(595, 842): "A4", (842, 1191): "A3", (612, 792): "Letter", (612, 1008): "Legal",
               (420, 595): "A5"}


_LAM_ALEF = re.compile("\u0644([\u0627\u0623\u0625\u0622])")


def normalize(text: str) -> str:
    t = unicodedata.normalize("NFKC", text)
    t = _BIDI.sub("", t).replace(_TATWEEL, "")
    # PDF extraction of the lam-alef ligature (ﻻ) often yields alef+lam; canonicalise
    # both orders to the same key so Arabic words compare equal.
    t = _LAM_ALEF.sub("\\1\u0644", t)
    return _WS.sub("", t)


def _has_rtl(s: str) -> bool:
    return any("֐" <= ch <= "ࣿ" or "ﭐ" <= ch <= "ﻼ" for ch in s)


class LineIndex:
    """Per-line character bags of the PDF text, for RTL runs whose extraction order is
    scrambled (e.g. LibreOffice output of Arabic mixed with digits). A cell matches when
    all of its characters occur in one line, or in a few consecutive lines (wrapped text;
    the window grows with the cell's length)."""

    WINDOW = 3

    def __init__(self, page_texts: list[str]):
        self.lines = [Counter(normalize(ln)) for t in page_texts for ln in t.splitlines()]
        # letters of every extracted word, order-insensitive: Excel's PDF text layer can swap
        # letters inside an Arabic word (e.g. "عصري" extracted as "عرصي") while drawing it right
        # (also adjacent pairs: a word can be split in two, e.g. "كوفي" -> "كوف" + "ي")
        self.word_keys: set[str] = set()
        for t in page_texts:
            toks = [normalize(w) for w in t.split()]
            for i, w in enumerate(toks):
                self.word_keys.add("".join(sorted(w)))
                if i + 1 < len(toks):
                    self.word_keys.add("".join(sorted(w + toks[i + 1])))
        self.by_char: dict[str, set[int]] = {}
        for i, c in enumerate(self.lines):
            for ch in c:
                self.by_char.setdefault(ch, set()).add(i)

    def contains(self, expected: str) -> bool:
        need = Counter(expected)
        try:
            rare = min(need, key=lambda ch: len(self.by_char.get(ch, ())))
        except ValueError:
            return True
        # Long wrapped cells span more lines, and extraction interleaves them with the
        # neighbouring cells' lines; measured on real files: ~1 line per 12-15 characters.
        window = min(60, max(self.WINDOW, len(expected) // 8 + 3))
        for i in sorted(self.by_char.get(rare, ())):
            for lo in range(max(0, i - window + 1), i + 1):
                bag: Counter = Counter()
                for j in range(lo, min(lo + window, len(self.lines))):
                    bag += self.lines[j]
                    if not need - bag:
                        return True
        return False


def _found(expected: str, haystack: str, raw_expected: str, index: LineIndex | None = None) -> bool:
    if not expected:
        return True
    if expected in haystack:
        return True
    if _has_rtl(raw_expected):
        if expected[::-1] in haystack:
            return True
        # word-level fallback: every word present in either direction
        words = [normalize(w) for w in raw_expected.split() if normalize(w)]
        if words and all(w in haystack or w[::-1] in haystack
                         or (index is not None and "".join(sorted(w)) in index.word_keys)
                         for w in words):
            return True
        return index is not None and index.contains(expected)
    return False


def paper_name(rect: pymupdf.Rect) -> str:
    w, h = sorted((round(rect.width), round(rect.height)))
    for (pw, ph), name in PAPER_NAMES.items():
        if abs(w - pw) <= 3 and abs(h - ph) <= 3:
            return name
    return f"{w}x{h}pt"


def verify_sheet_pdf(pdf_path: str, engine_sheet: dict, si: SheetInfo | None, plan: SheetPlan,
                     min_font_pt: float) -> tuple[SheetVerification, Counter]:
    """Returns the verification record and a Counter of measured chars per font size."""
    v = SheetVerification(sheet=plan.sheet, plan=plan.to_dict())
    sizes: Counter = Counter()
    parts: list[str] = []
    with pymupdf.open(pdf_path) as doc:
        v.pages = doc.page_count
        papers = set()
        top_band = plan.margins_mm[0] * MM * 0.9
        bottom_band = plan.margins_mm[1] * MM * 0.9
        for pno, page in enumerate(doc):
            rect = page.rect
            papers.add(paper_name(rect))
            v.orientations.append("landscape" if rect.width > rect.height else "portrait")
            text = page.get_text("text")
            parts.append(text)
            images = len(page.get_images(full=False))
            drawings = len(page.get_drawings())
            v.pdf_images += images
            v.pdf_drawings += drawings
            if not text.strip() and images == 0 and drawings == 0:
                v.blank_pages.append(pno + 1)
            d = page.get_text("dict", flags=pymupdf.TEXTFLAGS_TEXT)
            for block in d.get("blocks", []):
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        t = span.get("text", "").strip()
                        if not t:
                            continue
                        x0, y0, x1, y1 = span["bbox"]
                        if x0 < rect.x0 - 1 or y0 < rect.y0 - 1 or x1 > rect.x1 + 1 or y1 > rect.y1 + 1:
                            v.out_of_page_spans += 1
                        if y1 <= rect.y0 + top_band or y0 >= rect.y1 - bottom_band:
                            continue  # header / footer zone
                        sizes[round(float(span["size"]), 1)] += len(t)
        v.paper = ", ".join(sorted(papers)) if papers else None

    if sizes:
        v.min_font_pt = min(sizes)
    hay = normalize("\n".join(parts))
    texts = engine_sheet.get("texts") or []
    v.sampled = bool(engine_sheet.get("sampled"))
    v.hash_cells = len(engine_sheet.get("hash_cells") or [])
    v.expected_texts = len(texts)
    missing: list[str] = []
    has_rtl = any(_has_rtl(raw) for _, raw in texts)
    index = LineIndex(parts) if has_rtl else None
    page_bags = [Counter(normalize(p)) for p in parts] if has_rtl else []
    for addr, raw in texts:
        e = normalize(raw)
        if _found(e, hay, raw, index):
            v.found_texts += 1
        elif _has_rtl(raw) and any(not (Counter(e) - bag) for bag in page_bags):
            # Every character is on one page but the PDF text layer scrambles RTL order so
            # the sequence cannot be proven: not counted as found, reported for visual review.
            v.rtl_unverified.append(addr)
        else:
            missing.append(addr)
    v.missing_cells = missing[:200]
    if not v.hash_cells:
        # The engine didn't report it (LibreOffice renders ### only at paint time):
        # count ### runs in the PDF that are not part of genuine cell texts.
        in_pdf = len(HASH_RE.findall(hay))
        genuine = sum(len(HASH_RE.findall(normalize(raw))) for _, raw in texts)
        v.hash_cells = max(0, in_pdf - genuine)
    return v, sizes


def chars_below(sizes: Counter, threshold: float) -> tuple[int, int]:
    total = sum(sizes.values())
    below = sum(n for s, n in sizes.items() if s < threshold - 0.05)
    return below, total
