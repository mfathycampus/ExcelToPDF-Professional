"""Text-width estimates shared by the inspector and the planner (no engine needed).

Estimates are deliberately simple and slightly generous: an over-estimate costs a little width,
an under-estimate is caught by wrapping (text) or by verification (### on numbers).
Both engines break a word longer than its column at character boundaries when wrapping is on
(verified: URLs and codes are not clipped), so long tokens are wrapped like any text.
"""
from __future__ import annotations

import datetime as dt
import math
import re

DIGIT_EM = 0.58       # digits, separators, dates: a touch wider, ### is data loss
CELL_PAD_PT = 6.0     # inner cell padding both sides
NUM_PAD_PT = 9.0      # numbers: padding plus a safety margin against ###


def _char_em(ch: str) -> float:
    """Approximate advance width (em) of one character in Arial/Calibri-like fonts, a little
    generous. Calibrated on engine PDFs (Excel: Arial, LibreOffice: Tahoma for Arabic):
    Arabic ~0.55-0.58 em, Latin and digits within 3 % of these figures."""
    o = ord(ch)
    if ch == " ":
        return 0.28
    if 0x0600 <= o <= 0x06FF or 0x0750 <= o <= 0x077F or 0xFB50 <= o <= 0xFEFF:
        return 0.0 if 0x064B <= o <= 0x065F or o == 0x0670 else 0.60   # harakat take no width
    if "0" <= ch <= "9":
        return 0.56
    if "a" <= ch <= "z":
        return 0.62 if ch in "mw" else 0.30 if ch in "ijlft" else 0.52
    if "A" <= ch <= "Z":
        return 0.85 if ch in "MW" else 0.68
    if o < 0x80:
        return 0.33 if ch in ".,:;'!|()[]-/\\" else 0.58
    if 0x2E80 <= o <= 0x9FFF or 0xAC00 <= o <= 0xD7AF or o >= 0x1F000:
        return 1.0                                                     # CJK, emoji
    return 0.6


def text_em(s: str) -> float:
    """Width of a single-line text in em (multiply by the font size for points)."""
    return sum(_char_em(ch) for ch in s if ch not in "\r\n")


def text_width_pt(em: float, font_pt: float) -> float:
    return em * font_pt


def wrapped_lines(em: float, token_em: float, font_pt: float, col_pt: float) -> float:
    """Lines a text `em` wide (longest word `token_em`) needs in a wrapped column of col_pt points."""
    usable = max(col_pt - CELL_PAD_PT, 1.0)
    lines = text_width_pt(em, font_pt) / usable
    if lines <= 1.0 + 1e-3:
        return lines
    # word wrapping wastes the end of lines: count a little extra for longer texts
    return max(lines * 1.12, text_width_pt(token_em, font_pt) / usable)


_QUOTED = re.compile(r'"[^"]*"|\[[^\]]*\]|\\.')


def display_chars(value, number_format: str | None) -> int:
    """Approximate length of a non-text value as Excel displays it."""
    fmt = _QUOTED.sub("", number_format or "General")
    if isinstance(value, bool):
        return 5
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        if fmt.lower() in ("general", "@"):
            return 16 if isinstance(value, dt.datetime) else 10
        f = fmt.lower()
        f = f.replace("mmmm", "x" * 9).replace("dddd", "x" * 9).replace("am/pm", "xx")
        return max(5, min(30, len(f.split(";")[0])))
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return 6
        body = fmt.split(";")[0]
        if body.lower() == "general":
            s = repr(value) if isinstance(value, float) else str(value)
            return min(len(s), 11)          # General switches to scientific beyond 11 chars
        decimals = len(body.split(".", 1)[1].rstrip("%) _-")) if "." in body else 0
        try:
            s = f"{abs(value):,.{decimals}f}" if "," in body else f"{abs(value):.{decimals}f}"
        except (ValueError, OverflowError):
            s = str(value)
        extra = (1 if value < 0 else 0) + body.count("%") + (2 if any(ch in body for ch in "$€£¥") else 0)
        return len(s) + extra
    return len(str(value))


def number_width_pt(chars: int, font_pt: float) -> float:
    return chars * font_pt * DIGIT_EM + NUM_PAD_PT


def pt_to_col_chars(pt: float, wb_font_pt: float = 11.0) -> float:
    """Inverse of inspector.col_width_to_pt: points -> Excel column width (characters)."""
    mdw = 7.0 * wb_font_pt / 11.0
    return max(0.0, round((pt / 0.75 - 5) / mdw, 2))
