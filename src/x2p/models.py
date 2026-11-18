"""Plain data models shared by every layer (no Qt, no engine imports)."""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from enum import Enum
from typing import Any


class Severity(str, Enum):
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"


class LayoutMode(str, Enum):
    AUTO = "auto"                  # recommended automatic layout
    FIT_WIDTH = "fit_width"        # all columns on one page width
    ONE_PAGE = "one_page"          # whole sheet on one page (warns if text too small)
    KEEP_ORIGINAL = "keep_original"  # respect the workbook's own print setup
    MANUAL = "manual"              # user supplied values
    CARDS = "cards"                # one "field: value" card per record (explicit choice only)


class FinalStatus(str, Enum):
    PASSED = "passed"
    PASSED_WITH_WARNINGS = "passed_with_warnings"
    FAILED = "failed"


@dataclass
class Issue:
    code: str                 # stable machine code, e.g. "text_missing"
    severity: Severity
    sheet: str | None = None
    detail: dict[str, Any] = field(default_factory=dict)  # never contains cell contents

    def to_dict(self) -> dict[str, Any]:
        return {"code": self.code, "severity": self.severity.value, "sheet": self.sheet, "detail": self.detail}


@dataclass
class CellRect:
    """1-based inclusive rectangle."""
    min_row: int
    min_col: int
    max_row: int
    max_col: int

    @property
    def rows(self) -> int:
        return self.max_row - self.min_row + 1

    @property
    def cols(self) -> int:
        return self.max_col - self.min_col + 1

    def union(self, other: "CellRect | None") -> "CellRect":
        if other is None:
            return self
        return CellRect(min(self.min_row, other.min_row), min(self.min_col, other.min_col),
                        max(self.max_row, other.max_row), max(self.max_col, other.max_col))


@dataclass
class PageSetupInfo:
    orientation: str | None = None      # "portrait" | "landscape" | None (default)
    paper: str | None = None            # "A4", "Letter", ...
    scale: int | None = None            # percent
    fit_to_page: bool = False
    fit_width: int | None = None
    fit_height: int | None = None
    print_area: str | None = None
    title_rows: str | None = None       # "1:2"
    title_cols: str | None = None       # "A:B"
    row_breaks: int = 0
    col_breaks: int = 0
    has_header_footer: bool = False


@dataclass
class SheetInfo:
    index: int
    name: str
    kind: str = "worksheet"             # worksheet | chartsheet
    state: str = "visible"              # visible | hidden | veryHidden
    content: CellRect | None = None     # real content bounds (values + drawings), not formatting
    layout_rect: CellRect | None = None  # print area if defined, else content
    value_cells: int = 0
    formula_cells: int = 0
    merged_ranges: int = 0
    images: int = 0
    charts: int = 0
    shapes: int = 0
    tables: int = 0
    conditional_formats: int = 0
    comments: int = 0
    hidden_rows: int = 0
    hidden_cols: int = 0
    right_to_left: bool = False
    base_font_pt: float = 11.0          # most common font size among content cells
    min_font_pt_source: float | None = None  # smallest font size among content cells
    col_widths_pt: list[float] = field(default_factory=list)   # visible columns of content rect
    row_heights_pt: list[float] = field(default_factory=list)  # visible rows of content rect
    uncached_formulas: int = 0          # formulas with no saved result (computed by the engine)
    error_values: int = 0               # cells whose value is an error (#N/A, #DIV/0! ...)
    overflow_cells: int = 0             # unwrapped text cells estimated wider than their column
    overflow: list[tuple[int, int, float]] = field(default_factory=list)  # (row, col, estimated wrapped lines)
    long_texts: dict[tuple[int, int], str] = field(default_factory=dict)  # text of cells that may go to the appendix
    header_texts: dict[int, str] = field(default_factory=dict)            # col -> header caption
    key_texts: dict[int, str] = field(default_factory=dict)               # row -> key column value
    hidden_row_numbers: list[int] = field(default_factory=list)
    visible_cols: list[int] = field(default_factory=list)       # column numbers behind col_widths_pt
    frozen_rows: str | None = None                               # "1:2" from the sheet's frozen panes
    frozen_cols: str | None = None                               # "A:B"
    has_borders: bool = False                                    # any cell border in the first rows
    wb_font_pt: float = 11.0                                     # workbook default font (column width unit)
    text_cells: list[tuple[int, int, int, float, int]] = field(default_factory=list)  # unwrapped text: r, c, width em, pt, longest word em
    value_need_pt: dict[int, float] = field(default_factory=dict)  # col -> width numbers/dates need (no ###)
    widened_cols: dict[int, float] = field(default_factory=dict)   # col -> new width pt (working copy)
    header_row_guess: int | None = None
    header_confidence: float = 0.0
    key_col_guess: int | None = None
    key_confidence: float = 0.0
    page_setup: PageSetupInfo = field(default_factory=PageSetupInfo)

    @property
    def is_empty(self) -> bool:
        return self.content is None and self.charts == 0 and self.images == 0 and self.shapes == 0

    @property
    def content_width_pt(self) -> float:
        return sum(self.col_widths_pt)

    @property
    def content_height_pt(self) -> float:
        return sum(self.row_heights_pt)


@dataclass
class WorkbookInfo:
    path: str
    file_format: str
    size_bytes: int
    sheets: list[SheetInfo] = field(default_factory=list)
    has_macros: bool = False
    external_links: int = 0
    issues: list[Issue] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["issues"] = [i.to_dict() for i in self.issues]
        for s in d["sheets"]:  # keep reports small
            for k in ("col_widths_pt", "row_heights_pt", "hidden_row_numbers", "overflow", "long_texts",
                      "header_texts", "key_texts", "text_cells", "value_need_pt", "visible_cols", "widened_cols"):
                s.pop(k, None)
        return d


@dataclass
class SheetPlan:
    """What the engine must apply to one sheet before export (engine-neutral)."""
    sheet: str
    mode: LayoutMode
    keep_original: bool = False
    paper: str = "A4"
    orientation: str = "portrait"
    scale: int | None = 100              # percent; None when fit_* is used
    fit_width: int | None = None         # pages wide
    fit_height: int | None = None        # pages tall (0 = automatic)
    margins_mm: tuple[float, float, float, float] = (19.0, 19.0, 17.8, 17.8)  # top, bottom, left, right
    title_rows: str | None = None        # "1:1"
    title_cols: str | None = None        # "A:A"
    center_horizontally: bool = False
    wrap_ranges: list[str] = field(default_factory=list)     # e.g. "AY2:AY388" -> wrap text on the working copy
    autofit_rows: list[str] = field(default_factory=list)    # e.g. "2:50" -> optimal row height (visible rows only)
    long_cells: list[tuple[int, int]] = field(default_factory=list)  # cells whose text moves to the appendix
    endnotes: list[tuple[str, str]] = field(default_factory=list)    # (address, marker) set on the working copy
    col_widths: dict[str, float] = field(default_factory=dict)       # column letter -> width (Excel characters), working copy
    col_widths_pt: dict[str, float] = field(default_factory=dict)    # same widths in points (LibreOffice)
    wrapped_cells: int = 0
    print_gridlines: bool = False        # sheet has no borders: print gridlines so rows can be followed
    row_breaks: list[int] = field(default_factory=list)  # manual page break before these rows
    page_order: str = "down"             # "over": across then down (parts of the same rows are consecutive)
    print_headings: bool = False         # print row numbers / column letters (no repeated key column)
    cards: dict | None = None            # card-per-record build instructions (LayoutMode.CARDS)
    expected_pages_across: int = 1
    expected_pages: int = 1
    estimated_min_font_pt: float | None = None
    notes: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["mode"] = self.mode.value
        return d


@dataclass
class JobOptions:
    sheets: list[str] | None = None      # None = all visible sheets
    include_hidden: bool = False
    mode: LayoutMode = LayoutMode.AUTO
    per_sheet_mode: dict[str, LayoutMode] = field(default_factory=dict)
    manual: dict[str, dict[str, Any]] = field(default_factory=dict)  # sheet -> SheetPlan overrides
    paper: str = "A4"
    allow_larger_paper: bool = False
    min_font_pt: float = 8.0
    engine: str = "auto"                 # auto | libreoffice | excel
    timeout_s: int = 300                 # base limit; +1 s per 50 cells (pipeline)
    max_file_mb: int = 200
    max_text_cells: int = 30000          # per sheet, for coverage verification
    allow_failed_output: bool = False    # deliver even if verification failed (explicit)
    wrap_overflow: bool = True           # wrap text that would be clipped (working copy only, reported)
    templates: dict[str, dict] = field(default_factory=dict)  # sheet -> {title_rows, title_cols, page_order, row_numbers}
    gridlines: bool = True               # print gridlines for borderless tables (auto/manual)
    widen_columns: bool = True          # widen too-narrow columns on the working copy (auto/manual)
    long_text: str = "appendix"          # appendix | wrap : text too long to wrap legibly in its cell
    long_text_lines: int = 3             # wrapped-line estimate above which a cell goes to the appendix
    lang: str = "ar"                     # language of generated content (appendix captions)


@dataclass
class SheetVerification:
    sheet: str
    pages: int = 0
    paper: str | None = None
    orientations: list[str] = field(default_factory=list)
    min_font_pt: float | None = None
    expected_texts: int = 0
    found_texts: int = 0
    sampled: bool = False
    hash_cells: int = 0
    missing_cells: list[str] = field(default_factory=list)  # addresses only
    rtl_unverified: list[str] = field(default_factory=list)  # RTL cells: all characters on the page, order not provable
    blank_pages: list[int] = field(default_factory=list)
    out_of_page_spans: int = 0
    pdf_images: int = 0
    pdf_drawings: int = 0
    plan: dict[str, Any] | None = None
    attempt: int = 1

    @property
    def coverage(self) -> float | None:
        return None if self.expected_texts == 0 else self.found_texts / self.expected_texts


@dataclass
class ConversionReport:
    workbook: str
    engine: str = ""
    engine_version: str = ""
    output_pdf: str | None = None
    temp_pdf: str | None = None
    exported_sheets: list[str] = field(default_factory=list)
    excluded_sheets: dict[str, str] = field(default_factory=dict)   # name -> reason code
    total_pages: int = 0
    pdf_size_bytes: int = 0
    sheets: list[SheetVerification] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    status: FinalStatus = FinalStatus.FAILED
    delivered: bool = False
    source_unchanged: bool | None = None
    elapsed_s: float = 0.0
    sheet_page_ranges: dict[str, tuple[int, int]] = field(default_factory=dict)  # 0-based inclusive

    @property
    def min_font_pt(self) -> float | None:
        vals = [s.min_font_pt for s in self.sheets if s.min_font_pt is not None]
        return min(vals) if vals else None

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["status"] = self.status.value
        d["issues"] = [i.to_dict() for i in self.issues]
        d["min_font_pt"] = self.min_font_pt
        for s, sv in zip(d["sheets"], self.sheets):
            s["coverage"] = sv.coverage
        return d
