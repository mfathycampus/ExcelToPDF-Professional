"""Per-sheet manual print settings (spec §5.4 "manual"): paper, orientation, scale or
fit-to-width, margins, repeated title rows/columns. Starts from the automatic plan."""
from __future__ import annotations

import re

from PySide6.QtWidgets import (QCheckBox, QComboBox, QDialog, QDialogButtonBox, QDoubleSpinBox,
                               QFormLayout, QHBoxLayout, QLabel, QLineEdit, QSpinBox, QVBoxLayout,
                               QWidget)

from ..i18n import Translator
from ..models import SheetPlan

ROWS_RE = re.compile(r"^\d+:\d+$")
COLS_RE = re.compile(r"^[A-Za-z]{1,3}:[A-Za-z]{1,3}$")


class ManualDialog(QDialog):
    def __init__(self, tr: Translator, sheet: str, base: SheetPlan, current: dict | None, parent=None):
        super().__init__(parent)
        self.tr_ = tr
        self.setWindowTitle(f"{tr('manual')} — {sheet}")
        v = current or {}
        lay = QVBoxLayout(self)
        form = QFormLayout()
        self.paper = QComboBox()
        self.paper.addItems(["A4", "A3", "Letter", "Legal"])
        self.paper.setCurrentText(v.get("paper", base.paper))
        self.orient = QComboBox()
        self.orient.addItem(tr("portrait"), "portrait")
        self.orient.addItem(tr("landscape"), "landscape")
        self.orient.setCurrentIndex(1 if v.get("orientation", base.orientation) == "landscape" else 0)
        self.fit = QCheckBox(tr("fit_width"))
        self.scale = QSpinBox(minimum=10, maximum=400, suffix=" %")
        scale = v.get("scale", base.scale)
        self.fit.setChecked(scale is None)
        self.scale.setValue(int(scale or 100))
        self.scale.setEnabled(scale is not None)
        self.fit.toggled.connect(lambda on: self.scale.setEnabled(not on))
        scale_row = QWidget()
        h = QHBoxLayout(scale_row)
        h.setContentsMargins(0, 0, 0, 0)
        h.addWidget(self.scale)
        h.addWidget(self.fit)
        margins = v.get("margins_mm", base.margins_mm)
        self.margins = []
        m_row = QWidget()
        mh = QHBoxLayout(m_row)
        mh.setContentsMargins(0, 0, 0, 0)
        for key, val in zip(("top", "bottom", "left", "right"), margins):
            sp = QDoubleSpinBox(minimum=0, maximum=60, decimals=1, value=float(val), suffix=" mm")
            sp.setToolTip(tr(key))
            sp.setAccessibleName(tr(key))
            self.margins.append(sp)
            mh.addWidget(QLabel(tr(key)))
            mh.addWidget(sp)
        self.rows = QLineEdit(v.get("title_rows", base.title_rows) or "")
        self.rows.setPlaceholderText("1:1")
        self.cols = QLineEdit(v.get("title_cols", base.title_cols) or "")
        self.cols.setPlaceholderText("A:A")
        form.addRow(tr("paper"), self.paper)
        form.addRow(tr("orientation"), self.orient)
        form.addRow(tr("scale"), scale_row)
        form.addRow(tr("margins"), m_row)
        form.addRow(tr("title_rows"), self.rows)
        form.addRow(tr("title_cols"), self.cols)
        lay.addLayout(form)
        self.error = QLabel()
        self.error.setStyleSheet("color: palette(link);")
        lay.addWidget(self.error)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def _accept(self):
        rows, cols = self.rows.text().strip(), self.cols.text().strip().upper()
        if rows and not ROWS_RE.match(rows):
            self.error.setText("⚠ " + self.tr_("title_rows") + ": 1:2")
            return
        if cols and not COLS_RE.match(cols):
            self.error.setText("⚠ " + self.tr_("title_cols") + ": A:B")
            return
        self.accept()

    def values(self) -> dict:
        fit = self.fit.isChecked()
        return {
            "paper": self.paper.currentText(),
            "orientation": self.orient.currentData(),
            "scale": None if fit else self.scale.value(),
            "fit_width": 1 if fit else None,
            "fit_height": 0 if fit else None,
            "margins_mm": tuple(sp.value() for sp in self.margins),
            "title_rows": self.rows.text().strip() or None,
            "title_cols": self.cols.text().strip().upper() or None,
        }
