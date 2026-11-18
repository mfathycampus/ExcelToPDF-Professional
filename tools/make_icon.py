"""Draw the application icon with Qt (no extra dependency).

Writes src/x2p/resources/app.png (window icon) and packaging/app.ico (executable icon).
"""
import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtCore import QRectF, Qt  # noqa: E402
from PySide6.QtGui import QColor, QFont, QGuiApplication, QImage, QPainter, QPainterPath, QPen  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]


def draw(size: int) -> QImage:
    img = QImage(size, size, QImage.Format_ARGB32)
    img.fill(Qt.transparent)
    p = QPainter(img)
    p.setRenderHint(QPainter.Antialiasing)
    s = size / 256
    # sheet (green) behind, page (white with red band) in front
    path = QPainterPath()
    path.addRoundedRect(QRectF(16 * s, 28 * s, 150 * s, 190 * s), 18 * s, 18 * s)
    p.fillPath(path, QColor("#1E7145"))
    p.setPen(QPen(QColor(255, 255, 255, 110), 4 * s))
    for i in range(1, 5):
        y = (28 + i * 38) * s
        p.drawLine(int(16 * s), int(y), int(166 * s), int(y))
    p.drawLine(int(66 * s), int(28 * s), int(66 * s), int(218 * s))
    page = QPainterPath()
    page.addRoundedRect(QRectF(92 * s, 60 * s, 148 * s, 180 * s), 16 * s, 16 * s)
    p.fillPath(page, QColor("#FFFFFF"))
    p.setPen(QPen(QColor("#C62828"), 6 * s))
    p.drawPath(page)
    band = QPainterPath()
    band.addRoundedRect(QRectF(92 * s, 172 * s, 148 * s, 52 * s), 10 * s, 10 * s)
    p.fillPath(band, QColor("#C62828"))
    p.setPen(QColor("#FFFFFF"))
    f = QFont("Segoe UI", max(6, int(34 * s)))
    f.setBold(True)
    p.setFont(f)
    p.drawText(QRectF(92 * s, 172 * s, 148 * s, 52 * s), Qt.AlignCenter, "PDF")
    p.setPen(QPen(QColor("#1E7145"), 10 * s, Qt.SolidLine, Qt.RoundCap))
    p.drawLine(int(122 * s), int(116 * s), int(208 * s), int(116 * s))
    p.drawLine(int(186 * s), int(96 * s), int(208 * s), int(116 * s))
    p.drawLine(int(186 * s), int(136 * s), int(208 * s), int(116 * s))
    p.end()
    return img


def main() -> int:
    app = QGuiApplication(sys.argv)  # noqa: F841 - fonts need an app
    res = ROOT / "src" / "x2p" / "resources"
    res.mkdir(parents=True, exist_ok=True)
    draw(256).save(str(res / "app.png"))
    (ROOT / "packaging").mkdir(exist_ok=True)
    ok = draw(256).save(str(ROOT / "packaging" / "app.ico"), "ICO")
    print("png:", res / "app.png", "ico:", ok)
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
