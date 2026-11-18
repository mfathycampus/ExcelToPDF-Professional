"""Offscreen GUI smoke test: open a workbook, inspect, preview, save screenshots.

Usage: python tools/gui_smoke.py <workbook> <out_dir> [ar|en]
"""
import os
import sys
import time
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtWidgets import QApplication  # noqa: E402

from x2p.gui.main_window import MainWindow  # noqa: E402


def wait(app, cond, timeout=300):
    t0 = time.time()
    while not cond():
        app.processEvents()
        time.sleep(0.05)
        if time.time() - t0 > timeout:
            raise TimeoutError
    app.processEvents()


def main():
    wb, out = sys.argv[1], Path(sys.argv[2])
    out.mkdir(parents=True, exist_ok=True)
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    w = MainWindow()
    w.tr_.lang = sys.argv[3] if len(sys.argv) > 3 else "ar"
    w._retranslate()
    w.resize(1400, 900)
    w.show()
    w.open_file(wb)
    wait(app, lambda: w.info is not None)
    w.grab().save(str(out / "1_inspected.png"))
    w.preview()
    wait(app, lambda: w.report is not None and not w.runner.busy)
    w.fit_width()
    app.processEvents()
    w.grab().save(str(out / "2_preview.png"))
    print("status:", w.report.status.value, "pages:", w.report.total_pages)
    w.close()


if __name__ == "__main__":
    main()
