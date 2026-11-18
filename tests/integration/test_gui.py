"""GUI state: after each job finishes, the actions must be usable again (regression: they stayed
disabled because the result arrived while the worker thread was still shutting down)."""
import os
import time

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
QtWidgets = pytest.importorskip("PySide6.QtWidgets")

from conftest import available_engines  # noqa: E402


def _wait(app, cond, timeout=300):
    t0 = time.time()
    while not cond():
        app.processEvents()
        time.sleep(0.02)
        assert time.time() - t0 < timeout, "timed out"
    app.processEvents()


@pytest.fixture(scope="module")
def app():
    return QtWidgets.QApplication.instance() or QtWidgets.QApplication([])


def test_actions_enabled_after_inspect_and_preview(app, fixtures):
    if not available_engines():
        pytest.skip("no engine installed")
    from x2p.gui.main_window import MainWindow
    w = MainWindow()
    w.show()
    for _ in range(5):  # repeat: the original bug depended on thread timing
        w.open_file(str(fixtures["01_simple"]))
        assert w.runner.busy and not w.act_export.isEnabled() and w.act_cancel.isEnabled()
        _wait(app, lambda: not w.runner.busy)
        assert w.info is not None
        assert w.act_preview.isEnabled() and w.act_export.isEnabled() and w.act_open.isEnabled()
        assert not w.act_cancel.isEnabled() and not w.progress.isVisible()
    w.preview()
    assert w.runner.busy and w.progress.isVisible()
    _wait(app, lambda: not w.runner.busy)
    assert w.report is not None and w.report.temp_pdf
    assert w.act_preview.isEnabled() and w.act_export.isEnabled()
    assert not w.act_cancel.isEnabled()
    w.close()
