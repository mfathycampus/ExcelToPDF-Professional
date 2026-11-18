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


def test_update_button_finds_new_release(app, monkeypatch, tmp_path):
    """Manual check runs off the UI thread, then offers the update and relabels the button."""
    import json
    import http.server
    import threading
    from functools import partial
    srv = tmp_path / "srv"
    srv.mkdir()
    (srv / "latest.json").write_text(json.dumps({"tag_name": "v99.0.0", "body": "", "html_url": "x", "assets": [
        {"name": "a-win64.zip", "browser_download_url": "x"}, {"name": "a-win64.zip.sha256", "browser_download_url": "x"}]}))
    handler = partial(http.server.SimpleHTTPRequestHandler, directory=str(srv))
    handler.log_message = lambda *a, **k: None
    httpd = http.server.ThreadingHTTPServer(("127.0.0.1", 0), handler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    monkeypatch.setenv("X2P_UPDATE_URL", f"http://127.0.0.1:{httpd.server_address[1]}/latest.json")
    from x2p.gui.main_window import MainWindow
    w = MainWindow()
    offered = []
    monkeypatch.setattr(w.updates, "_offer", offered.append)   # the real one opens a modal box
    w.updates.check(manual=True)
    _wait(app, lambda: offered, timeout=30)
    httpd.shutdown()
    assert offered[0].version == "99.0.0"
    assert "99.0.0" in w.act_update.text()
    w.close()
