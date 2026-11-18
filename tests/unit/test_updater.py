"""Updater: version compare, release parsing, checksum enforcement, safe extraction.
Uses a local HTTP server standing in for the GitHub API (X2P_UPDATE_URL)."""
import hashlib
import http.server
import json
import threading
import zipfile
from functools import partial

import pytest

from x2p import __version__, updater


def _make_pkg(path, extra=None):
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("ExcelToPDF/ExcelToPDF.exe", b"MZ fake")
        z.writestr("ExcelToPDF/_internal/lib.dll", b"lib")
        for name, data in (extra or {}).items():
            z.writestr(name, data)


@pytest.fixture
def server(tmp_path, monkeypatch):
    root = tmp_path / "srv"
    root.mkdir()
    handler = partial(http.server.SimpleHTTPRequestHandler, directory=str(root))
    handler.log_message = lambda *a, **k: None
    httpd = http.server.ThreadingHTTPServer(("127.0.0.1", 0), handler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    base = f"http://127.0.0.1:{httpd.server_address[1]}"

    def publish(version, sha=None, zip_extra=None):
        z = root / f"pkg-{version}-win64.zip"
        _make_pkg(z, zip_extra)
        digest = sha or hashlib.sha256(z.read_bytes()).hexdigest()
        (root / (z.name + ".sha256")).write_text(f"{digest}  {z.name}\n")
        (root / "latest.json").write_text(json.dumps({
            "tag_name": f"v{version}", "body": "notes", "html_url": base + "/page",
            "assets": [{"name": z.name, "size": z.stat().st_size, "browser_download_url": f"{base}/{z.name}"},
                       {"name": z.name + ".sha256", "browser_download_url": f"{base}/{z.name}.sha256"}]}))
    monkeypatch.setenv("X2P_UPDATE_URL", base + "/latest.json")
    yield publish
    httpd.shutdown()


def test_version_compare():
    assert updater.is_newer("0.10.0", "0.9.9")
    assert updater.is_newer("1.0.0", "0.3.2")
    assert not updater.is_newer("0.3.1", "0.3.1")
    assert not updater.is_newer("0.3.0", "0.3.1")


def test_no_update_when_same_version(server):
    server(__version__)
    assert updater.check_latest() is None


def test_download_verify_and_stage(server, tmp_path):
    server("99.0.0")
    info = updater.check_latest()
    assert info and info.version == "99.0.0" and info.notes == "notes"
    seen = []
    z = updater.download(info, tmp_path / "dl", progress=seen.append)
    assert seen and seen[-1] == 1.0
    app = updater.stage(z, tmp_path / "stage")
    assert (app / "ExcelToPDF.exe").is_file() and (app / "_internal" / "lib.dll").is_file()


def test_checksum_mismatch_is_rejected(server, tmp_path):
    server("99.0.0", sha="0" * 64)
    info = updater.check_latest()
    with pytest.raises(updater.UpdateError) as e:
        updater.download(info, tmp_path / "dl")
    assert e.value.code == "update_bad_checksum"
    assert not list((tmp_path / "dl").glob("*.zip"))


@pytest.mark.parametrize("evil", ["ExcelToPDF/../../escape.txt", "other/file.txt"])
def test_stage_rejects_unsafe_paths(tmp_path, evil):
    z = tmp_path / "evil.zip"
    _make_pkg(z, {evil: b"x"})
    with pytest.raises(updater.UpdateError):
        updater.stage(z, tmp_path / "stage")
    assert not (tmp_path / "escape.txt").exists()


def test_unreachable_server_reports_error(monkeypatch):
    monkeypatch.setenv("X2P_UPDATE_URL", "http://127.0.0.1:9/none.json")
    with pytest.raises(updater.UpdateError) as e:
        updater.check_latest(timeout=3)
    assert e.value.code == "update_check_failed"


def test_not_self_updating_outside_package():
    assert updater.install_dir() is None and not updater.can_self_update()
