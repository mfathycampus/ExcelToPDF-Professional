"""End-to-end tests with every installed engine (skipped when none is installed)."""
import threading

import pymupdf
import pytest

from conftest import available_engines
from x2p.models import FinalStatus, JobOptions
from x2p.pipeline import Converter, sha256

ENGINES = available_engines()
pytestmark = [pytest.mark.engine,
              pytest.mark.skipif(not ENGINES, reason="no conversion engine installed")]

# fixture -> must the automatic mode deliver it?
EXPECT_PASS = ["13_long_text_appendix", "01_simple", "02_legacy", "03_multi_sheet", "05_arabic_rtl", "07_merged",
               "08_chart", "09_formulas_cf", "10_numbers", "11_long_text", "12_hidden_rows_cols"]
EXPECT_FAIL = {"90_narrow_hash": "cell_overflow_hash", "91_clipped_text": "text_missing"}


@pytest.mark.parametrize("engine", ENGINES)
@pytest.mark.parametrize("name", EXPECT_PASS)
def test_reference_files_pass(engine, name, fixtures, tmp_path):
    src = fixtures[name]
    before = sha256(src)
    out = tmp_path / f"{name}.pdf"
    rep = Converter().convert(src, out, JobOptions(engine=engine))
    assert rep.status is not FinalStatus.FAILED, [i.to_dict() for i in rep.issues]
    assert rep.delivered and out.exists()
    assert rep.source_unchanged is True and sha256(src) == before
    assert rep.min_font_pt is None or rep.min_font_pt >= 8 - 0.1
    with pymupdf.open(out) as doc:
        assert doc.page_count == rep.total_pages > 0
        assert [t[1] for t in doc.get_toc()] == rep.exported_sheets   # sheet order kept


@pytest.mark.parametrize("engine", ENGINES)
@pytest.mark.parametrize("name,code", sorted(EXPECT_FAIL.items()))
def test_data_loss_is_detected_and_blocks_delivery(engine, name, code, fixtures, tmp_path):
    # automatic fixes off: the raw rendering must be caught by verification
    out = tmp_path / "x.pdf"
    rep = Converter().convert(fixtures[name], out, JobOptions(engine=engine, wrap_overflow=False))
    assert rep.status is FinalStatus.FAILED
    assert code in {i.code for i in rep.issues}
    assert not rep.delivered and not out.exists()


@pytest.mark.parametrize("engine", ENGINES)
def test_clipped_text_is_fixed_and_reported(engine, fixtures, tmp_path):
    rep = Converter().convert(fixtures["91_clipped_text"], tmp_path / "c.pdf", JobOptions(engine=engine))
    assert rep.status is not FinalStatus.FAILED, [i.to_dict() for i in rep.issues]
    codes = {i.code for i in rep.issues}
    assert codes & {"overflow_text_wrapped", "long_text_to_appendix"}


@pytest.mark.parametrize("engine", ENGINES)
def test_long_text_goes_to_verified_appendix(engine, fixtures, tmp_path):
    out = tmp_path / "l.pdf"
    rep = Converter().convert(fixtures["13_long_text_appendix"], out, JobOptions(engine=engine))
    assert rep.status is not FinalStatus.FAILED, [i.to_dict() for i in rep.issues]
    assert "long_text_to_appendix" in {i.code for i in rep.issues}
    assert len(rep.exported_sheets) == 2                       # sheet + appendix, in that order
    appendix = rep.sheets[1]
    assert appendix.expected_texts > 0 and not appendix.missing_cells
    with pymupdf.open(out) as doc:
        first = doc[rep.sheet_page_ranges[rep.exported_sheets[0]][0]].get_text()
        assert "[1]" in first                                   # marker printed in the cell


@pytest.mark.parametrize("engine", ENGINES)
def test_wide_sheet_is_split_and_readable(engine, fixtures, tmp_path):
    rep = Converter().convert(fixtures["06_wide_long"], tmp_path / "w.pdf", JobOptions(engine=engine))
    assert rep.status is not FinalStatus.FAILED, [i.to_dict() for i in rep.issues]
    assert rep.min_font_pt >= 8 - 0.1
    assert rep.total_pages > 4


@pytest.mark.parametrize("engine", ENGINES)
def test_no_overwrite_without_consent(engine, fixtures, tmp_path):
    out = tmp_path / "exists.pdf"
    out.write_bytes(b"keep me")
    with pytest.raises(FileExistsError):
        Converter().convert(fixtures["01_simple"], out, JobOptions(engine=engine))
    assert out.read_bytes() == b"keep me"


@pytest.mark.parametrize("engine", ENGINES)
def test_two_concurrent_jobs(engine, fixtures, tmp_path):
    results = {}

    def run(name):
        results[name] = Converter().convert(fixtures[name], tmp_path / f"{name}.pdf", JobOptions(engine=engine))

    threads = [threading.Thread(target=run, args=(n,)) for n in ("01_simple", "07_merged")]
    for t in threads:
        t.start()
    for t in threads:
        t.join(300)
    assert all(r.status is not FinalStatus.FAILED for r in results.values())


@pytest.mark.parametrize("engine", ENGINES)
def test_cancel_cleans_up(engine, fixtures, tmp_path):
    from x2p import tempstore
    before = set(tempstore.root().iterdir())
    cancel = threading.Event()
    cancel.set()
    rep = Converter().convert(fixtures["01_simple"], tmp_path / "c.pdf", JobOptions(engine=engine), cancel=cancel)
    assert "cancelled" in {i.code for i in rep.issues}
    assert set(tempstore.root().iterdir()) <= before


@pytest.mark.parametrize("engine", ENGINES)
def test_card_mode_is_complete(engine, fixtures, tmp_path):
    from x2p.models import LayoutMode
    out = tmp_path / "cards.pdf"
    rep = Converter().convert(fixtures["06_wide_long"], out, JobOptions(engine=engine, mode=LayoutMode.CARDS))
    assert rep.status is not FinalStatus.FAILED, [i.to_dict() for i in rep.issues]
    v = rep.sheets[0]
    # 300 records x 40 fields + a title per record, every one found in the PDF
    assert v.expected_texts >= 300 * 40 and v.found_texts + len(v.rtl_unverified) == v.expected_texts
    assert rep.exported_sheets == ["Wide"] and v.min_font_pt >= 8
