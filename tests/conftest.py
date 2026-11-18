import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))


@pytest.fixture(scope="session")
def fixtures(tmp_path_factory) -> dict[str, Path]:
    import make_fixtures
    out = tmp_path_factory.mktemp("fixtures")
    return {p.stem: p for p in make_fixtures.build(out)}


def available_engines() -> list[str]:
    from x2p.engines import engine_status
    return [name for name, ok in engine_status().items() if ok]
