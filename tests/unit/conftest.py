import pytest

from small_business_reporting_automation import synthetic

from sample_oracle import SEED


@pytest.fixture(scope="session")
def sample_root(tmp_path_factory):
    """Dataset generado una vez con la semilla confirmada (directorio temporal; no toca ``data/``)."""
    root = tmp_path_factory.mktemp("sample") / "data"
    synthetic.generate(SEED, root)
    return root
