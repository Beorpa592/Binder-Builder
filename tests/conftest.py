import shutil
from pathlib import Path

import pytest

FIXTURE = Path(__file__).parent / "fixtures" / "synthetic_binder"


@pytest.fixture
def binder(tmp_path) -> Path:
    """A fresh copy of the synthetic binder. Tests never sync the fixture in place."""
    dest = tmp_path / "North Star Media Management Binder"
    shutil.copytree(FIXTURE, dest)
    return dest
