"""Regenerate tests/fixtures/golden/: text dumps of every generated document after a first sync.

    python tests/fixtures/make_golden.py

Review the diff before committing: a golden change is a change in what users see.
"""

from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from binder_builder.engine import sync  # noqa: E402
from tests.dump import dump  # noqa: E402

HERE = Path(__file__).resolve().parent
FIXTURE = HERE / "synthetic_binder"
GOLDEN = HERE / "golden"

GENERATED_FILES = [
    "00_Cover.docx",
    "01_TOC_m.xlsx",
    "02_CKL.xlsx",
    "Tabs.docx",
    "0_Control_Section/00_Cover_Sheet_CS.docx",
    "A_Administration/00_Cover.docx",
    "B_Policies_and_Procedures/00_Cover.docx",
    "C_Campaign_Management/00_Cover.docx",
    "D_Website_Media/00_Cover.docx",
    "E_Print_Media/00_Cover.docx",
]


def golden_name(rel: str) -> str:
    return rel.replace("/", "__") + ".txt"


def synced_copy(dest: Path) -> Path:
    shutil.copytree(FIXTURE, dest)
    r = sync.apply(sync.plan(dest))
    assert r.ok, r.failures
    return dest


def main() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        b = synced_copy(Path(tmp) / "binder")
        if GOLDEN.exists():
            shutil.rmtree(GOLDEN)
        GOLDEN.mkdir()
        for rel in GENERATED_FILES:
            (GOLDEN / golden_name(rel)).write_text(dump(b / rel), encoding="utf-8")
    print(f"wrote {len(GENERATED_FILES)} golden dumps to {GOLDEN}")


if __name__ == "__main__":
    main()
