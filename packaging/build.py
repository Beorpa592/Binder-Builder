"""Build the single-folder Windows app: dist/BinderBuilder/BinderBuilder.exe, plus a release zip.

    python packaging/build.py            # needs: pip install -e .[dev]
"""

from __future__ import annotations

import shutil
import sys
import zipfile
from pathlib import Path

import PyInstaller.__main__

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from binder_builder import __version__  # noqa: E402

DIST = ROOT / "dist"
APP = DIST / "BinderBuilder"


def build() -> Path:
    PyInstaller.__main__.run([
        str(ROOT / "packaging" / "entry.py"),
        "--name", "BinderBuilder",
        "--onedir",
        "--windowed",
        "--noconfirm",
        "--clean",
        "--distpath", str(DIST),
        "--workpath", str(ROOT / "build"),
        "--specpath", str(ROOT / "build"),
        "--add-data", f"{ROOT / 'binder_builder' / 'templates'};binder_builder/templates",
        "--collect-data", "docx",
        "--collect-data", "docxtpl",
        "--hidden-import", "binder_builder.engine.generate.cover",
        "--hidden-import", "binder_builder.engine.generate.section_cover",
        "--hidden-import", "binder_builder.engine.generate.tabs",
        "--hidden-import", "binder_builder.engine.generate.toc",
    ])
    zpath = DIST / f"BinderBuilder-{__version__}-win64.zip"
    if zpath.exists():
        zpath.unlink()
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for f in APP.rglob("*"):
            z.write(f, Path("BinderBuilder") / f.relative_to(APP))
    return zpath


if __name__ == "__main__":
    print(f"built {build()}")
