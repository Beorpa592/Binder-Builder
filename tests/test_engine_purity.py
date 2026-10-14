"""SPEC §13: the engine package must not import GUI or COM code."""

import ast
from pathlib import Path

ENGINE = Path(__file__).parent.parent / "binder_builder" / "engine"
FORBIDDEN = ("tkinter", "tkinterdnd2", "win32com", "win32api", "pythoncom", "pywintypes", "PySide6",
             "binder_builder.gui", "binder_builder.office")


def test_engine_has_no_gui_or_com_imports():
    bad = []
    for py in ENGINE.rglob("*.py"):
        for node in ast.walk(ast.parse(py.read_text(encoding="utf-8"))):
            names = []
            if isinstance(node, ast.Import):
                names = [a.name for a in node.names]
            elif isinstance(node, ast.ImportFrom) and node.module:
                names = [node.module]
                if node.level:
                    names = []  # relative imports stay inside the engine
            for n in names:
                if any(n == f or n.startswith(f + ".") for f in FORBIDDEN):
                    bad.append(f"{py.name}: {n}")
    assert not bad, bad
