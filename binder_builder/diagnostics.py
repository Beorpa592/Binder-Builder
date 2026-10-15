"""Problem reports (phase 5 error reporting): one plain-text summary a user can copy into an email.

Contains versions, the binder's lock and manifest status, recent sync log entries and the tail
of the error log. No document content is included.
"""

from __future__ import annotations

import os
import platform
import sys
from pathlib import Path

from . import __version__


def _app_dir() -> Path:
    return Path(os.environ.get("APPDATA") or Path.home()) / "BinderBuilder"


def _tail(path: Path, lines: int) -> list[str]:
    try:
        return path.read_text(encoding="utf-8", errors="replace").splitlines()[-lines:]
    except OSError:
        return []


def report(folder: Path | str | None = None) -> str:
    out = [
        f"Binder Builder {__version__}",
        f"Python {platform.python_version()} ({'packaged' if getattr(sys, 'frozen', False) else 'source'})",
        f"Windows {platform.release()} {platform.version()}",
    ]
    try:
        import win32com.client  # noqa: F401

        out.append("Office automation: available")
    except ImportError:
        out.append("Office automation: pywin32 not installed")

    if folder:
        from .engine import sync
        from .engine.manifest import Manifest, ManifestError
        from .engine.storage import LocalFolderStorage

        folder = Path(folder)
        out += ["", f"Binder folder: {folder}"]
        try:
            m = Manifest.load(folder)
            out.append(f"Manifest: format {m.raw.get('format_version')}, "
                       f"{sum(1 for _ in _walk(m.raw.get('tree')))} entries, page numbers {m.binder.page_numbers}")
        except ManifestError as e:
            out.append(f"Manifest problem: {e}")
        storage = LocalFolderStorage(folder)
        info = storage.lock_info()
        out.append(f"Lock: {info.user} on {info.machine} since {info.since:%Y-%m-%d %H:%M}" if info else "Lock: none")
        out.append(f"Backups: {len(sync.list_backups(folder))}")
        out.append("Recent syncs:")
        for entry in storage.history()[-10:]:
            out.append(f"  {entry}")

    out += ["", "Recent errors:"]
    out += [f"  {line}" for line in _tail(_app_dir() / "error.log", 60)] or ["  (none)"]
    return "\n".join(out)


def _walk(nodes):
    for n in nodes or []:
        yield n
        yield from _walk(n.get("children"))
