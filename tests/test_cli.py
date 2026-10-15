"""Command line (SPEC §13): sync, import, backups/restore, unlock, new-from, diagnose."""

import json

from binder_builder.cli import main
from binder_builder.engine.storage import LocalFolderStorage


def test_unlock_shows_and_breaks_only_stale_locks(binder, capsys):
    lock = binder / ".binder" / "binder.lock"
    assert main(["unlock", str(binder)]) == 0
    assert "Not locked" in capsys.readouterr().out
    lock.write_text(json.dumps({"user": "someone", "machine": "DESK-07", "since": "2099-01-01T09:00:00"}))
    assert main(["unlock", str(binder), "--break"]) == 2  # fresh lock: refused
    assert lock.exists()
    lock.write_text(json.dumps({"user": "someone", "machine": "DESK-07", "since": "2020-01-01T09:00:00"}))
    assert main(["unlock", str(binder), "--break"]) == 0
    assert LocalFolderStorage(binder).lock_info() is None


def test_new_from_and_diagnose(binder, tmp_path, capsys):
    dest = tmp_path / "Next"
    assert main(["new-from", str(binder), str(dest), "--name", "Next Year"]) == 0
    assert main(["sync", str(dest), "--apply", "--yes"]) == 0
    capsys.readouterr()
    assert main(["diagnose", str(dest)]) == 0
    out = capsys.readouterr().out
    assert "Binder Builder" in out and "Manifest: format 1" in out and "Lock: none" in out
    assert "Recent syncs:" in out


def test_backups_and_restore(binder, capsys):
    assert main(["sync", str(binder), "--apply", "--yes"]) == 0
    capsys.readouterr()
    assert main(["backups", str(binder)]) == 0
    backup = capsys.readouterr().out.split()[0]
    assert main(["restore", str(binder), backup]) == 0
    assert "Restored" in capsys.readouterr().out


def test_export_pdf_refuses_merged_on_unsynced_binder(binder, capsys):
    import pytest

    pytest.importorskip("pythoncom")
    assert main(["export-pdf", str(binder), "--merged"]) == 2
    assert "Sync the binder first" in capsys.readouterr().err
