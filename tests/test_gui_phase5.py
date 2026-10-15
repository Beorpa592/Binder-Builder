"""GUI pieces added in phase 5: problem report, descriptive restore list, new-from-existing."""

import pytest

from .test_gui import _pump, app  # noqa: F401 - the shared Tk fixture

tk = pytest.importorskip("tkinter")


def test_problem_report_dialog(app, binder):  # noqa: F811
    from binder_builder.gui.app import ProblemReport

    app.open_binder(binder)
    dlg = ProblemReport(app, binder)
    _pump(app)
    text = dlg.text.get("1.0", "end")
    assert "Binder Builder" in text and str(binder) in text and "Lock:" in text
    dlg._copy()
    assert "Binder Builder" in app.clipboard_get()
    dlg.destroy()


def test_restore_dialog_describes_backups(app, binder):  # noqa: F811
    from binder_builder.gui.sync_view import RestoreDialog

    app.open_binder(binder)
    app.session.apply(app.session.plan())
    infos = [app.session.backup_info(b) for b in app.session.backups()]
    dlg = RestoreDialog(app, infos)
    _pump(app)
    assert "Sync" in dlg.lb.get(0) and "file(s)" in dlg.lb.get(0)
    dlg._ok()
    assert dlg.choice == infos[0]["name"]


def test_new_from_existing_flow(app, binder, tmp_path, monkeypatch):  # noqa: F811
    from binder_builder.gui import app as app_mod

    answers = iter([str(binder), str(tmp_path)])
    monkeypatch.setattr(app_mod.filedialog, "askdirectory", lambda **k: next(answers))
    monkeypatch.setattr(app_mod.simpledialog, "askstring", lambda *a, **k: "Next Year")
    monkeypatch.setattr(app_mod.messagebox, "askyesno", lambda *a, **k: False)
    app.new_from_existing()
    assert app.session is not None and app.session.folder == tmp_path / "Next Year"
    assert app.session.layout().codes["top-10-list"] == "A.03"
