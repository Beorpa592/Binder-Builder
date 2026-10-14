"""GUI smoke tests: build the real tkinter screens and drive them (no dialogs left open).

Skipped where Tk has no display.
"""

import pytest

tk = pytest.importorskip("tkinter")


@pytest.fixture
def app(monkeypatch, tmp_path):
    from binder_builder.gui import app as app_mod

    monkeypatch.setenv("APPDATA", str(tmp_path / "appdata"))
    for name in ("showinfo", "showwarning", "showerror"):
        monkeypatch.setattr(app_mod.messagebox, name, lambda *a, **k: None)
    try:
        a = app_mod.App(check_updates=False)
    except tk.TclError as e:
        pytest.skip(f"no display: {e}")
    a.geometry("1200x760+0+0")
    a.update()
    yield a
    import gc

    a.on_close()
    gc.collect()  # free Tk objects on this (the Tk) thread


def _pump(app, ms=50):
    for _ in range(max(1, ms // 10)):
        app.update()
        app.after(10)


def test_start_screen_and_open_binder(app, binder):
    from binder_builder.gui.binder_view import BinderView

    app.open_binder(binder)
    assert isinstance(app.view, BinderView)
    assert app.session is not None and not app.session.read_only
    assert (binder / ".binder" / "binder.lock").exists()
    tv = app.view.tv
    assert tv.set("top-10-list", "code") == "A.03"
    assert app.recent.load()[0] == str(binder)
    app.close_binder()
    assert not (binder / ".binder" / "binder.lock").exists()


def test_toolbar_edits_and_details(app, binder):
    app.open_binder(binder)
    view = app.view
    tv = view.tv
    tv.selection_set("org-chart")
    _pump(app)
    assert view.details.vars["title"].get() == "Organizational Chart"
    view._edit(app.session.move_down)
    assert tv.set("org-chart", "code") == "A.03"
    view._edit(app.session.move_up)
    assert tv.set("org-chart", "code") == "A.02"

    # include toggle by clicking the Include column
    view._toggle_include("org-chart")
    assert tv.set("org-chart", "incl") == "☐"
    view._toggle_include("org-chart")

    # details pane save renames without changing the id
    tv.selection_set("org-chart")
    _pump(app)
    view.details.vars["title"].set("Org Chart")
    view.details._save()
    assert app.session.tree.find("org-chart").text == "Org Chart"
    assert tv.item("org-chart", "text").startswith("Org Chart")


def test_drag_and_drop_moves_node_with_live_code(app, binder):
    app.open_binder(binder)
    view = app.view
    tv = view.tv
    tv.see("policies")
    _pump(app)
    src = tv.bbox("outreach-efforts")
    dst = tv.bbox("media-policy")
    assert src and dst
    x = src[0] + 20
    y0 = src[1] + src[3] // 2
    y1 = dst[1] + 2  # top edge of the target row: drop before it
    tv.event_generate("<ButtonPress-1>", x=x, y=y0)
    tv.event_generate("<B1-Motion>", x=x, y=y0 + 10)
    tv.event_generate("<B1-Motion>", x=x, y=y1)
    _pump(app)
    assert "becomes B.01" in view.busy.cget("text")
    tv.event_generate("<ButtonRelease-1>", x=x, y=y1)
    _pump(app)
    assert app.session.layout().codes["outreach-efforts"] == "B.01"
    assert tv.set("outreach-efforts", "code") == "B.01"


def test_sync_dialog_plans_and_applies(app, binder, monkeypatch):
    from binder_builder.gui import sync_view

    monkeypatch.setattr(sync_view.messagebox, "askyesno", lambda *a, **k: True)
    monkeypatch.setattr(sync_view.messagebox, "showinfo", lambda *a, **k: None)
    app.open_binder(binder)
    dlg = sync_view.SyncDialog(app.view, app, app.session)
    for _ in range(300):
        _pump(app, 20)
        if dlg.plan is not None:
            break
    assert dlg.plan is not None and dlg.plan.changes
    assert len(dlg.table.get_children()) == len(dlg.plan.actions)
    dlg.filters["stamp_header"].set(False)
    dlg._fill()
    assert len(dlg.table.get_children()) < len(dlg.plan.actions)
    dlg._apply()
    for _ in range(600):
        _pump(app, 20)
        if dlg.plan is not None and dlg.plan.is_empty:
            break
    assert dlg.plan.is_empty
    dlg.destroy()
    _pump(app)


def test_settings_dialog_saves(app, binder):
    from binder_builder.gui.settings import SettingsDialog

    app.open_binder(binder)
    dlg = SettingsDialog(app.view, app.session)
    _pump(app)
    assert "Packet Ref # A.01" in dlg.preview.cget("text")
    dlg.ref_label.set("Ref #:")
    dlg.vars_text.delete("1.0", "end")
    dlg.vars_text.insert("1.0", "org_short = NS")
    assert dlg._save()
    assert app.session.binder.ref_label == "Ref #:"
    assert app.session.binder.vars == {"org_short": "NS"}


def test_read_only_when_locked_by_someone_else(app, binder):
    import json

    (binder / ".binder" / "binder.lock").write_text(json.dumps(
        {"user": "someone", "machine": "DESK-07", "since": "2099-01-01T09:00:00"}))
    app.open_binder(binder)
    assert app.session.read_only
    app.update()
    assert app.view.banner.winfo_ismapped()
    app.close_binder()
    assert (binder / ".binder" / "binder.lock").exists()  # someone else's lock is left alone
