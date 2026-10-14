"""The GUI-free session behind every screen: edits, validation, lock, live codes."""

import json

import pytest

from binder_builder.engine import word
from binder_builder.engine.session import BinderSession, ReadOnlyError


def test_create_binder_and_add_structure(tmp_path):
    s = BinderSession.create(tmp_path / "New Binder", "New Binder")
    a = s.add_section("Administration")
    s.add_document("Top 10 List", "docx", a.id, "child")
    s.add_document("Budget", "xlsx", "top-10-list", "after")
    b = s.add_section("Policies", a.id, "after")
    s.add_document("Media Policy", "docx", b.id)  # "after" a section means inside it
    s.close()

    s2 = BinderSession(tmp_path / "New Binder")
    codes = s2.layout().codes
    assert codes == {"administration": "A", "top-10-list": "A.01", "budget": "A.02", "policies": "B",
                     "media-policy": "B.01"}
    text = (tmp_path / "New Binder" / ".binder" / "binder.yaml").read_text()
    assert "- id: administration" in text  # block style, readable
    assert "# line 1 left" in text  # template comments kept
    s2.close()


def test_invalid_placements_rejected(binder):
    s = BinderSession(binder)
    with pytest.raises(ValueError):
        s.add_document("Loose", "docx")  # top level
    with pytest.raises(ValueError):
        s.move("administration", "top-10-list", 0)  # section inside item
    with pytest.raises(ValueError):
        s.add_document("", "docx", "administration", "child")
    s.close()


def test_indent_outdent_and_move(binder):
    s = BinderSession(binder)
    s.indent("org-chart")  # under mrp-summary -> A.01.01
    assert s.layout().codes["org-chart"] == "A.01.01"
    s.outdent("org-chart")
    assert s.layout().codes["org-chart"] == "A.02"
    s.move_down("org-chart")
    assert s.layout().codes["org-chart"] == "A.03"
    s.move_up("org-chart")
    assert s.layout().codes["org-chart"] == "A.02"
    assert s.preview_move("org-chart", "policies", 0) == "B.01"
    assert s.layout().codes["org-chart"] == "A.02"  # preview did not change anything
    assert s.preview_move("org-chart", None, 0) is None  # invalid drop
    s.close()


def test_update_node_keeps_id_and_validates(binder):
    s = BinderSession(binder)
    s.update_node("top-10-list", title="Top Ten List", locked_header=True)
    node = s.tree.find("top-10-list")
    assert node.text == "Top Ten List" and node.get("locked_header") is True
    s.update_node("top-10-list", locked_header=False, included=True)
    assert "locked_header" not in node.raw and "included" not in node.raw
    with pytest.raises(Exception):
        s.update_node("top-10-list", ext="pptx")
    assert node.get("ext") == "docx"
    s.close()


def test_lock_held_while_open_and_read_only_for_others(binder):
    s = BinderSession(binder)
    lock = json.loads((binder / ".binder" / "binder.lock").read_text())
    s.close()
    assert not (binder / ".binder" / "binder.lock").exists()

    (binder / ".binder" / "binder.lock").write_text(json.dumps(
        {"user": "someone", "machine": "DESK-07", "since": "2026-10-14T09:00:00"}))
    other = BinderSession(binder)
    assert other.read_only and other.lock_info.user == "someone"
    with pytest.raises(ReadOnlyError):
        other.add_section("Nope")
    assert other.lock_is_stale()  # older than 4 hours
    other.break_lock()
    assert not other.read_only
    other.add_section("Now allowed")
    other.close()
    assert lock["user"]


def test_sync_through_session_and_export_templates(binder):
    s = BinderSession(binder)
    r = s.apply(s.plan())
    assert r.ok
    st = s.status(s.plan())
    assert st.changes == 0 and st.orphans == 0
    written = s.export_templates()
    assert len(written) == 5
    assert s.export_templates() == []  # never overwrites
    assert s.status(s.plan()).changes == 0  # identical templates: nothing to regenerate
    assert word.read_code(s.path_of("top-10-list")) == "A.03"
    s.close()
