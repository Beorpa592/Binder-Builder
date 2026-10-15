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


def test_new_binder_from_existing_structure(binder, tmp_path):
    from binder_builder.engine.manifest import Manifest
    from binder_builder.engine.tree import Tree

    s = BinderSession(binder)
    s.apply(s.plan())
    (binder / ".binder" / "templates").mkdir()
    s.export_templates()
    s.close()

    dest = tmp_path / "Next Year Binder"
    n = BinderSession.create_from_structure(binder, dest, "Next Year Binder")
    m = Manifest.load(dest)
    assert m.binder.name == "Next Year Binder"
    assert m.binder.header_lines == Manifest.load(binder).binder.header_lines
    tree = Tree(m)
    assert {x.id for x in tree.walk()} == {x.id for x in Tree(Manifest.load(binder)).walk()}
    assert tree.find("open-points").get("filename") is None  # legacy names dropped
    assert (dest / ".binder" / "templates" / "Section_Cover.docx").is_file()
    assert n.layout().codes == BinderSession(binder, lock=False).layout().codes
    p = n.plan()
    creates = [a for a in p.changes if a.kind == "create_file"]
    assert len(creates) == 21 and not p.orphans  # 20 Word/Excel items + the CKL; the PDF item needs a file
    assert n.apply(p).ok
    assert word.read_code(dest / "A_Administration" / "03_Top_10_List.docx") == "A.03"
    assert (dest / "0_Control_Section" / "00_Cover.docx").is_file()  # computed name, not the legacy one
    n.close()


def test_new_binder_copying_documents(binder, tmp_path):
    from docx import Document

    dest = tmp_path / "Copy"
    n = BinderSession.create_from_structure(binder, dest, "Copy", copy_documents=True)
    assert n.tree.find("top-10-list").get("copied_from")["source_path"].endswith("03_Top_10_List.docx")
    assert n.apply(n.plan()).ok
    text = " ".join(p.text for p in Document(str(dest / "A_Administration" / "03_Top_10_List.docx")).paragraphs)
    assert "Placeholder text" in text
    assert (dest / "B_Policies_and_Procedures" / "02_Signed_Policy_Acknowledgement.pdf").is_file()
    assert not [a for a in n.plan().warnings if "source changed" in a.detail]
    n.close()
