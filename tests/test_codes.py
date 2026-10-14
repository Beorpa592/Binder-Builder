from pathlib import Path

from binder_builder.engine.codes import compute_layout, header_spec
from binder_builder.engine.manifest import Manifest
from binder_builder.engine.tree import Tree

from .conftest import FIXTURE


def layout_of(text: str):
    m = Manifest.loads(text)
    return compute_layout(Tree(m), m.binder)


def fixture_layout():
    m = Manifest.load(FIXTURE)
    return compute_layout(Tree(m), m.binder), m


def test_fixture_codes():
    layout, _ = fixture_layout()
    c = layout.codes
    assert c["control"] == "0"
    assert c["open-points"] == "0.01" and isinstance(c["open-points"], str)
    assert c["administration"] == "A"
    assert c["top-10-list"] == "A.03"
    assert c["top-10-list-backup"] == "A.03.PB"
    assert c["points-of-contact"] == "A.04"
    assert c["outreach-efforts"] == "A.07"
    assert c["policies"] == "B"
    assert c["retired-policy"] == "B.04"  # included: false keeps its code
    assert c["website-media"] == "D"  # clarifying header consumes no letter
    assert c["web-index"] == "D.01"
    assert c["weekly-digest"] == "D.01.01"
    assert c["print-media"] == "E"
    assert "content-mediums" not in c


def test_fixture_paths_match_files_on_disk():
    layout, _ = fixture_layout()
    for e in layout.entities:
        assert (FIXTURE / e.rel_path).is_file(), e.rel_path


def test_specific_paths():
    layout, _ = fixture_layout()
    assert layout.entity("open-points").rel_path == "0_Control_Section/01_Open_Points.xlsx"
    assert layout.entity("control#cover").rel_path == "0_Control_Section/00_Cover_Sheet_CS.docx"
    assert layout.entity("control#cover").code == "0.00"
    assert layout.entity("top-10-list-backup").rel_path == "A_Administration/03_PB_Top_10_List_Backup.docx"
    assert layout.entity("weekly-digest").rel_path == "D_Website_Media/01_01_Weekly_Digest_Monthly_Summary.xlsx"
    assert layout.entity("#toc").rel_path == "01_TOC_m.xlsx"
    assert layout.entity("#toc").code == "TOC"
    assert layout.entity("#ckl").code == "CKL"
    assert layout.entity("retired-policy").included is False


def test_three_digits_past_99():
    items = "\n".join(f"      - {{id: i{n}, kind: item, title: 'Item {n}', ext: docx}}" for n in range(1, 101))
    layout = layout_of(f"format_version: 1\nbinder: {{name: B}}\ntree:\n  - id: s\n    kind: section\n    label: Big\n    children:\n{items}\n")
    assert layout.codes["i1"] == "A.001"
    assert layout.codes["i100"] == "A.100"
    assert layout.entity("i7").rel_path == "A_Big/007_Item_7.docx"


def test_sub_sections_take_next_letter_and_nest():
    layout = layout_of("""
format_version: 1
binder: {name: B}
tree:
  - id: a
    kind: section
    label: Alpha
    children:
      - {id: a1, kind: item, title: One, ext: docx}
      - id: sub
        kind: section
        label: Sub Part
        children:
          - {id: s1, kind: item, title: Deep, ext: xlsx}
  - id: c
    kind: section
    label: Gamma
""")
    assert layout.codes == {"a": "A", "a1": "A.01", "sub": "B", "s1": "B.01", "c": "C"}
    assert layout.entity("s1").rel_path == "A_Alpha/B_Sub_Part/01_Deep.xlsx"


def test_tab_labels_and_fit_warning():
    layout = layout_of("""
format_version: 1
binder: {name: B}
tree:
  - {id: a, kind: section, label: Administration}
  - {id: b, kind: section, label: Policies and Procedures}
  - {id: c, kind: section, label: "An Extremely Long Section Label Over Forty"}
""")
    labels = {s.node.id: s.tab_label for s in layout.sections}
    assert labels["a"] == "A) Administration"
    assert labels["b"] == "Policies and Procedures"  # 26 chars with letter -> label only
    assert any("over the fit limit" in w for w in layout.warnings)


def test_header_spec_renders_placeholders():
    layout, m = fixture_layout()
    spec = header_spec(layout.entity("web-index"), m.binder)
    assert spec.lines == ("North Star", "North Star Media", "Management Binder", "Index of Genres & Channels")
    assert spec.code == "D.01"
    assert spec.ref_label == "Packet Ref #"
