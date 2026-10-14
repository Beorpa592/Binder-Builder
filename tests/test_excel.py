import zipfile

from binder_builder.engine import excel
from binder_builder.engine.codes import HeaderSpec

from .conftest import FIXTURE

SPEC = HeaderSpec(("North Star", "Index of Genres & Channels"), "D.01", "Packet Ref #", ("", "", ""))
ROWS = [("top-10-list", "A.03", "Top 10 List", "03_Top_10_List.docx"),
        ("mrp-summary", "A.01", "North Star MRP Summary", "01_North_Star_MRP_Summary.docx")]


def test_header_and_binder_sheet(tmp_path):
    src = FIXTURE / "0_Control_Section/01_Open_Points.xlsx"
    dst = tmp_path / "out.xlsx"
    excel.update(src, dst, SPEC, ROWS)
    wb = excel.load(dst)
    assert excel.header_matches(wb, SPEC)
    ws = wb.worksheets[0]
    assert ws.oddHeader.center.text == "North Star\nIndex of Genres && Channels"
    assert ws.oddHeader.right.text == "Packet Ref # D.01\nPage &P of &N"
    assert wb[excel.BINDER_SHEET].sheet_state == "hidden"
    assert wb.active.title != excel.BINDER_SHEET
    assert excel.resolve_name(dst, "bb_top_10_list") == "A.03"
    assert wb[excel.BINDER_SHEET]["B2"].number_format == "@"
    assert excel.read_code(dst) == "D.01"
    assert excel.body_hash(src) == excel.body_hash(dst)
    st = excel.inspect(dst, SPEC, ROWS)
    assert st.header_ok and st.refs_ok and not st.broken


def test_rows_change_detected_and_codes_are_strings(tmp_path):
    src = FIXTURE / "0_Control_Section/01_Open_Points.xlsx"
    dst = tmp_path / "out.xlsx"
    excel.update(src, dst, SPEC, ROWS)
    rows2 = [("top-10-list", "A.04", "Top 10 List", "04_Top_10_List.docx"), ROWS[1]]
    assert not excel.inspect(dst, SPEC, rows2).refs_ok
    excel.update(dst, dst, SPEC, [("control", "0", "Control", "0_Control_Section"),
                                  ("open-points", "0.01", "Open Points", "01_Open_Points.xlsx")] + rows2)
    assert excel.resolve_name(dst, "bb_open_points") == "0.01"
    assert excel.resolve_name(dst, "bb_top_10_list") == "A.04"


def test_broken_reference_gets_question_mark(tmp_path):
    src = FIXTURE / "0_Control_Section/01_Open_Points.xlsx"
    dst = tmp_path / "out.xlsx"
    only_mrp = [ROWS[1]]
    st = excel.inspect(src, SPEC, only_mrp)
    assert "bb_top_10_list" in st.broken
    assert st.broken["bb_top_10_list"] == ["Open Points - Action Items!B2"]
    excel.update(src, dst, SPEC, only_mrp)
    assert excel.resolve_name(dst, "bb_top_10_list") == excel.BROKEN


def test_body_hash_detects_cell_change(tmp_path):
    src = FIXTURE / "A_Administration/05_Budget_Overview.xlsx"
    wb = excel.load(src)
    h0 = excel.body_hash_wb(wb)
    wb.worksheets[0]["B2"] = 101
    assert excel.body_hash_wb(wb) != h0


def test_unsafe_parts_detected(tmp_path):
    src = FIXTURE / "A_Administration/05_Budget_Overview.xlsx"
    assert excel.unsafe_parts(src) == []
    p = tmp_path / "chart.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(p, "w") as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        zout.writestr("xl/charts/chart1.xml", "<c/>")
    assert excel.unsafe_parts(p) == ["xl/charts/"]


def test_sheet_title_sanitised(tmp_path):
    assert excel.sheet_title("Open Points / Action Items") == "Open Points  Action Items"
    assert len(excel.sheet_title("x" * 40)) == 31
