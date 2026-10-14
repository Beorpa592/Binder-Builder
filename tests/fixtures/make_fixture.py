"""Regenerate tests/fixtures/synthetic_binder/ (see README.md in this folder).

Everything here is placeholder text. Never put real binder content in this repo.

    python tests/fixtures/make_fixture.py     # writes synthetic_binder/ and legacy_binder/
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import Workbook
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

HERE = Path(__file__).resolve().parent
OUT = HERE / "synthetic_binder"
LEGACY_OUT = HERE / "legacy_binder"

# Legacy fixture differences (SPEC §14 phase 3): references typed as plain text, one stale
# header code, one stale filename number, no .binder/ folder.
LEGACY_NAME_CODES = {"bb_top_10_list": "A.03", "bb_mrp_summary": "A.01", "bb_open_points": "0.01",
                     "bb_campaign_briefs": "C.02"}
LEGACY_STALE_HEADER = {"A_Administration/04_Points_of_Contact.docx": "A.05"}
LEGACY_RENAMES = {"A_Administration/06_Meeting_Minutes.docx": "A_Administration/05_Meeting_Minutes.docx"}

LOREM = (
    "Placeholder text. This paragraph stands in for author-written body content and "
    "exists only so the body invariant check has something to protect."
)

MANIFEST = """\
# Synthetic fixture binder. Placeholder content only.
format_version: 1
binder:
  name: North Star Media Management Binder
  header_lines:                      # author-defined, propagated to every document
    - "{{org_short}}"                # line 1 left; the ref code goes line 1 right automatically
    - "North Star Media"
    - "Management Binder"
    - "{{doc_title}}"
  ref_label: "Packet Ref #"
  vars:
    org_short: North Star
  footer:
    left: ""
    center: ""
    right: ""
  page_numbers: per_document
  filenames:
    toc: TOC_m                       # legacy name 01_TOC_m.xlsx preserved
  tabs:
    product: "Avery 658843"
    template: Tabs.docx
    label_style: letter_and_label
    fit_limit: 24
  cover:
    template: 00_Cover.docx
    lines: ["North Star Media", "(DCCS Social Media Outreach)", "Management Binder"]
tree:
  - id: control
    kind: section
    label: Control Section
    numeric: true
    filename: Cover_Sheet_CS         # legacy cover name 00_Cover_Sheet_CS.docx
    description: "Used to communicate with your packet reviewer and keep track of your time."
    children:
      - {id: open-points, kind: item, title: "Open Points / Action Items", ext: xlsx, filename: Open_Points}
      - {id: notes-to-reviewer, kind: item, title: "Notes to Reviewer", ext: xlsx}
      - {id: time-sheet, kind: item, title: "Time Sheet", ext: xlsx}
  - id: administration
    kind: section
    label: Administration
    children:
      - {id: mrp-summary, kind: item, title: "North Star MRP Summary", ext: docx}
      - {id: org-chart, kind: item, title: "Organizational Chart", ext: docx}
      - {id: top-10-list, kind: item, title: "Top 10 List", ext: docx}
      - {id: top-10-list-backup, kind: item, title: "Top 10 List Backup", ext: docx, suffix: PB}
      - {id: points-of-contact, kind: item, title: "Points of Contact", ext: docx}
      - {id: budget-overview, kind: item, title: "Budget Overview", ext: xlsx}
      - {id: meeting-minutes, kind: item, title: "Meeting Minutes", ext: docx}
      - {id: outreach-efforts, kind: item, title: "Social Media Outreach Efforts", ext: docx}
  - id: policies
    kind: section
    label: Policies and Procedures
    children:
      - {id: media-policy, kind: item, title: "Media Policy", ext: docx}
      - {id: signed-acknowledgement, kind: item, title: "Signed Policy Acknowledgement", ext: pdf}
      - {id: licensed-release-form, kind: item, title: "Licensed Release Form", ext: docx, locked_header: true}
      - {id: retired-policy, kind: item, title: "Retired Policy", ext: docx, included: false}
  - id: campaign-management
    kind: section
    label: Campaign Management
    children:
      - {id: campaign-calendar, kind: item, title: "Campaign Calendar", ext: xlsx}
      - {id: campaign-briefs, kind: item, title: "Campaign Briefs", ext: docx}
  - id: content-mediums
    kind: header                     # clarifying header: no letter, no folder, no tab
    label: Content Mediums
    children:
      - id: website-media
        kind: section
        label: Website Media
        children:
          - id: web-index
            kind: item
            title: "Index of Genres & Channels"
            ext: xlsx
            children:
              - {id: weekly-digest, kind: item, title: "Weekly Digest Monthly Summary", ext: xlsx}
          - {id: website-screenshots, kind: item, title: "Website Screenshots", ext: docx}
      - id: print-media
        kind: section
        label: Print Media
        children:
          - {id: print-inventory, kind: item, title: "Print Inventory", ext: xlsx}
"""

# (relative path, code, title, ext, extras)
DOCS = [
    ("0_Control_Section/00_Cover_Sheet_CS.docx", "0.00", "Control Section", {}),
    ("0_Control_Section/01_Open_Points.xlsx", "0.01", "Open Points / Action Items",
     {"rows": [("Item", "From ref", "Owner", "Status"),
               ("Placeholder action one", "=bb_top_10_list", "Placeholder", "Open"),
               ("Placeholder action two", "=bb_mrp_summary", "Placeholder", "Closed")]}),
    ("0_Control_Section/02_Notes_to_Reviewer.xlsx", "0.02", "Notes to Reviewer",
     {"rows": [("Date", "Note"), ("2026-01-01", "Placeholder note to reviewer.")]}),
    ("0_Control_Section/03_Time_Sheet.xlsx", "0.03", "Time Sheet",
     {"rows": [("Date", "Hours", "Task", "Ref"), ("2026-01-02", 2.5, "Placeholder task", "=bb_open_points")]}),
    ("A_Administration/00_Cover.docx", "A.00", "Administration", {}),
    ("A_Administration/01_North_Star_MRP_Summary.docx", "A.01", "North Star MRP Summary",
     {"refs": [("Complete the Top 10 List (", "top-10-list", "A.03", ") first, then the outreach summary in "),
               ("", "outreach-efforts", "A.07", ".")], "table": True}),
    ("A_Administration/02_Organizational_Chart.docx", "A.02", "Organizational Chart", {}),
    ("A_Administration/03_Top_10_List.docx", "A.03", "Top 10 List", {"label": "Ref #:"}),
    ("A_Administration/03_PB_Top_10_List_Backup.docx", "A.03.PB", "Top 10 List Backup", {}),
    ("A_Administration/04_Points_of_Contact.docx", "A.04", "Points of Contact", {}),
    ("A_Administration/05_Budget_Overview.xlsx", "A.05", "Budget Overview",
     {"rows": [("Line", "Amount"), ("Placeholder line", 100), ("Total", "=SUM(B2:B2)")]}),
    ("A_Administration/06_Meeting_Minutes.docx", "A.06", "Meeting Minutes", {"first_page": True}),
    ("A_Administration/07_Social_Media_Outreach_Efforts.docx", "A.07", "Social Media Outreach Efforts",
     {"refs": [("See ", "top-10-list-backup", "A.03.PB", " for the backup list.")]}),
    ("B_Policies_and_Procedures/00_Cover.docx", "B.00", "Policies and Procedures", {}),
    ("B_Policies_and_Procedures/01_Media_Policy.docx", "B.01", "Media Policy", {}),
    ("B_Policies_and_Procedures/02_Signed_Policy_Acknowledgement.pdf", "B.02", "Signed Policy Acknowledgement", {}),
    ("B_Policies_and_Procedures/03_Licensed_Release_Form.docx", "B.03", "Licensed Release Form",
     {"locked_text": "LICENSED FORM - HEADER MUST NOT CHANGE"}),
    ("B_Policies_and_Procedures/04_Retired_Policy.docx", "B.04", "Retired Policy", {}),
    ("C_Campaign_Management/00_Cover.docx", "C.00", "Campaign Management", {}),
    ("C_Campaign_Management/01_Campaign_Calendar.xlsx", "C.01", "Campaign Calendar",
     {"rows": [("Week", "Campaign", "See"), (1, "Placeholder campaign", "=bb_campaign_briefs")]}),
    ("C_Campaign_Management/02_Campaign_Briefs.docx", "C.02", "Campaign Briefs",
     {"refs": [("Channels are indexed in ", "web-index", "D.01", " and summarised in "),
               ("", "weekly-digest", "D.01.01", ".")]}),
    ("D_Website_Media/00_Cover.docx", "D.00", "Website Media", {}),
    ("D_Website_Media/01_Index_of_Genres_and_Channels.xlsx", "D.01", "Index of Genres & Channels",
     {"rows": [("Genre", "Channel"), ("Placeholder genre", "Placeholder channel")]}),
    ("D_Website_Media/01_01_Weekly_Digest_Monthly_Summary.xlsx", "D.01.01", "Weekly Digest Monthly Summary",
     {"rows": [("Month", "Digests"), ("January", 4)]}),
    ("D_Website_Media/02_Website_Screenshots.docx", "D.02", "Website Screenshots", {}),
    ("E_Print_Media/00_Cover.docx", "E.00", "Print Media", {}),
    ("E_Print_Media/01_Print_Inventory.xlsx", "E.01", "Print Inventory",
     {"rows": [("Piece", "Count"), ("Placeholder flyer", 250)]}),
]


def _ref_sdt(target: str, text: str):
    sdt = OxmlElement("w:sdt")
    pr = OxmlElement("w:sdtPr")
    tag = OxmlElement("w:tag")
    tag.set(qn("w:val"), f"bb:ref:{target}")
    pr.append(tag)
    sdt.append(pr)
    content = OxmlElement("w:sdtContent")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    content.append(r)
    sdt.append(content)
    return sdt


def _legacy_word_header(section_hf, code: str, title: str, label: str = "Packet Ref #") -> None:
    section_hf.is_linked_to_previous = False
    paras = section_hf.paragraphs
    first = paras[0]
    first.text = f"North Star\t\t{label} {code}"
    for line in ("North Star Media", "Management Binder", title):
        section_hf.add_paragraph(line)


def make_docx(path: Path, code: str, title: str, extras: dict, legacy: bool = False) -> None:
    doc = Document()
    sec = doc.sections[0]
    if "locked_text" in extras:
        sec.header.is_linked_to_previous = False
        sec.header.paragraphs[0].text = extras["locked_text"]
        sec.footer.is_linked_to_previous = False
        sec.footer.paragraphs[0].text = "Form 123 (placeholder) - do not alter"
    else:
        _legacy_word_header(sec.header, code, title, extras.get("label", "Packet Ref #"))
        if extras.get("first_page"):
            sec.different_first_page_header_footer = True
            _legacy_word_header(sec.first_page_header, code, title)
    doc.add_heading(title, level=1)
    doc.add_paragraph(LOREM)
    if "refs" in extras:
        p = doc.add_paragraph()
        for before, target, text, after in extras["refs"]:
            if before:
                p.add_run(before)
            if legacy:
                p.add_run(text)  # typed by hand: not yet a live reference
            else:
                p._p.append(_ref_sdt(target, text))
            if after:
                p.add_run(after)
    if extras.get("table"):
        t = doc.add_table(rows=2, cols=2)
        for i, row in enumerate(t.rows):
            for j, cell in enumerate(row.cells):
                cell.text = f"Placeholder cell {i + 1}.{j + 1}"
    doc.add_paragraph("Second placeholder paragraph with\ta tab and a line break.")
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


def _legacy_excel_header(ws, code: str, title: str) -> None:
    ws.oddHeader.center.text = "North Star Media\nManagement Binder\n" + title.replace("&", "&&")
    ws.oddHeader.right.text = f"Packet Ref # {code}\nPage &P of &N"


def make_xlsx(path: Path, code: str, title: str, extras: dict, legacy: bool = False) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = title.replace("/", "-")[:31]
    for row in extras.get("rows", [("Placeholder",)]):
        if legacy:
            row = [LEGACY_NAME_CODES.get(v[1:], v) if isinstance(v, str) and v.startswith("=bb_") else v for v in row]
        ws.append(list(row))
    _legacy_excel_header(ws, code, title)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))


def make_pdf(path: Path, title: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(path), pagesize=letter, invariant=1)
    c.drawString(72, 700, f"{title} (placeholder scan)")
    c.showPage()
    c.save()


def make_root(out: Path) -> None:
    # Binder cover
    doc = Document()
    for line in ("North Star Media", "(DCCS Social Media Outreach)", "Management Binder"):
        doc.add_heading(line, level=0)
    doc.save(str(out / "00_Cover.docx"))

    # Legacy TOC: X | code | title | sub-item title. Control codes stored as floats, as in the samples.
    wb = Workbook()
    ws = wb.active
    ws.title = "TOC"
    ws.append(["", "", "Table of Contents"])
    ws.append(["", 0, "Control Section"])
    ws.append(["X", 0.01, "Open Points / Action Items"])
    ws.append(["X", 0.02, "Notes to Reviewer"])
    ws.append(["X", 0.03, "Time Sheet"])
    ws.append(["", "A", "Administration"])
    for code, title in [("A.01", "North Star MRP Summary"), ("A.02", "Organizational Chart"),
                        ("A.03", "Top 10 List"), ("A.03.PB", "Top 10 List Backup"),
                        ("A.04", "Points of Contact"), ("A.05", "Budget Overview"),
                        ("A.06", "Meeting Minutes"), ("A.07", "Social Media Outreach Efforts")]:
        ws.append(["X", code, title])
    ws.append(["", "B", "Policies and Procedures"])
    ws.append(["X", "B.01", "Media Policy"])
    ws.append(["X", "B.02", "Signed Policy Acknowledgement"])
    ws.append(["X", "B.03", "Licensed Release Form"])
    ws.append(["", "B.04", "Retired Policy"])
    ws.append(["", "C", "Campaign Management"])
    ws.append(["X", "C.01", "Campaign Calendar"])
    ws.append(["X", "C.02", "Campaign Briefs"])
    ws.append(["", "", "Content Mediums"])
    ws.append(["", "D", "Website Media"])
    ws.append(["X", "D.01", "Index of Genres & Channels"])
    ws.append(["X", "", "D.01.01", "Weekly Digest Monthly Summary"])
    ws.append(["X", "D.02", "Website Screenshots"])
    ws.append(["", "E", "Print Media"])
    ws.append(["X", "E.01", "Print Inventory"])
    _legacy_excel_header(ws, "TOC", "Table of Contents")
    wb.save(str(out / "01_TOC_m.xlsx"))

    # Legacy Master Checklist: references as plain text, user columns filled in.
    wb = Workbook()
    ws = wb.active
    ws.title = "Master Checklist"
    ws.append(["Reference", "Description", "Hours", "Reviewer", "Date Reviewed", "Final Review"])
    for code, title in [("0.01", "Open Points / Action Items"), ("A.01", "North Star MRP Summary"),
                        ("A.03", "Top 10 List"), ("B.01", "Media Policy"), ("D.01", "Index of Genres & Channels")]:
        ws.append([code, title, 1.0, "Placeholder Reviewer", "2026-02-01", ""])
    _legacy_excel_header(ws, "CKL", "Master Checklist")
    wb.save(str(out / "02_CKL.xlsx"))

    # Tabs: a table of label cells.
    doc = Document()
    labels = ["0) Control Section", "A) Administration", "B) Policies and Procedures",
              "C) Campaign Management", "D) Website Media", "E) Print Media"]
    t = doc.add_table(rows=len(labels), cols=1)
    for row, label in zip(t.rows, labels):
        row.cells[0].text = label
    doc.save(str(out / "Tabs.docx"))


def build(out: Path = OUT) -> Path:
    if out.exists():
        shutil.rmtree(out)  # regenerating the fixture's own output folder
    (out / ".binder").mkdir(parents=True)
    (out / ".binder" / "binder.yaml").write_text(MANIFEST, encoding="utf-8")
    make_root(out)
    for rel, code, title, extras in DOCS:
        p = out / rel
        if p.suffix == ".docx":
            make_docx(p, code, title, extras)
        elif p.suffix == ".xlsx":
            make_xlsx(p, code, title, extras)
        else:
            make_pdf(p, title)
    return out


def build_legacy(out: Path = LEGACY_OUT) -> Path:
    """The same binder before Binder Builder ever saw it (input for the import tests)."""
    if out.exists():
        shutil.rmtree(out)  # regenerating the fixture's own output folder
    out.mkdir(parents=True)
    make_root(out)
    for rel, code, title, extras in DOCS:
        code = LEGACY_STALE_HEADER.get(rel, code)
        p = out / LEGACY_RENAMES.get(rel, rel)
        if p.suffix == ".docx":
            make_docx(p, code, title, extras, legacy=True)
        elif p.suffix == ".xlsx":
            make_xlsx(p, code, title, extras, legacy=True)
        else:
            make_pdf(p, title)
    return out


if __name__ == "__main__":
    print(f"wrote {build(OUT)}")
    print(f"wrote {build_legacy(LEGACY_OUT)}")
