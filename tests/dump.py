"""Readable text dumps of Office files, used for golden-file comparison.

Dumps hold what a reader sees: paragraph text, table cells, header/footer
content (app controls shown as <tag>text</>), cell values, number formats of
code cells, hidden columns/sheets and defined names. They leave out zip
timestamps and XML noise, so goldens stay stable and reviewable in diffs.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from binder_builder.engine import excel, word


def dump_docx(path: Path) -> str:
    doc = Document(str(path))
    out = []
    sec = doc.sections[0]
    for name, hf in (("HEADER", sec.header), ("FOOTER", sec.footer)):
        if not hf.is_linked_to_previous:
            out.append(f"[{name}]")
            out.extend(word._part_signature(hf._element).rstrip("\n").split("\n"))
    out.append("[BODY]")
    body = doc.element.body
    for child in body:
        if child.tag == qn("w:p"):
            text = word._part_signature(child).rstrip("\n")
            style = child.find(qn("w:pPr") + "/" + qn("w:pStyle"))
            prefix = f"({style.get(qn('w:val'))}) " if style is not None else ""
            if text or prefix:
                out.append(prefix + text)
            if child.findall(".//" + qn("w:br")) and any(
                    b.get(qn("w:type")) == "page" for b in child.iter(qn("w:br"))):
                out.append("<page break>")
        elif child.tag == qn("w:tbl"):
            out.append("<table>")
            for tr in child.iter(qn("w:tr")):
                cells = [word._part_signature(tc).strip().replace("\n", " / ") for tc in tr.iter(qn("w:tc"))]
                out.append("| " + " | ".join(cells) + " |")
            out.append("</table>")
    return "\n".join(out) + "\n"


def dump_xlsx(path: Path) -> str:
    wb = excel.load(path)
    out = []
    for ws in wb.worksheets:
        out.append(f"[SHEET {ws.title}] state={ws.sheet_state}")
        out.append(f"header={tuple(p.text for p in (ws.oddHeader.left, ws.oddHeader.center, ws.oddHeader.right))}")
        hidden = [k for k, d in sorted(ws.column_dimensions.items()) if d.hidden]
        if hidden:
            out.append(f"hidden columns={hidden}")
        for row in ws.iter_rows():
            vals = [(c.coordinate, c.value, c.number_format) for c in row if c.value is not None]
            if vals:
                out.append("  ".join(f"{k}={v!r}" + (" @" if f == "@" else "") for k, v, f in vals))
    for name, dn in sorted(wb.defined_names.items()):
        out.append(f"NAME {name} -> {dn.attr_text}")
    return "\n".join(out) + "\n"


def dump(path: Path) -> str:
    return dump_docx(path) if path.suffix == ".docx" else dump_xlsx(path)
