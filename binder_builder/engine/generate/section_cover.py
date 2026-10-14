"""Section cover sheet (SPEC §8): `<Letter>) <Label>`, description, then `__X__  code  title` per child."""

from __future__ import annotations

from pathlib import Path

from docxtpl import DocxTemplate

from ..codes import Entity
from ..tree import Node

MARK_IN = "__X__"
MARK_OUT = "__ __"


def rows_for(section: Node, ctx) -> list[dict]:
    """Every item in the section (sub-items included, in order) and each direct sub-section."""
    layout = ctx.layout
    rows = []

    def add(node: Node) -> None:
        code = layout.codes.get(node.id)
        if code is None:
            return
        e = layout.entity(node.id)
        included = e.included if e else node.included
        rows.append({"mark": MARK_IN if included else MARK_OUT, "code": code, "title": node.text,
                     "kind": node.kind})

    def walk(nodes: list[Node]) -> None:
        for n in nodes:
            if n.is_item:
                add(n)
                walk(n.children)
            elif n.is_section:
                add(n)
            elif n.is_header:
                walk([c for c in n.children if c.is_section])

    walk(section.children)
    return rows


def render(ctx, entity: Entity, template: Path, out: Path) -> list[str]:
    section = entity.node
    code = ctx.layout.codes[section.id]
    tpl = DocxTemplate(str(template))
    tpl.render({
        "heading": f"{code}) {section.text}",
        "code": code,
        "label": section.text,
        "description": section.get("description") or "",
        "rows": rows_for(section, ctx),
        "binder_name": ctx.binder.name,
    }, autoescape=True)
    tpl.save(str(out))
    return []
