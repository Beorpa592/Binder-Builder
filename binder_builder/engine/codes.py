"""Tree position -> codes, folders, filenames and tab labels (SPEC §3).

Nothing here touches the disk. `compute_layout` is a pure function of the tree
and binder settings; it is recomputed on every sync, so codes are never stored
as the source of truth.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from posixpath import join as pjoin

from .manifest import BinderModel, FooterModel
from .tree import Node, Tree

# Entity kinds that are documents on disk.
ITEM = "item"
SECTION_COVER = "section_cover"
BINDER_COVER = "binder_cover"
TOC = "toc"
CKL = "ckl"
TABS = "tabs"


def filename_part(title: str) -> str:
    """Title -> filename title part, exactly as SPEC §3.1."""
    s = re.sub(r"\s+", "_", title)
    s = s.replace("&", "and")
    s = re.sub(r"[^A-Za-z0-9_-]", "", s)
    s = re.sub(r"_+", "_", s)
    return s.strip("_")


def letter_sequence():
    """A, B, ..., Z, AA, AB, ..., AZ, BA, ... (spreadsheet-style)."""
    n = 0
    while True:
        n += 1
        k, s = n, ""
        while k:
            k, r = divmod(k - 1, 26)
            s = chr(65 + r) + s
        yield s


@dataclass
class Entity:
    """One document the binder owns a position for."""

    entity_id: str
    kind: str
    rel_path: str  # posix, relative to the binder folder
    ext: str
    code: str | None
    title: str
    section_label: str | None = None
    node: Node | None = None
    included: bool = True
    locked_header: bool = False
    footer: FooterModel = field(default_factory=FooterModel)
    stamp: bool = True  # False: the app never writes a header/footer for this entity

    @property
    def filename(self) -> str:
        return self.rel_path.rsplit("/", 1)[-1]


@dataclass
class SectionInfo:
    node: Node
    code: str
    folder: str  # posix, relative to binder folder
    depth: int
    tab_label: str | None
    parent: "SectionInfo | None" = None


@dataclass
class Layout:
    entities: list[Entity]
    codes: dict[str, str]  # node id -> code, for items and sections
    sections: list[SectionInfo]
    warnings: list[str]

    def entity(self, entity_id: str) -> Entity | None:
        for e in self.entities:
            if e.entity_id == entity_id:
                return e
        return None

    def items(self) -> list[Entity]:
        return [e for e in self.entities if e.kind == ITEM]


def section_cover_id(section_id: str) -> str:
    return f"{section_id}#cover"


def tab_label(code: str, label: str, binder: BinderModel) -> tuple[str, str | None]:
    """Returns (label, warning or None) per SPEC §8 Tabs."""
    limit = binder.tabs.fit_limit
    full = f"{code}) {label}"
    if binder.tabs.label_style == "letter_and_label" and len(full) <= limit:
        return full, None
    if len(label) <= limit:
        return label, None
    return label, f"tab label {label!r} is {len(label)} characters, over the fit limit of {limit}"


def _footer_for(node: Node | None, binder: BinderModel) -> FooterModel:
    if node is not None and node.get("footer"):
        return FooterModel.model_validate(dict(node.get("footer")))
    return binder.footer


def compute_layout(tree: Tree, binder: BinderModel) -> Layout:
    entities: list[Entity] = []
    codes: dict[str, str] = {}
    sections: list[SectionInfo] = []
    warnings: list[str] = []
    letters = letter_sequence()
    names = binder.filenames

    entities.append(
        Entity("#binder-cover", BINDER_COVER, f"00_{names.cover}.docx", "docx", None,
               binder.name, stamp=False, footer=binder.footer)
    )
    entities.append(Entity("#toc", TOC, f"01_{names.toc}.xlsx", "xlsx", "TOC", "Table of Contents",
                           footer=binder.footer))
    entities.append(Entity("#ckl", CKL, f"02_{names.ckl}.xlsx", "xlsx", "CKL", "Master Checklist",
                           footer=binder.footer))

    def number_items(children: list[Node], code_prefix: str, parts: list[str], folder: str,
                     section: SectionInfo) -> None:
        ordinals = [c for c in children if c.is_item and not c.get("suffix")]
        width = 3 if len(ordinals) > 99 else 2
        n = 0
        prev: tuple[str, list[str]] | None = None
        for c in children:
            if not c.is_item:
                if c.kind == "section" and section is not None:
                    continue  # handled by walk_sections
                warnings.append(f"{c.kind} {c.id!r} cannot sit inside an item; ignored")
                continue
            suffix = c.get("suffix")
            if suffix:
                if prev is None:
                    warnings.append(f"suffixed item {c.id!r} has no preceding sibling to hang off")
                    prev = (f"{code_prefix}.{'0' * width}", parts + ["0" * width])
                code = f"{prev[0]}.{suffix}"
                my_parts = prev[1] + [suffix]
            else:
                n += 1
                nn = str(n).zfill(width)
                code = f"{code_prefix}.{nn}"
                my_parts = parts + [nn]
                prev = (code, my_parts)
            codes[c.id] = code
            title_part = c.get("filename") or filename_part(c.text)
            fname = "_".join(my_parts) + f"_{title_part}.{c.get('ext')}"
            entities.append(
                Entity(
                    c.id, ITEM, pjoin(folder, fname) if folder else fname, c.get("ext"), code, c.text,
                    section_label=section.node.text if section else None, node=c,
                    included=c.included and all(a.included for a in c.ancestors()),
                    locked_header=bool(c.get("locked_header", False)),
                    footer=_footer_for(c, binder),
                )
            )
            if c.children:
                number_items(c.children, code, my_parts, folder, section)

    def walk(nodes: list[Node], folder: str, parent: SectionInfo | None, depth: int) -> None:
        for node in nodes:
            if node.is_header:
                if any(c.is_item for c in node.children):
                    warnings.append(f"clarifying header {node.id!r} has items of its own; they are ignored")
                walk([c for c in node.children if not c.is_item], folder, parent, depth)
            elif node.is_section:
                code = "0" if node.get("numeric") else next(letters)
                codes[node.id] = code
                folder_part = node.get("folder") or filename_part(node.text)
                sec_folder = pjoin(folder, f"{code}_{folder_part}") if folder else f"{code}_{folder_part}"
                label, warn = tab_label(code, node.text, binder)
                if warn:
                    warnings.append(warn)
                info = SectionInfo(node, code, sec_folder, depth, label, parent)
                sections.append(info)
                cover_part = node.get("filename") or "Cover"
                entities.append(
                    Entity(
                        section_cover_id(node.id), SECTION_COVER, f"{sec_folder}/00_{cover_part}.docx", "docx",
                        f"{code}.00", node.text, section_label=node.text, node=node,
                        included=node.included and all(a.included for a in node.ancestors()),
                        locked_header=bool(node.get("locked_header", False)),
                        footer=_footer_for(node, binder),
                    )
                )
                number_items(node.children, code, [], sec_folder, info)
                walk([c for c in node.children if not c.is_item], sec_folder, info, depth + 1)
            else:
                warnings.append(f"item {node.id!r} is not inside a section; it gets no position")

    walk(tree.roots, "", None, 0)

    entities.append(Entity("#tabs", TABS, f"{names.tabs}.docx", "docx", None, "Tabs", stamp=False,
                           footer=binder.footer))
    return Layout(entities, codes, sections, warnings)


# ---- header text (SPEC §5) ------------------------------------------------

_PLACEHOLDER = re.compile(r"\{\{\s*([A-Za-z0-9_]+)\s*\}\}")


@dataclass(frozen=True)
class HeaderSpec:
    """What the app-owned header and footer of one document should say."""

    lines: tuple[str, ...]
    code: str | None
    ref_label: str
    footer: tuple[str, str, str]  # left, center, right
    page_numbers: bool = True


def render_text(template: str, context: dict[str, str]) -> str:
    return _PLACEHOLDER.sub(lambda m: context.get(m.group(1), m.group(0)), template)


def header_spec(entity: Entity, binder: BinderModel) -> HeaderSpec:
    context = dict(binder.vars)
    context.update(
        binder_name=binder.name,
        doc_title=entity.title,
        doc_code=entity.code or "",
        section_label=entity.section_label or "",
    )
    lines = tuple(render_text(line, context) for line in binder.header_lines)
    f = entity.footer
    footer = tuple(render_text(x, context) for x in (f.left, f.center, f.right))
    return HeaderSpec(lines, entity.code, binder.ref_label, footer)  # type: ignore[arg-type]
