"""Node model, traversal, edits and id slugging (SPEC §4).

`Node` wraps the raw round-trip mapping from the manifest, so every edit made
through `Tree` is written straight back into the YAML document and survives a
save with comments and ordering intact.
"""

from __future__ import annotations

import re
from typing import Any, Iterator

from ruamel.yaml.comments import CommentedMap, CommentedSeq

from .manifest import Manifest


def slugify(text: str) -> str:
    """Slug used for permanent node ids: lowercase, `&` -> and, hyphen separated."""
    s = text.replace("&", " and ").lower()
    s = re.sub(r"[^a-z0-9]+", "-", s)
    return s.strip("-") or "node"


class Node:
    def __init__(self, raw: CommentedMap, parent: "Node | None"):
        self.raw = raw
        self.parent = parent
        self.children: list[Node] = [Node(c, self) for c in raw.get("children", []) or []]

    # ---- field access -------------------------------------------------
    @property
    def id(self) -> str:
        return self.raw["id"]

    @property
    def kind(self) -> str:
        return self.raw["kind"]

    @property
    def is_section(self) -> bool:
        return self.kind == "section"

    @property
    def is_header(self) -> bool:
        return self.kind == "header"

    @property
    def is_item(self) -> bool:
        return self.kind == "item"

    @property
    def text(self) -> str:
        """Display text: label for sections/headers, title for items."""
        return self.raw.get("label") if self.kind != "item" else self.raw.get("title")

    def get(self, key: str, default: Any = None) -> Any:
        v = self.raw.get(key, default)
        return default if v is None else v

    def set(self, key: str, value: Any) -> None:
        if value is None:
            self.raw.pop(key, None)
        else:
            self.raw[key] = value

    @property
    def included(self) -> bool:
        return bool(self.get("included", True))

    def __repr__(self) -> str:
        return f"Node({self.id!r}, {self.kind})"

    # ---- structure ----------------------------------------------------
    def _children_seq(self) -> CommentedSeq:
        seq = self.raw.get("children")
        if seq is None:
            seq = CommentedSeq()
            self.raw["children"] = seq
        if not seq and isinstance(seq, CommentedSeq):
            seq.fa.set_block_style()
        return seq

    def walk(self) -> Iterator["Node"]:
        yield self
        for c in self.children:
            yield from c.walk()

    def ancestors(self) -> Iterator["Node"]:
        p = self.parent
        while p is not None:
            yield p
            p = p.parent


# Where each kind may sit: items inside sections (or under another item, as sub-items);
# sections at top level, under a clarifying header, or nested in a section; headers at
# top level or inside a section.
VALID_PARENTS = {
    "item": ("section", "item"),
    "section": (None, "header", "section"),
    "header": (None, "section"),
}


def can_place(kind: str, parent: "Node | None") -> bool:
    return (parent.kind if parent else None) in VALID_PARENTS[kind]


class Tree:
    """The binder tree. Top-level nodes have `parent=None`."""

    def __init__(self, manifest: Manifest):
        self.manifest = manifest
        seq = manifest.raw.get("tree")
        if seq is None:
            seq = CommentedSeq()
            manifest.raw["tree"] = seq
        if not seq:
            seq.fa.set_block_style()  # an empty `tree: []` must not make every node flow-style
        self._root_seq: CommentedSeq = seq
        self.roots: list[Node] = [Node(r, None) for r in seq]

    def walk(self) -> Iterator[Node]:
        for r in self.roots:
            yield from r.walk()

    def ids(self) -> set[str]:
        return {n.id for n in self.walk()}

    def find(self, node_id: str) -> Node:
        for n in self.walk():
            if n.id == node_id:
                return n
        raise KeyError(node_id)

    def get(self, node_id: str) -> Node | None:
        try:
            return self.find(node_id)
        except KeyError:
            return None

    # ---- id generation ------------------------------------------------
    def new_id(self, title: str) -> str:
        base = slugify(title)
        existing = self.ids()
        if base not in existing:
            return base
        i = 2
        while f"{base}-{i}" in existing:
            i += 1
        return f"{base}-{i}"

    # ---- edits --------------------------------------------------------
    def _siblings(self, parent: Node | None) -> tuple[list[Node], CommentedSeq]:
        if parent is None:
            return self.roots, self._root_seq
        return parent.children, parent._children_seq()

    def insert(self, data: dict, parent_id: str | None = None, index: int | None = None) -> Node:
        """Insert a new node. `data` needs kind and title/label; id is generated if absent."""
        raw = CommentedMap()
        if "id" not in data:
            raw["id"] = self.new_id(data.get("title") or data.get("label") or "node")
        for k, v in data.items():
            raw[k] = v
        if raw["id"] in self.ids():
            raise ValueError(f"duplicate id {raw['id']!r}")
        parent = self.find(parent_id) if parent_id else None
        nodes, seq = self._siblings(parent)
        if index is None:
            index = len(nodes)
        seq.insert(index, raw)
        node = Node(raw, parent)
        nodes.insert(index, node)
        self.manifest.validate()
        return node

    def insert_after(self, sibling_id: str, data: dict) -> Node:
        sib = self.find(sibling_id)
        nodes, _ = self._siblings(sib.parent)
        return self.insert(data, sib.parent.id if sib.parent else None, nodes.index(sib) + 1)

    def remove(self, node_id: str) -> Node:
        """Remove a node (and its subtree) from the TOC. Files are never touched."""
        node = self.find(node_id)
        nodes, seq = self._siblings(node.parent)
        i = nodes.index(node)
        del nodes[i]
        del seq[i]
        self.manifest.validate()
        return node

    def move(self, node_id: str, parent_id: str | None, index: int) -> Node:
        node = self.find(node_id)
        if parent_id is not None and parent_id in {n.id for n in node.walk()}:
            raise ValueError("cannot move a node into itself")
        if not can_place(node.kind, self.find(parent_id) if parent_id else None):
            where = f"inside {self.find(parent_id).kind} {parent_id!r}" if parent_id else "at the top level"
            raise ValueError(f"a {node.kind} cannot go {where}")
        if node.parent is (self.find(parent_id) if parent_id else None):
            current = self._siblings(node.parent)[0].index(node)
            if index > current:
                index -= 1  # removing the node first shifts later positions up
        old_nodes, old_seq = self._siblings(node.parent)
        i = old_nodes.index(node)
        del old_nodes[i]
        del old_seq[i]
        parent = self.find(parent_id) if parent_id else None
        nodes, seq = self._siblings(parent)
        index = max(0, min(index, len(nodes)))
        seq.insert(index, node.raw)
        nodes.insert(index, node)
        node.parent = parent
        self.manifest.validate()
        return node

    # ---- convenience edits used by the GUI ------------------------------
    def index_of(self, node: Node) -> int:
        return self._siblings(node.parent)[0].index(node)

    def move_by(self, node_id: str, delta: int) -> Node:
        node = self.find(node_id)
        nodes, _ = self._siblings(node.parent)
        i = nodes.index(node)
        j = max(0, min(len(nodes) - 1, i + delta))
        if i == j:
            return node
        target = j + 1 if j > i else j
        return self.move(node_id, node.parent.id if node.parent else None, target)

    def indent(self, node_id: str) -> Node:
        """Make the node the last child of its previous sibling."""
        node = self.find(node_id)
        nodes, _ = self._siblings(node.parent)
        i = nodes.index(node)
        if i == 0:
            raise ValueError("nothing above to indent under")
        new_parent = nodes[i - 1]
        return self.move(node_id, new_parent.id, len(new_parent.children))

    def outdent(self, node_id: str) -> Node:
        """Move the node out of its parent, to just after the parent."""
        node = self.find(node_id)
        parent = node.parent
        if parent is None:
            raise ValueError("already at the top level")
        grand = parent.parent
        idx = self.index_of(parent) + 1
        return self.move(node_id, grand.id if grand else None, idx)
