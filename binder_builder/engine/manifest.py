"""Load, validate, upgrade and save `.binder/binder.yaml` (SPEC §4).

The manifest is kept as a ruamel.yaml round-trip document so comments and key
order survive edits. Pydantic models validate it; the engine reads node data
through `tree.Node`, which wraps the raw mappings so edits land back in the
round-trip document.
"""

from __future__ import annotations

import io
from pathlib import Path
from typing import Any, Callable, Literal

from pydantic import BaseModel, ConfigDict, Field, ValidationError, model_validator
from ruamel.yaml import YAML
from ruamel.yaml.comments import CommentedMap

FORMAT_VERSION = 1
BINDER_DIR = ".binder"
MANIFEST_NAME = "binder.yaml"

DEFAULT_REF_LABEL = "Packet Ref #"
DEFAULT_TAB_FIT_LIMIT = 24
DEFAULT_LOCK_MAX_AGE_HOURS = 4


class ManifestError(Exception):
    pass


class _Strict(BaseModel):
    model_config = ConfigDict(extra="forbid")


class FooterModel(_Strict):
    left: str = ""
    center: str = ""
    right: str = ""

    def is_blank(self) -> bool:
        return not (self.left or self.center or self.right)


class CopiedFrom(_Strict):
    binder_path: str
    id: str
    synced_at: str
    source_hash: str | None = None
    source_path: str | None = None  # the exact file copied, for the drift check (SPEC §9.3)


class NodeModel(_Strict):
    id: str
    kind: Literal["section", "header", "item"]
    label: str | None = None
    title: str | None = None
    ext: Literal["docx", "xlsx", "pdf"] | None = None
    included: bool = True
    suffix: str | None = None
    filename: str | None = None
    folder: str | None = None
    locked_header: bool = False
    footer: FooterModel | None = None
    copied_from: CopiedFrom | None = None
    numeric: bool = False
    description: str | None = None
    hybrid_rows: bool = False
    hybrid_columns: list[str] | None = None
    children: list["NodeModel"] = Field(default_factory=list)

    @model_validator(mode="after")
    def _check_kind(self) -> "NodeModel":
        if self.kind in ("section", "header") and not self.label:
            raise ValueError(f"{self.kind} {self.id!r} needs a label")
        if self.kind == "item":
            if not self.title:
                raise ValueError(f"item {self.id!r} needs a title")
            if not self.ext:
                raise ValueError(f"item {self.id!r} needs an ext")
        return self


class TabsModel(_Strict):
    product: str = "Avery 658843"
    template: str = "Tabs.docx"
    label_style: Literal["letter_and_label", "label_only"] = "letter_and_label"
    fit_limit: int = DEFAULT_TAB_FIT_LIMIT


class CoverModel(_Strict):
    template: str = "00_Cover.docx"
    lines: list[str] = Field(default_factory=list)


class RootFilenames(_Strict):
    """Title-part overrides for the binder root documents (see DEVIATIONS.md)."""

    cover: str = "Cover"
    toc: str = "TOC"
    ckl: str = "CKL"
    tabs: str = "Tabs"


class BinderModel(_Strict):
    name: str
    header_lines: list[str] = Field(default_factory=list)
    ref_label: str = DEFAULT_REF_LABEL
    vars: dict[str, str] = Field(default_factory=dict)
    footer: FooterModel = Field(default_factory=FooterModel)
    page_numbers: Literal["per_document", "continuous"] = "per_document"
    tabs: TabsModel = Field(default_factory=TabsModel)
    cover: CoverModel = Field(default_factory=CoverModel)
    filenames: RootFilenames = Field(default_factory=RootFilenames)
    template_folder: str | None = None
    lock_max_age_hours: float = DEFAULT_LOCK_MAX_AGE_HOURS


class ManifestModel(_Strict):
    format_version: int
    binder: BinderModel
    tree: list[NodeModel] = Field(default_factory=list)


def _yaml() -> YAML:
    y = YAML()  # round-trip mode
    y.indent(mapping=2, sequence=4, offset=2)
    y.width = 4096
    y.preserve_quotes = True
    return y


def manifest_path(folder: Path) -> Path:
    return Path(folder) / BINDER_DIR / MANIFEST_NAME


# Schema upgrades: UPGRADES[n] turns a format_version n manifest into n + 1, editing the
# round-trip document in place (so comments survive). Bump FORMAT_VERSION and register a step
# whenever the schema changes; old binders are upgraded the next time they are opened.
UPGRADES: dict[int, Callable[[CommentedMap], None]] = {}


def upgrade(raw: CommentedMap) -> int | None:
    """Upgrade an older manifest in place. Returns the original version if it changed, else None."""
    version = raw.get("format_version")
    if version is None:
        raise ManifestError("manifest has no format_version")
    if not isinstance(version, int) or isinstance(version, bool):
        raise ManifestError(f"format_version must be an integer, got {version!r}")
    if version > FORMAT_VERSION:
        raise ManifestError(
            f"this binder was saved by a newer Binder Builder (format {version}; this version reads up to "
            f"{FORMAT_VERSION}). Update Binder Builder to open it."
        )
    original = version
    while version < FORMAT_VERSION:
        step = UPGRADES.get(version)
        if step is None:
            raise ManifestError(f"no upgrade path from format_version {version}")
        step(raw)
        version += 1
        raw["format_version"] = version
    return original if version != original else None


NEW_BINDER_TEMPLATE = """# Binder Builder manifest. Edit through the app; hand edits are allowed but validated.
format_version: 1
binder:
  name: New Binder
  header_lines:                      # author-defined, propagated to every document
    - "{{org_short}}"                # line 1 left; the ref code goes line 1 right automatically
    - "{{binder_name}}"
    - "{{doc_title}}"
  ref_label: "Packet Ref #"
  vars:
    org_short: ""
  footer:
    left: ""
    center: ""
    right: ""
  page_numbers: per_document
  cover:
    lines: []
tree:
"""


class Manifest:
    """A loaded manifest: the round-trip document plus its validated model."""

    def __init__(self, raw: CommentedMap, path: Path | None = None):
        self.raw = raw
        self.path = path
        self.model = self.validate()

    # ---- construction -------------------------------------------------
    @classmethod
    def load(cls, folder: Path | str) -> "Manifest":
        path = manifest_path(Path(folder))
        if not path.exists():
            raise ManifestError(f"no manifest at {path}")
        text = path.read_text(encoding="utf-8")
        raw = _yaml().load(text)
        if not isinstance(raw, CommentedMap):
            raise ManifestError(f"{path} is not a YAML mapping")
        upgraded_from = upgrade(raw)
        m = cls(raw, path)
        if upgraded_from is not None:
            backup = path.with_name(f"binder.v{upgraded_from}.yaml.bak")
            if not backup.exists():
                backup.write_text(text, encoding="utf-8")  # the original, before the upgrade
            m.save()
        return m

    @classmethod
    def create(cls, folder: Path | str, name: str) -> "Manifest":
        """Write a new, empty binder manifest. Refuses to overwrite an existing one."""
        path = manifest_path(Path(folder))
        if path.exists():
            raise ManifestError(f"{path} already exists")
        raw = _yaml().load(NEW_BINDER_TEMPLATE)
        raw["binder"]["name"] = name
        raw["binder"]["vars"]["org_short"] = name
        raw["binder"]["cover"]["lines"] = [name]
        from ruamel.yaml.comments import CommentedSeq

        raw["tree"] = CommentedSeq()
        m = cls(raw, path)
        m.save()
        return m

    @classmethod
    def loads(cls, text: str) -> "Manifest":
        raw = _yaml().load(text)
        upgrade(raw)
        return cls(raw)

    # ---- validation ---------------------------------------------------
    def validate(self) -> ManifestModel:
        try:
            model = ManifestModel.model_validate(_plain(self.raw))
        except ValidationError as e:
            raise ManifestError(f"invalid manifest: {e}") from e
        seen: set[str] = set()
        dupes: list[str] = []

        def walk(nodes: list[NodeModel]) -> None:
            for n in nodes:
                if n.id in seen:
                    dupes.append(n.id)
                seen.add(n.id)
                walk(n.children)

        walk(model.tree)
        if dupes:
            raise ManifestError(f"duplicate node ids: {', '.join(sorted(set(dupes)))}")
        self.model = model
        return model

    @property
    def binder(self) -> BinderModel:
        return self.model.binder

    # ---- saving -------------------------------------------------------
    def dumps(self) -> str:
        buf = io.StringIO()
        _yaml().dump(self.raw, buf)
        return buf.getvalue()

    def save(self, path: Path | None = None) -> None:
        self.validate()
        target = Path(path) if path else self.path
        if target is None:
            raise ManifestError("no path to save manifest to")
        target.parent.mkdir(parents=True, exist_ok=True)
        tmp = target.with_suffix(".yaml.tmp")
        tmp.write_text(self.dumps(), encoding="utf-8")
        tmp.replace(target)
        self.path = target


def _plain(obj: Any) -> Any:
    """Convert ruamel containers to plain dicts/lists for pydantic."""
    if isinstance(obj, dict):
        return {str(k): _plain(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_plain(v) for v in obj]
    return obj
