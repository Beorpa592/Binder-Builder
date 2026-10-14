# Binder Builder: Code-Ready Specification

Version 0.1, 2026-10-14. Owner: Ben. This file is the source of truth for the coding agent. Human-readable rationale lives in the "Binder Builder Spec" doc; where they differ, this file wins.

---

## 1. What this is

A Windows desktop application (full GUI, no terminal) that manages physical binders made of Word and Excel documents. The user edits a Table of Contents inside the app; the app keeps every document's header, reference code, filename, folder, checklist row, and tab label in sync with that TOC, working **in place** on the binder's folder on the office shared drive.

The one rule everything depends on: **every document has a permanent internal ID; its reference code (A.07) is computed from its position in the TOC tree every time the app syncs.** No code is ever stored as the source of truth. Move a document, sync, and everything that pointed at it follows.

### 1.1 Non-negotiable safety rules

1. **The app never deletes or rewrites body content an author put in a document.** It touches only: page headers, page footers, tagged reference fields (Word content controls, Excel defined names), the hidden `_binder` sheet in Excel files, and the generated documents it owns (TOC, section cover sheets, tabs sheet). Everything else in a file is invariant across a sync, and the app verifies that after every write (see §7.4).
2. **Every sync is previewed before it is applied.** The user sees the full change list (renames, header changes, new files, reference updates, warnings) and confirms.
3. **Every sync snapshots every file it is about to modify** into `.binder/backups/<timestamp>/` before writing. Restore is one click.
4. **The app never deletes a file.** Removing a TOC row moves nothing and deletes nothing; the file stays where it is and is listed as "orphaned" in the status panel. The user deletes files themselves in Explorer if they want to.
5. **Documents can be marked `locked_header: true`** and the app will not touch their header or footer at all (licensed forms, signed pages, scans).

---

## 2. Environment and delivery

| Item | Decision |
| --- | --- |
| OS | Windows 11 (Windows 10 tolerated) |
| Office | Microsoft Word and Excel are installed on every machine that runs the app |
| Python | Whatever version the dev uses; irrelevant to users because of packaging |
| Packaging | PyInstaller single-folder build, all dependencies bundled, one `BinderBuilder.exe`. Users never see Python, pip, or a terminal |
| Updates | GitHub Releases on the public repo. At startup the app checks the latest release, shows version and notes, downloads and swaps itself on confirm. Fallback: user downloads the zip from the Releases page |
| Repo | Public, personal GitHub account, name `binder-builder` |
| GUI | Full GUI from the start. v1 in tkinter + ttk (matches Packet Printer); engine is GUI-agnostic so a PySide front end can replace it later |
| PDF | Word and Excel COM automation (pywin32) for PDF export and page counts, since Office is always present. No LibreOffice dependency |
| Packet Printer | Separate project. Integration is a JSON print manifest the app writes and Packet Printer reads (§11) |

**Public repo caution:** no real binder content may be committed. Test fixtures are synthetic documents that mirror the structure of the North Star samples (same header layout, same TOC layout, same file naming) with placeholder text. A `fixtures/README.md` explains how to regenerate them.

---

## 3. Binder folder layout and naming

A binder is one folder. The app stores its state in a `.binder/` subfolder; everything else is the user's.

```
North Star Media Management Binder/
  .binder/
    binder.yaml            # the manifest (source of truth)
    binder.lock            # present while someone holds the structure lock
    sync.log               # append-only log of every sync
    state.json             # per-file hashes from the last sync and last print
    backups/
      2026-10-14T15-30-00/ # copies of files modified by that sync
  00_Cover.docx
  01_TOC_m.xlsx
  02_CKL.xlsx
  0_Control_Section/
    00_Cover_Sheet_CS.docx
    01_Open_Points.xlsx
    02_Notes_to_Reviewer.xlsx
    03_Time_Sheet.xlsx
  A_Administration/
    00_Cover.docx
    01_North_Star_MRP_Summary.docx
    ...
    07_Social_Media_Outreach_Efforts.docx
  D_Website_Media/
    00_Cover.docx
    01_Index_of_Genres_and_Channels.xlsx
    01_01_Weekly_Digest_Monthly_Summary.xlsx
  Tabs.docx
```

### 3.1 Naming rules

- **Item file:** `NN_<Title>.<ext>` where `NN` is the item's two-digit ordinal within its section (`A.07` → `07_`). Past 99 use three digits for the whole section.
- **Sub-item file:** `NN_MM_<Title>.<ext>` (`D.01.01` → `01_01_`). Deeper nesting appends another `_MM`.
- **Section cover sheet:** `00_Cover.docx` inside the section folder. The Control section's cover keeps its legacy name `00_Cover_Sheet_CS.docx` via override (below).
- **Section folder:** `<Letter>_<Section Label>` (`A_Administration`). Control section: `0_Control_Section`. Sub-sections nest inside their parent section folder. A clarifying header section with no items of its own (e.g. "Content Mediums") gets **no folder, no cover, no tab**; its child sections are placed as if they were top level.
- **Binder root:** `00_Cover.docx`, `01_TOC.xlsx`, `02_CKL.xlsx`, `Tabs.docx`.
- **Title → filename:** replace runs of whitespace with `_`; replace `&` with `and`; drop every character outside `[A-Za-z0-9_-]`; collapse repeated `_`; trim `_` at both ends. `Open Points / Action Items` → `Open_Points_Action_Items`.
- **Filename override:** any node may carry `filename: <exact name>`. Import sets this for every existing file so legacy names (`01_TOC_m.xlsx`, `01_Open_Points.xlsx`, `00_Cover_Sheet_CS.docx`) are preserved exactly. The override holds the title part only; the `NN_` prefix is still computed, so renumbering still renames. The user can clear an override to adopt the computed name.
- **Suffixes:** a node with `suffix: PB` gets code `A.03.PB` and filename `03_PB_<Title>`. Suffixed nodes are not counted in the ordinal sequence; they hang off the node they follow.

### 3.2 Code rules

- Sections are coded in TOC order: the Control section is `0`, then `A`, `B`, `C`, ... Letters never skip. Past `Z` use `AA`, `AB`.
- Clarifying header sections consume no letter.
- Items: `<Section>.<NN>` with two digits, `A.01` ... `A.10` ... `A.99`.
- Sub-items: `<ItemCode>.<MM>`, `D.01.01`. Any depth, same rule.
- Control section items: `0.01`, `0.02`. Always stored and written as **strings**, never numbers (the legacy TOC stores `0.01` as a float; the app must never do that).
- Section cover code: `<Section>.00` (`A.00`, `0.00`). Binder root documents: `TOC`, `CKL`. The binder cover has no code.
- `included: false` on a node: the node keeps its position and code (so other codes do not shift), its file stays on disk, its TOC row loses its `X`, it is skipped from print manifests and merged PDFs. The checklist row stays, flagged.

---

## 4. Manifest schema (`.binder/binder.yaml`)

```yaml
format_version: 1
binder:
  name: North Star Media Management Binder
  header_lines:                      # author-defined, propagated to every document (§5)
    - "{{org_short}}"                # line 1 left; the ref code goes line 1 right automatically
    - "North Star Media"
    - "Management Binder"
    - "{{doc_title}}"
  vars:
    org_short: North Star
  footer:                            # binder default; empty means no footer
    left: ""
    center: ""
    right: ""
  page_numbers: per_document         # per_document | continuous
  tabs:
    product: "Avery 658843"          # product number as given by the user
    template: Tabs.docx              # the legacy table layout is the template
    label_style: letter_and_label    # "C) Campaign Management"; falls back to label_only if too long
  cover:
    template: 00_Cover.docx
    lines: ["North Star Media", "(DCCS Social Media Outreach)", "Management Binder"]
tree:
  - id: control
    kind: section
    label: Control Section
    numeric: true                    # produces 0, 0.01, 0.02
    description: "Used to communicate with your packet reviewer and keep track of your time."
    children:
      - {id: open-points,       kind: item, title: "Open Points / Action Items", ext: xlsx, filename: Open_Points}
      - {id: notes-to-reviewer, kind: item, title: "Notes to Reviewer",           ext: xlsx}
      - {id: time-sheet,        kind: item, title: "Time Sheet",                  ext: xlsx}
  - id: administration
    kind: section
    label: Administration
    children:
      - {id: mrp-summary,      kind: item, title: "North Star MRP Summary",        ext: docx}
      - {id: top-10-list,      kind: item, title: "Top 10 List",                   ext: docx}
      - {id: outreach-efforts, kind: item, title: "Social Media Outreach Efforts", ext: docx}
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
```

Node fields:

| Field | Required | Notes |
| --- | --- | --- |
| `id` | yes | Permanent. Slug of the title at creation; never changes afterwards, even if the title does. Unique within the binder |
| `kind` | yes | `section`, `header`, `item` |
| `label` / `title` | yes | Display text. Sections use `label`, items use `title` |
| `ext` | items | `docx`, `xlsx`, `pdf` |
| `included` | no | default `true` |
| `suffix` | no | e.g. `PB` |
| `filename` | no | Title-part override (§3.1) |
| `locked_header` | no | default `false`. `true` = app never touches header or footer |
| `footer` | no | Per-document footer override `{left, center, right}` |
| `copied_from` | no | `{binder_path, id, synced_at}` when the file was pulled from another binder (§9.3) |
| `numeric` | sections | `true` only for the control section |
| `description` | sections | One line printed under the section title on its cover sheet |
| `children` | no | Nested nodes |

The manifest is read and written with `ruamel.yaml` so comments and ordering survive. `format_version` is bumped on schema changes and the loader upgrades old files in place.

---

## 5. Header and footer model

There is no fixed header template. The author defines `header_lines` once per binder (the app shows the North Star layout as a sample to copy from). Lines may contain `{{org_short}}`, `{{binder_name}}`, `{{doc_title}}`, `{{doc_code}}`, `{{section_label}}`, and any key under `vars`. The reference code is always placed at the right of line 1 as `Packet Ref # <code>`; the author may change the label text (`Ref #:`) in binder settings but it is one setting for the whole binder.

**Page numbers are required on every document that can run more than one page, so on every document.** Word: `Page {PAGE} of {NUMPAGES}` fields in the footer. Excel: `Page &P of &N` in the header's right block under the code, as in the samples.

**Footer:** binder default (`binder.footer`), per-document override (`node.footer`). Blank means no footer. The BMC revision block from the samples is offered as a preset the author can pick, nothing more.

### 5.1 Word implementation

The app owns the primary header and footer of section 1 of the document. It replaces their contents wholesale with paragraphs built from `header_lines`, each line wrapped in a content control (`w:sdt`) tagged `bb:header:<n>` and the code in one tagged `bb:code`. Why content controls: python-docx cannot address header text reliably across formatted runs, but an `sdt` found by tag can be rewritten without touching anything else. Fonts and sizes come from the document's existing header style if one exists, else Normal.

If a document has different first-page or even/odd headers, the app writes the same content to all of them. If `locked_header` is true, nothing in the header or footer is touched.

### 5.2 Excel implementation

openpyxl `ws.oddHeader.center.text` and `.right.text` (and `evenHeader` if set) on every worksheet. Header strings are rebuilt from `header_lines` with `&P of &N` appended on the right. Footer the same way.

### 5.3 PDF items

A PDF item gets a header band overlaid on every page with `pypdf` (a small reportlab-drawn overlay merged onto each page) and `Page n of N` in the footer. The original PDF is kept as `<name>.source.pdf` inside `.binder/pdf_sources/` and regenerated from it on every sync, so overlays never stack. If this proves fragile it is dropped from v1 and PDF items are simply listed and printed as-is with a warning.

---

## 6. Live references

A reference is any place one document names another by code: a sentence in a Word body ("complete the Top 10 List (A.03) first"), the `Reference` column in the Master Checklist, the `From ref` column in Open Points.

### 6.1 Word

A reference is a content control tagged `bb:ref:<target-id>` whose displayed text is the target's current code. The author inserts one from the app ("Insert reference" picks a document from the tree and writes the control at the cursor through Word COM into the open document) or converts typed text during import claim (§9.2). On sync the app finds every `bb:ref:*` control and rewrites its text. The author sees plain "A.03" and can use it anywhere in a sentence.

### 6.2 Excel

Each sync writes a hidden sheet `_binder` with columns `id | code | title | filename` and defines one workbook-level name per row: `bb_<id_with_hyphens_as_underscores>` pointing at that row's code cell. A reference cell contains `=bb_top_10_list`. Excel keeps it current itself; the app only maintains the hidden sheet and names. Codes are written as text (a leading apostrophe is not needed with openpyxl string cells; ensure `number_format = '@'`).

### 6.3 Resolution rules

- A reference whose target ID no longer exists in the tree is a **broken reference**: the control text becomes `[?]`, the Excel name points at a cell containing `[?]`, and the sync report lists each one with the sentence or cell it sits in.
- A reference to a node with `included: false` still resolves to its code (the document exists, it is just not in this physical binder).

---

## 7. Sync (the core operation)

"Sync" replaces the earlier idea of a separate build folder. It makes the binder folder match the manifest.

### 7.1 Plan

Walk the tree, compute codes, folders and filenames, then produce a list of actions:

| Action | When |
| --- | --- |
| `create_file` | Item exists in the tree with no file on disk. Creates a blank document of the chosen `ext` with the correct header, footer, and page numbers. For `docx`: empty body. For `xlsx`: one sheet named after the title, empty. For `pdf`: cannot be created blank; the row is flagged "needs a file" |
| `move_file` | Computed folder or filename differs from where the file is now |
| `stamp_header` | Header, footer, or code text differs from what the file currently carries |
| `update_refs` | Any `bb:ref` control or `_binder` row would change |
| `regenerate` | TOC, a section cover sheet, Tabs sheet, or the binder cover differs from what the manifest would produce |
| `checklist_rows` | Master Checklist is missing rows for new documents or has rows whose code changed |
| `orphan` | A file is inside the binder folder, not in `.binder/`, and not claimed by any node. Reported, never moved |
| `warn` | Tab label too long, broken reference, duplicate ID, locked file in use, PDF item without a file |

### 7.2 Review

The plan is shown as a table (action, file, detail). Nothing happens until the user clicks Apply. Warnings do not block.

### 7.3 Apply

1. Take the structure lock (§10). Abort if someone else holds it.
2. Snapshot every file that will be modified or moved into `.binder/backups/<timestamp>/` preserving relative paths.
3. Apply actions in order: moves, then creates, then stamps and refs, then regenerated documents, then checklist rows.
4. For every modified Word or Excel file run the invariant check (§7.4). On failure, restore that file from the snapshot and report.
5. Update `state.json` with the new per-file content hashes.
6. Append to `sync.log`: user, timestamp, action count, warnings.
7. Release the lock.

Files open in Word or Excel by another user cannot be written (Office holds a lock). The plan detects this before apply and lists them; the user can apply everything else and re-sync later.

### 7.4 Body invariant check

For `.docx`: before and after, compute a hash over the concatenated text of every body paragraph and table cell (`document.element.body`), excluding header and footer parts and excluding the text inside `bb:*` content controls. The two hashes must be equal. For `.xlsx`: hash every cell value on every non-hidden sheet except `_binder`, plus the list of defined names not starting with `bb_`. Equal or the write is rolled back. This is the mechanical guarantee behind safety rule 1 and it is tested on every fixture.

### 7.5 Idempotence

A second sync immediately after the first produces an empty plan. Tested.

---

## 8. Generated documents

All generated documents are rendered from template files that live in the binder folder (so staff can restyle them in Office) with a default set shipped in the app's `templates/` directory. Placeholders use `docxtpl` (Jinja) in Word and `{{ }}` tokens in Excel cells.

| Document | Template | Content |
| --- | --- | --- |
| Binder cover | `00_Cover.docx` | `binder.cover.lines`. No code in header |
| TOC | `01_TOC.xlsx` | Rows in tree order: `X` if included, code, title. Section rows show the letter and label; clarifying headers show the label only. Column D for sub-items as in the samples. In continuous mode an extra `Page` column. Header code `TOC` |
| Section cover | `<section>/00_Cover.docx` | `<Letter>) <Label>`, the optional description line, then `__X__  <code>  <title>` per child in order (`__ __` when not included). Header code `A.00` |
| Tabs | `Tabs.docx` | One cell per section and sub-section in order, skipping clarifying headers. Label is `<Letter>) <Label>`; if the label exceeds the fit limit, fall back to label only; if still too long, warn. Fit limit is a character count derived from the template cell width and font (measured once, stored in binder settings, default 24 characters). More sections than cells → additional copies of the table appended |
| Master Checklist | `02_CKL.xlsx` | Hybrid (§8.1) |

### 8.1 Master Checklist (hybrid)

The CKL is the user's working document with app-maintained columns. Columns `Reference` and `Description` are app-owned: `Reference` is `=bb_<id>` and `Description` is written as text. Every other column (hours, reviewer, dates, final review) is user-owned and never touched. The app also writes a hidden column `bb_id` so rows can be matched after the user re-sorts or inserts rows.

- A document with no CKL row gets one appended in tree order among existing rows (inserted after its nearest preceding sibling's row).
- A document removed from the tree keeps its row; the Reference cell shows `[removed]` and the row is highlighted. The user deletes it if they want.
- The user may add as many extra rows as they like, including rows that reference no document; the app ignores rows without a `bb_id`.
- The TOC, CKL, section covers, and the Tabs sheet all have rows.

The same hybrid mechanism applies to any other Excel document where the user wants a document-per-row list (e.g. Index of Prepared Media): mark the node `hybrid_rows: true` and name the two app-owned columns.

---

## 9. Import and fill

### 9.1 Import an existing binder

1. **Read the TOC.** Parse the TOC workbook into a tree: section rows become sections, `X`/code/title rows become items, nesting from the code pattern. Codes stored as floats (`0.01`) are converted to strings. IDs are slugged from titles.
2. **Match files.** For each node, search the binder folder (recursively) for a file by, in order: header code match, filename ordinal plus title similarity, title similarity alone. Show a pairing table: matched, unmatched nodes, unmatched files. The user confirms or fixes pairs. Confirmed pairs set `filename` overrides so nothing is renamed on first sync.
3. **TOC disagreements.** When the TOC and the files disagree (a file whose header says `A.05` paired with the TOC's `A.06`), the TOC wins and the disagreement is shown. The user can edit the tree before accepting.
4. **Take over headers.** First sync replaces each file's header with the app-owned controls or strings. The previous header text is saved in the import log.
5. **Claim references** (§9.2).

Import changes nothing until the user accepts the pairing table, and the first sync goes through the normal plan/review/apply flow.

### 9.2 Claim references

Scan every Word body and every Excel cell for code-shaped strings: `[A-Z]{1,2}\.\d{2}(\.\d{2})*(\.[A-Z]{1,3})?` and `0\.\d{2}`. For each hit, show the sentence or cell, the code found, and the document it would resolve to. The user approves per hit, or "approve all in this file". Approved hits become `bb:ref` controls or `=bb_<id>` formulas. Unapproved hits stay as plain text and are listed in the report so they can be found later.

### 9.3 Fill a TOC position from any source

From the tree, "Fill from file..." opens a file picker that can browse any folder (another binder, a template library, a desktop). The chosen file is **copied** into the computed position, its header is taken over, and `copied_from` is recorded. The original is never modified. On later syncs, if the original's hash differs from the one recorded, the plan shows a warning "source changed since copy" and offers to re-copy (which replaces this binder's copy wholesale after a backup, so the user must confirm).

---

## 10. Collaboration

v1 uses a lock file. Opening the structure editor writes `.binder/binder.lock` containing user name, machine name, and timestamp. Other users opening the same binder see "Locked by Ben on DESK-07 since 15:30" and get a read-only tree; they can still open and edit documents in Word or Excel as usual (Office's own file locking handles that). Sync requires the lock. A lock older than a configurable age (default 4 hours) can be broken with a confirmation. The storage layer is an interface (`open`, `save`, `lock`, `unlock`, `history`) so a git or server backend can replace it later without touching the engine.

---

## 11. Outputs and printing

All three outputs are available per binder and per selection; the user picks any combination:

1. **Office files in place** (always; this is the sync itself).
2. **PDF per document** into `.binder/pdf/<same relative path>.pdf`, exported through Word and Excel COM so fidelity matches what Office prints.
3. **Merged PDF** of the whole binder in TOC order, with PDF bookmarks for every section and document. The binder cover is first and is not numbered. In `continuous` mode the TOC is page 1 and the running number is stamped bottom center of every page from the TOC onward; the TOC's `Page` column is filled from the real page counts. In `per_document` mode no stamping happens and each document keeps its own `Page X of Y`.

**Print manifest:** `.binder/print_manifest.json`, an ordered list of `{path, pdf_path, code, title, pages, changed_since_last_print}`. Packet Printer gains an "open print manifest" action and prints in order. Packet Printer is not modified by this project beyond that.

**Changed-only printing:** `state.json` records each file's hash at the last print. The print screen offers "all", "changed since last print", and a manual selection. Printing marks the selected files as printed.

---

## 12. GUI

| Screen | Content |
| --- | --- |
| Start | Recent binders, Open folder, New binder, Import existing folder. Version and update notice |
| Binder (main) | Left: tree with drag-and-drop reorder, indent/outdent, Include checkbox, live codes shown next to each node as you drag. Right: node details (title, kind, ext, suffix, filename override, locked header, footer override, copied-from status). Top: lock banner. Bottom: status (orphans, broken refs, warnings count) |
| Add document | Title, type (Word, Excel, PDF file), position (after selected node, as child), optional "fill from file" |
| Binder settings | Name, header lines editor with a live sample preview, ref label, vars, footer editor with presets, page number mode, tabs product and fit limit, cover lines, template folder |
| Sync | Plan table with filters by action type, warnings panel, Apply, Restore from backup |
| Print and export | Checkboxes for PDF per document, merged PDF, print manifest; scope selector (all, changed, selected); Send to Packet Printer |
| Import wizard | Steps: pick folder → pairing table → TOC disagreements → claim references → first sync plan |
| Insert reference | Document picker; writes a `bb:ref` control into the active Word document or `=bb_<id>` into the active Excel cell via COM |

Design intent: simple enough that someone who only authors documents can add a row and sync without help, but with every option reachable from the node details pane.

---

## 13. Tech stack and repo layout

| Layer | Library |
| --- | --- |
| Manifest | `ruamel.yaml`, `pydantic` for validation |
| Word | `python-docx` for reading and body hashing; direct `lxml` for content controls (`w:sdt`) and header/footer parts; `docxtpl` for generated documents |
| Excel | `openpyxl` for everything including defined names, hidden sheets, headers |
| Office automation | `pywin32` COM for PDF export, page counts, and Insert Reference into open documents. Never used for the sync writes themselves |
| PDF | `pypdf` for merge, bookmarks, overlays; `reportlab` for overlay drawing |
| GUI | `tkinter` + `ttk`; `tkinterdnd2` only if drag-and-drop in Treeview needs it |
| Packaging | `PyInstaller`; GitHub Actions workflow builds the release zip on tag |
| Tests | `pytest`; synthetic fixtures |

```
binder-builder/
  README.md
  SPEC.md                     # this file
  pyproject.toml
  binder_builder/
    __init__.py
    engine/                   # no GUI or COM imports in this package
      manifest.py             # load, save, validate, upgrade
      tree.py                 # node model, traversal, id slugging
      codes.py                # tree position -> codes, folders, filenames, tab labels
      word.py                 # sdt helpers, header/footer writer, body hash
      excel.py                # header writer, _binder sheet, defined names, body hash
      pdfdoc.py               # header overlay for PDF items
      refs.py                 # find, resolve, claim
      generate/
        cover.py, toc.py, section_cover.py, tabs.py, checklist.py
      sync.py                 # plan, apply, backup, invariant check, log
      importer.py
      storage/
        base.py, local_folder.py
    office/                   # COM-dependent code, isolated
      export_pdf.py, insert_reference.py, page_count.py
    output/
      merge_pdf.py, print_manifest.py
    gui/
      app.py, start.py, binder_view.py, settings.py, sync_view.py, print_view.py, import_wizard.py
    templates/                # default template set
    updater.py
  tests/
    fixtures/synthetic_binder/
    test_codes.py, test_naming.py, test_word.py, test_excel.py, test_sync.py, test_import.py, test_generate.py
  .github/workflows/release.yml
```

The engine exposes a CLI (`python -m binder_builder sync <folder> --plan`, `--apply`, `import`, `export-pdf`) used by tests and by the GUI alike.

---

## 14. Phases and acceptance tests

Each phase ends with a tagged release and a passing test suite. The phases are ordered so that the office gets value from phase 1 onward.

**Phase 0: Core engine proof.** `manifest`, `tree`, `codes`, `word`, `excel`, `sync` (plan and apply with backup and invariant check), synthetic fixture. CLI only.
Acceptance: load the fixture; sync; assert empty second plan. Insert a new item between A.02 and A.03; sync; assert every file from A.03 onward was renamed, every header code shifted, every `bb:ref` and `bb_` name resolves to the new code, and every body hash is unchanged. Delete a node; assert the file still exists and is reported orphaned.

**Phase 1: Generated documents and checklist.** TOC, section covers, tabs (with fit warning), binder cover, CKL hybrid, template folder.
Acceptance: generated files match golden fixtures; CKL user columns survive three syncs with reorders; tabs warn on a 40-character label.

**Phase 2: GUI.** All screens in §12 except import and insert-reference. Lock file. PyInstaller build and GitHub release workflow. Updater.
Acceptance: a non-technical user creates a binder with five sections and twelve documents, syncs, opens the files in Word, reorders two sections, syncs again, and the files, headers, TOC, covers, and tabs are correct, with zero body changes.

**Phase 3: Import, claim, fill.** Import wizard, claim references, fill from any source with `copied_from` drift warning.
Acceptance: a legacy-style fixture (floats in the TOC, inconsistent header labels, one stale code) imports with every disagreement shown and nothing changed until accepted; after first sync all headers are uniform and claimed references resolve.

**Phase 4: Outputs.** PDF per document and merged PDF via COM, continuous numbering, bookmarks, print manifest, changed-only printing, PDF items with header overlay. Insert Reference via COM.
Acceptance: merged PDF has TOC as page 1, correct bookmarks, correct running numbers; print manifest lists exactly the files changed since the last print.

**Phase 5: Hardening.** Manifest `format_version` upgrade path, "new binder from existing binder's structure", lock breaking, restore-from-backup UI, error reporting.

---

## 15. Open items with chosen defaults

| Item | Default used until Ben says otherwise |
| --- | --- |
| Sub-item filename pattern | `01_01_Title.ext` |
| Section folder name pattern | `A_Administration`; control section `0_Control_Section` |
| Tab fit limit | 24 characters, adjustable in settings |
| Header label text | `Packet Ref #` |
| Lock max age | 4 hours |
| PDF item header overlay | Attempted in phase 4; dropped if fragile |
| Packet Printer location | Separate repo; this project only writes the manifest |

---

## 16. Kickoff prompt for the coding agent

```
Read SPEC.md in full before writing any code. Build phase 0 exactly as specified in
section 14, using the repo layout in section 13. Start by generating the synthetic
fixture binder in tests/fixtures/synthetic_binder/ with the same folder layout, header
layout, TOC layout, and file naming described in sections 3 and 5, using placeholder
text only. Then implement engine/manifest.py, tree.py, codes.py, word.py, excel.py,
and sync.py with the CLI. Every safety rule in section 1.1 is a hard requirement and
section 7.4 must be enforced on every write. Write the phase 0 acceptance tests first
and make them pass. Do not build any GUI in this phase. Commit after each module with a
clear message. When phase 0 passes, stop and summarize what was built, what deviated
from the spec and why, and what you recommend for phase 1.
```
