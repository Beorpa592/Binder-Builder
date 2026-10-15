# Deviations from SPEC.md

Every place the implementation departs from SPEC.md, fills a gap the spec leaves
open, or interprets something ambiguous. Each entry gives the smallest reasonable
choice and the reason for it. Ben can overrule any of them.

## Phase 0

### Repo and tooling
1. **Spec file renamed.** The spec arrived as `BINDER_BUILDER_SPEC.md`. It is now `SPEC.md` because §13 and the kickoff prompt both name it that way.
2. **Extra files outside the §13 tree:** `binder_builder/cli.py` and `__main__.py` (needed for `python -m binder_builder`, which §13 requires), `tests/conftest.py` and `tests/helpers.py`, `tests/test_engine_purity.py`, `tests/fixtures/make_fixture.py` and `.gitattributes`. The repo root is the existing `Binder-Builder` checkout, not a nested `binder-builder/` folder.
3. **Tags are local.** Each phase is tagged `phase-N`. Nothing is pushed to GitHub, because publishing to the public repo is Ben's call.

### Manifest schema additions (§4)
4. **`binder.filenames: {cover, toc, ckl, tabs}`** holds title-part overrides for the root documents. §3.1 says import preserves `01_TOC_m.xlsx`, but the TOC is not a tree node, so a node-level `filename` cannot hold that override.
5. **On a section, `filename` overrides the section cover's title part** (`Cover_Sheet_CS` → `00_Cover_Sheet_CS.docx`). This is how §3.1 makes the Control cover keep its legacy name. A new optional `folder` field overrides the section folder's label part.
6. **Other new optional fields.** Binder level: `ref_label` (§5 calls it "one setting for the whole binder"), `template_folder`, `lock_max_age_hours` (§15 default 4) and `tabs.fit_limit` (§15 default 24). Node level: `hybrid_columns` (§8.1 mentions naming the two app-owned columns but gives no field). `copied_from` also gains `source_hash`, so §9.3 can detect drift.
7. **Unknown keys are rejected** (`extra="forbid"`), so typos in hand-edited manifests fail loudly.

### Codes and naming (§3)
8. **Sub-section codes.** §3.2 doesn't say how a section nested in a section is coded. It takes the next letter in the global sequence (`A`, then its sub-section `B`, then the next top-level section `C`). Its folder nests inside the parent's folder (`A_Alpha/B_Sub_Part/`), as §3.1 says.
9. **A clarifying header with items of its own** is reported as a warning and its items get no position. §3.1 defines a header as having no items.
10. **`included: false` cascades.** An excluded section's documents are also treated as excluded for print purposes. They keep their codes, as §3.2 requires.
11. **Title → filename** follows §3.1 literally, so non-ASCII letters are dropped (`Café` → `Caf`).

### Headers (§5)
12. **Word header layout.** Line 1 is `bb:header:1`, a right tab, the plain-text label `Packet Ref # `, then `bb:code`. The code control holds only the code, so references and tests can read it directly. The label is outside the control but still app-owned, because the whole header part is replaced. Lines 2..n are each one paragraph with one `bb:header:n` control. All header controls are content-locked.
13. **Word footer.** Footer text, when not blank, is one line with `bb:footer:left/center/right` controls on center and right tabs. `Page {PAGE} of {NUMPAGES}` goes on its own right-aligned line, so it never collides with the bottom-center running number added in continuous mode (§11).
14. **Excel header layout** follows §5.2: all header lines in the center block, and `Packet Ref # <code>` plus `Page &P of &N` in the right block. The §4 YAML comment's "line 1 left" is taken to describe Word only. User text has `&` escaped as `&&`.
15. **What counts as "header differs".** The comparison uses a canonical rendering of the header and footer parts (app controls, plain text, tabs, and field codes with cached results ignored), not raw XML. Re-saving a document in Word, which turns simple fields into complex ones and adds rsids, does not cause a restamp.
16. **The binder cover and `Tabs.docx` are never stamped.** §8 says the cover has no code in its header, and the tabs sheet is a label sheet. Section covers, the TOC and the CKL are stamped (`A.00`, `TOC`, `CKL`).

### Sync (§7)
17. **Finding a node's file.** The engine first uses the path recorded in `.binder/state.json` from the last sync, then the computed path. It does not embed IDs inside files. Consequence: a file renamed by hand in Explorer is reported as an orphan, and a blank file is planned at the computed path. Nothing is lost, but the plan is noisy. *Recommendation:* in phase 3, add an orphan-to-node "re-link" action, or match orphans by their `bb:code` header.
18. **The invariant check runs before anything is overwritten.** Each Word or Excel write goes to `.binder/tmp/`, is body-hashed, and replaces the original only if the hash matches. On a mismatch the original was never touched. That is stronger than §7.3 step 4's restore-after-failure, and the snapshot still exists either way.
19. **Moves are two-phase** (rename to `*.bbmoveN`, then to the target), so chains (`03→04`, `04→05`) and swaps cannot collide. Each moved file's bytes are hashed before and after the move.
20. **Folders left empty by moves are removed.** Only directories that are completely empty are removed, never files. Otherwise every section-letter change would leave an empty `B_Old_Name/` folder behind. A folder that still holds an orphan stays.
21. **Missing generated documents.** A missing section cover, TOC or CKL is created blank with the right header. Phase 1 renders their content. A missing binder cover or tabs sheet is left for phase 1.
22. **"Empty plan" means no change actions.** Orphans and warnings are reports. The fixture has neither, so its second plan is literally empty.
23. **Restore** (engine and CLI now, UI in phase 5) snapshots the current versions into `<ts>_before-restore/`, moves moved files back, copies the backed-up bytes over them, and restores the pre-sync `state.json`. Files the sync created are left in place (never delete).
24. **CLI `--apply`** prints the plan and asks `y/N` unless `--yes` is given. This keeps "preview before apply" true at the command line too.

### Safety additions (stricter than the spec)
25. **Workbooks openpyxl cannot round-trip are never written.** openpyxl silently drops charts, images, pivot tables, slicers, form controls and macros on save, and the §7.4 value hash would not notice. An `.xlsx` containing any of those parts gets a `warn` ("cannot be saved safely; header and refs not updated") and is left byte-for-byte alone. *Recommendation:* if real binders contain such workbooks, stamp them through Excel COM in a later phase.
26. **Rich text is preserved.** Workbooks are loaded with `rich_text=True`, so formatted runs inside cells survive a save.
27. **Body hashes are stricter than §7.4.** The Excel hash includes hidden user sheets (§7.4 says non-hidden) and sheet-scoped names. The Word hash includes deleted-revision text and field codes, and records where each `bb:*` control sits (but not its text).

## Phase 1

### Templates (§8)
28. **Where templates live.** §8's template column reuses the output names (`00_Cover.docx`, `Tabs.docx`), and a file can't be both its own template and its output. Templates are looked up in the binder's template folder (`.binder/templates/`, or `binder.template_folder`), then in the app's shipped defaults. Names: `binder.cover.template` (default `00_Cover.docx`), `Section_Cover.docx`, `01_TOC.xlsx`, `binder.tabs.template` (default `Tabs.docx`) and `02_CKL.xlsx`. The template folder is never reported as orphans. A binder's legacy `Tabs.docx` layout is used by copying it into the template folder. *Recommendation:* the phase 3 importer should do that copy automatically.
29. **The TOC template mechanism.** Rows whose first cell starts with `{{#section}}`, `{{#header}}`, `{{#item}}` or `{{#subitem}}` are row templates, copied (values and styles) once per TOC row. That's how §8's "`{{ }}` tokens in Excel cells" becomes a TOC that can be restyled. See `binder_builder/templates/README.md`.
30. **The default Tabs template is generic:** an 8-cell, single-column grid. The real Avery 658843 cell geometry isn't known, so the fit limit comes from settings (default 24), not from measuring the template.
31. **Section cover rows** list every item in the section in TOC order (sub-items and suffixed items included) and each direct sub-section. Items inside a sub-section appear on that sub-section's own cover.

### Regeneration (§7.1 `regenerate`)
32. **How "differs from what the manifest would produce" is decided.** The plan renders each generated document to scratch, stamps its header, and compares a content signature (text, header/footer, cell values, names) with the file on disk. A changed template (hash recorded in `state.json`) also triggers regeneration. Generated documents are app-owned, so they're replaced wholesale (after the snapshot) and the body invariant does not apply to them. *Cost:* every plan renders every generated document, which takes about a second for the fixture binder. Fingerprinting the inputs could skip this later if real binders turn out slow.
33. **A missing generated document is rendered**, not created blank. A missing CKL is created from the `02_CKL.xlsx` template.
34. **Continuous mode:** the TOC gets its `Page` heading now. The page values arrive in phase 4, once COM page counts exist.

### Master Checklist (§8.1)
35. **Which documents get rows.** Every document except the binder cover, in TOC order: TOC, CKL, then each section cover and its items, then Tabs (§8.1: "the TOC, CKL, section covers, and the Tabs sheet all have rows"). Generated documents get `_binder` rows and names too (`bb__toc`, `bb__ckl`, `bb_<section>__cover`). Tabs has no code, so its Reference is the plain text `Tabs`.
36. **Flagging.** A removed document's Reference becomes `[removed]`, with a yellow fill on the app-owned cells only, not the whole row, because user cells' formatting belongs to the user. An `included: false` document gets ` [not included]` appended to its Description and a grey fill. Fills are cleared again if the document comes back.
37. **Inserting versus appending.** openpyxl's `insert_rows` does not adjust formulas, merged cells, validations or conditional formats. If the checklist sheet has any of those (or anything else refers to that sheet), new rows are appended after the last row instead of inserted after their sibling, and the plan says so.
38. **The checklist invariant is order-independent.** It is a hash of every user-owned cell, grouped by row and sorted, plus every other sheet's full body hash. Rows may be inserted or re-sorted, but no user value may change.
39. **Legacy adoption.** When a checklist has no `bb_id` column yet, rows whose Reference text equals a current code (float `0.01` accepted) are claimed for that document, and other rows stay user rows. *Assumption:* the manifest's codes match the legacy checklist at adoption time, which holds right after an import.
40. **Hybrid rows on other workbooks:** `hybrid_rows: true` plus `hybrid_columns: [<reference heading>, <description heading>]` (default `Reference`, `Description`). These get one row per item document.

## Phase 2

### Structure
41. **`engine/session.py` is new (not in §13).** It is the GUI-free controller behind every screen: it opens a binder, holds the lock, makes validated edits that save themselves, computes live codes and runs plan/apply/restore. The tkinter modules only present it, so a PySide front end could reuse it unchanged, and the screens are testable without clicking.
42. **Structure edits save the manifest immediately.** There is no "Save" for the tree, so nothing is lost by forgetting. The node details pane has a "Save changes" button for its fields.
43. **Placement rules** (§4 doesn't state them). Documents sit inside a section, or under another document as sub-items. Sections sit at the top level, under a clarifying heading, or inside a section. Headings sit at the top level or inside a section. The GUI explains a refused move, and adding a document "after" a section puts it inside that section.

### Screens (§12)
44. **Drag and drop uses plain Treeview mouse bindings**, so `tkinterdnd2` is not needed. Dropping on the top 30% of a row puts the node before it, the bottom 30% after it, and the middle inside it.
45. **Live codes while dragging show only the dragged node's would-be code** ("Top 10 List becomes B.01"), in the status line. §12 asks for codes "next to each node as you drag". Re-coding every row on each mouse move flickers; the full tree re-codes on drop.
46. **Not yet built, but visible.** The Import button explains that import arrives in phase 3, "Fill from file..." is disabled until phase 3, and the Print & export screen shows its options disabled until phase 4.
47. **The "BMC revision block" footer preset is a placeholder** (`Revision {{revision}}` / `Revised {{revision_date}}`, adding those two variables), because the real block from the samples isn't in the repo.
48. **Opening a folder that isn't a binder** offers to make it an empty binder. Files already in it are listed as orphans and nothing is changed.
49. **The lock is taken when a binder is opened** (the binder screen is the structure editor, §10) and released when it is closed. A lock held by the same user on the same machine is re-entrant. Someone else's lock gives a read-only screen with a banner, plus a "Break lock..." button once the lock is older than `lock_max_age_hours`.
50. **App-level files** live in `%APPDATA%\BinderBuilder\`: `recent.json`, plus `error.log` for error reporting (the user sees a short message and the full traceback goes to the log).
51. **Threading.** Plan and apply run on worker threads. Worker functions never hold Tk objects, and garbage from destroyed screens is collected on the Tk thread. Without this, Tcl crashed when a widget was freed on a worker thread; the GUI tests caught it.

### Packaging and updates (§2)
52. **The updater targets `Beorpa592/Binder-Builder`** (the actual git remote; §2 names the repo `binder-builder`). It self-updates only in the packaged build. It downloads only from `https://github.com/`, rejects archive paths that escape the folder, and copies the new build over the app folder with `robocopy /E`, so old files are never deleted. There is no code signing or signature check yet. *Recommendation:* sign releases before wide rollout.
53. **`BinderBuilder.exe --cli <args>`** runs the command line inside the packaged app (output goes to `%APPDATA%\BinderBuilder\cli.log` when there's no console). It's used to verify the frozen build and is useful for support.
54. **The release workflow** (`.github/workflows/release.yml`) runs on `v*` tags. It fails if the tag doesn't equal `binder_builder.__version__`, runs the tests (the Office COM test skips on CI, which has no Office), builds, and attaches `BinderBuilder-<version>-win64.zip`. It has never run, because nothing has been pushed.

### Acceptance
55. **The phase 2 acceptance test names a non-technical user, which I can't be.** It is covered by three things:
    - `test_phase2_scenario.py`: five sections and twelve documents through the same session the GUI uses, author edits, two sections reordered, re-sync, and every file, header, TOC, cover and tab is correct with zero body changes.
    - `test_gui.py`: drives the real tkinter screens, including genuine mouse drag-and-drop events, include toggles, the details pane, settings, sync apply and read-only lock mode.
    - `test_office_smoke.py`: opens the synced files in real Word and Excel.

    **Still open:** a hands-on session with a real non-technical user.

## Phase 3

### Reading the legacy TOC (§9.1 step 1)
56. **Row parsing.** The include-mark column is whichever column holds the most `X`s. A row is a section if its first non-mark cell is a section code (`0`, `A`, `AA`), an item if it is an item code (floats like `0.01` are converted to `"0.01"`), and a clarifying header if it is a lone text cell after the first section. Title rows above the first section are skipped. An item without an `X` is imported as `included: false`.
57. **Structure the TOC can't express.** A clarifying header groups every section after it until the next header. Sub-sections can't be recovered (TOC letters are flat), so all sections import at the top level or under a header. If the TOC's own numbering has a gap (`A.01, A.03`), the computed codes close it, and each such case is listed as a "numbering" disagreement.
58. **New ids are slugs of the TOC titles** (`open-points-action-items`). A file's type comes from its paired file. An entry with no file defaults to `docx`, and a blank one is created on the first sync.

### Pairing (§9.1 step 2)
59. **Matching thresholds.** A header-code match is accepted only if the titles are also at least 50% similar, so a stale header code can never pair the wrong file. The fixture's Points of Contact header says `A.05`, which is Budget Overview's code. Filename-number matching runs inside the section's folder with the same 50% bar. Title-only matching needs 85%. A file is never paired twice, and manual pairing in the wizard overrides everything.
60. **Header, variables and reference label are inferred** from the most common legacy Word header pattern. In the fixture this yields `{{org_short}}`, `North Star Media`, `Management Binder`, `{{doc_title}}` with label `Packet Ref #`, so the minority `Ref #:` becomes uniform.
61. **What accept writes.** Only `.binder/`: the manifest, `state.json` and `import.json` (pairs, disagreements, every file's previous header text, and unclaimed references). `state.json` records each file's current path, so a file whose number is stale is renamed by the first sync rather than recreated. The legacy `Tabs.docx` is copied into `.binder/templates/` as the tab layout (DEVIATIONS #28).
62. **The tree is edited after accepting, not inside the wizard.** The user makes changes in the binder screen before the first sync. Accepting changes no document, and the first sync is previewed like any other.
63. **Locked headers can't be inferred.** Nothing in a legacy file says "licensed form". The first sync after import stamps every document, and the plan lists each one. *Recommendation:* tick "Locked header" on licensed forms and scans before applying the first sync. The wizard's final message says to review the first sync.

### Claiming references (§9.2)
64. **When claims are applied.** In the wizard, after accept and before the first sync. Each touched file is backed up (`<ts>_claims`) and written through a claim-specific invariant:
    - Word: the full body text, including text inside controls, must be identical.
    - Excel: only the approved cells may change, and each only into a formula.

    A code split by a tab or field can't be wrapped. It is skipped and reported, and so are unapproved hits (in `import.json`).
65. **What isn't scanned or claimed.**
    - Not scanned: the TOC, the CKL and other generated documents (the CKL adopts its rows through §8.1), and PDFs.
    - Not claimed: numeric Excel cells, because `0.01` may be a number rather than a code. A code inside longer cell text becomes a concatenation formula (`="see "&bb_top_10_list`).

### Fill from any source (§9.3)
66. **`copied_from.source_path`** records the exact file copied, and drift is that file's hash against the one recorded. Filling a position that already has a file needs confirmation, and the current file is backed up (`<ts>_fill`). Re-copy replaces this binder's copy wholesale after a backup. If the source is inside another binder, `binder_path` and `id` are filled from that binder's state.
67. **CLI:** `python -m binder_builder import <folder> [--accept] [--claim-all]`.

## Phase 4

### Export (§11, §2)
68. **Word exports with `SaveAs2(FileFormat=wdFormatPDF)`, not `ExportAsFixedFormat`.** On this machine's Word build, `ExportAsFixedFormat` crashes Word on the next `Close`; the test against real Word showed it. Excel uses `ExportAsFixedFormat`.
69. **A per-document watchdog (180 s).** If Word or Excel stops responding (typically a hidden dialog), the exporter ends the automation process it started itself and reports that file as failed. It identifies its own process by the `/Automation -Embedding` command line and the before/after process list, so a user's open Word or Excel is never touched. The batch then carries on with a fresh instance. A disconnected Office process is retried once.
70. **Blank workbooks** (just created, nothing typed yet) would make Excel refuse to print ("nothing to print"). In the read-only copy Excel has open, never saved, the exporter puts a space in A1 of each empty sheet, so the PDF shows one page with the header.
71. **Page counts come from the PDF Word or Excel exported** (`pdf_pages`), because that is exactly what prints. `office/page_count.office_pages` asks Office directly and is kept for quick estimates.
72. **Per-document PDFs are always made**, because they are the merge's inputs. They're cached: a document is re-exported only when its file hash changes.

### Merged PDF
73. **Order and contents.** Binder cover, TOC, CKL, then every section cover and document in TOC order. Excluded (`included: false`) documents and the tab sheet are left out (tab labels print on tab stock). "Selected" scope covers the node, everything inside it and section covers.
74. **Page labels** are set as well as stamped numbers, so a PDF viewer shows "Cover" then 1, 2, 3 from the TOC onward. Clarifying headings get a bookmark too, above their sections.
75. **Continuous mode needs a synced binder.** The merged export refuses with "Sync first" if the plan isn't empty, so page numbers can't describe a TOC that differs from the binder. The running start page of each code is saved in `state.json["pages"]`. A normal sync then regenerates the TOC with its Page column (with backup), the TOC is re-exported, and this repeats (up to 3 rounds) until the TOC's own length settles. Sync always renders the TOC with the saved pages, so the next plan is empty.
76. **Running numbers** are stamped 14 pt from the bottom, centred. Word footers put `Page X of Y` on the right (#13), so the two never overlap.

### Printing
77. **The print manifest** is a top-level JSON array of `{path, pdf_path, code, title, pages, changed_since_last_print}` with absolute paths, in TOC order, covering the chosen scope. "Send to Packet Printer" exports with a manifest and launches Packet Printer with the manifest path, if its location is set (asked once, stored in `%APPDATA%\BinderBuilder\settings.json`). It then asks before marking the documents as printed.
78. **"Changed since last print" compares what prints, not the file's bytes.** It uses:
    - Word: body plus header and footer.
    - Excel: visible sheets, with `=bb_` formulas shown as the codes they display.
    - PDF: the file itself.

    Without this, every workbook would count as changed whenever any code anywhere changed, because its hidden `_binder` sheet does.

### PDF items (§5.3)
79. **The original is kept as `.binder/pdf_sources/<id>.source.pdf`** (by id rather than name, so renames don't matter). A stamp key (original hash plus header text) in `state.json` decides when to restamp. If the file on disk no longer matches what the last sync wrote, someone replaced the PDF, so it becomes the new original and the previous original goes into that sync's backup. The invariant requires the same page count, with each stamped page still containing all of its source page's text. Rotated pages are not specially handled; the overlay is drawn in the page's unrotated frame.

### Insert Reference (§6.1, §12)
80. **Word:** a locked plain-text content control tagged `bb:ref:<id>` is inserted at the cursor of the active document (`GetActiveObject`). The user is told if that document isn't in this binder. **Excel:** `=bb_<id>` goes into the active cell. It shows the code once the workbook has been synced (sync defines the name). No name is added through COM, so all writes to the `_binder` sheet stay with sync.

## Phase 5

81. **Manifest upgrades.** `UPGRADES[n]` holds a step that turns format `n` into `n + 1`, applied one step at a time to the round-trip document (so comments survive). Before the upgraded manifest is written back in place, the original is kept as `.binder/binder.v<n>.yaml.bak`. A manifest from a newer app version is refused with "Update Binder Builder". No real upgrade exists yet because the format is still 1. The path is tested with a simulated v1→v2 step.
82. **New binder from an existing structure.** It copies the tree, keeping node ids so references still resolve, the binder settings and the binder's own templates. It drops legacy filename and folder overrides (the new binder uses computed names) and `copied_from`. "Copy documents" fills each position from the old binder through Fill from file, so the copies are tracked and drift is warned about. Without it, sync creates blank documents with the right headers.
83. **Every backup describes itself** (`index.json` with kind `sync`, `claims`, `fill` or `before-restore`), so all of them can be restored. Before this, claim and fill backups couldn't be. Restore also brings back kept PDF originals. Each sync backup includes the pre-sync `state.json`, including the empty one before a first sync, so restoring a first sync returns the binder exactly to never-synced. Restoring a fill puts the old file back but keeps the node's `copied_from` (the manifest isn't part of a backup).
84. **Error reporting** is "Report a problem..." on the Start and binder screens (and `diagnose` on the CLI). It produces a copyable summary of versions, Office availability, manifest and lock status, recent syncs and the tail of `error.log`, with no document content.
85. **CLI additions:** `unlock [--break]` (refuses unless the lock is older than `lock_max_age_hours`), `new-from`, `export-pdf [--merged] [--scope] [--select]`, and `diagnose`. §13 named `export-pdf`; it is built now.
86. **The version is 0.5.0, not 1.0.0.** Every phase's automated acceptance passes, but no real user has tried it yet (DEVIATIONS #55). 1.0 should follow that.
