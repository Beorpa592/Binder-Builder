# Phase log

## Phase 0: Core engine proof (tag `phase-0`)

**Built:** `engine/manifest.py` (round-trip YAML, pydantic validation, version gate),
`tree.py` (node edits that keep YAML comments), `codes.py` (codes, folders, filenames,
tab labels, header text), `word.py` (`bb:*` content controls, header/footer writer,
`bb:ref` updates, body hash), `excel.py` (header writer, hidden `_binder` sheet, `bb_`
names, body hash), `storage/` (interface plus local-folder backend with lock file),
`sync.py` (plan, apply, backups, invariant check, `state.json`, `sync.log`, restore),
and the CLI. There is also a synthetic fixture binder with its generator.

**Acceptance (all passing):**
- Load the fixture, sync, and the second plan is empty.
- Insert an item between A.02 and A.03, then sync. A.03 onward (including `A.03.PB`)
  is renamed, every header code shifts, every `bb:ref` and every `bb_` name in every
  workbook resolves to the new code, and every body hash is unchanged.
- Delete a node. Its file still exists and is reported as an orphan.
- Plus: locked headers stay untouched, swapping sections loses no files, backups
  hold the original bytes, a tampered write is rejected by the invariant check,
  restore works, someone else's lock blocks apply, the CLI works, and the engine
  contains no GUI or COM imports.
- Manually verified in real Word and Excel through COM. Stamped headers render,
  `Page X of Y` computes, and `=bb_top_10_list` evaluates to the code.

**Deviations:** DEVIATIONS.md items 1–27. The ones with real consequences:
- Workbooks containing charts, images, pivots or macros are never written (#25).
- Files renamed by hand in Explorer show up as orphans (#17).

## Phase 1: Generated documents and checklist (tag `phase-1`)

**Built:** `engine/generate/` (binder cover, section covers, TOC, tabs, hybrid checklist),
a default template set with its builder (`binder_builder/templates/`), template-folder
override with change detection, `regenerate` and `checklist_rows` actions in sync, and
text-dump goldens (`tests/fixtures/golden/`, rebuilt by `make_golden.py`).

**Acceptance (all passing, 60 tests total):**
- Generated files match the golden fixtures: binder cover, TOC, CKL, Tabs and all six section covers.
- CKL user columns survive three syncs with reorders (section swap, cross-section
  move, insertion), plus a user re-sort and a free user row. Each document keeps exactly
  one row, the new one lands after its sibling, and the next plan is empty.
- Tabs warn on a 40-character label, and overflow appends a second table copy.
- Plus: removed documents are flagged `[removed]`, a CKL with user formulas gets rows
  appended rather than inserted, a missing CKL is rebuilt from the template, the
  template folder overrides defaults and triggers regeneration, and backups include
  regenerated docs.
- Manually verified in real Office through COM. CKL names evaluate (`TOC`, `CKL`, `0.00`,
  `0.01`), the `bb_id` column is hidden, TOC codes are text, and generated docs open.

**Deviations:** DEVIATIONS.md items 28–40. The ones with real consequences:
- Templates live in `.binder/templates/`, not as the output files (#28).
- A checklist with user formulas gets new rows at the end rather than after their sibling (#37).

## Phase 2: GUI (tag `phase-2`, version 0.2.0)

**Built:** `engine/session.py` (GUI-free controller), the tkinter screens (`gui/app.py`,
`start.py`, `binder_view.py`, `settings.py`, `sync_view.py`, `print_view.py`, with
`import_wizard.py` as a phase 3 placeholder), the lock-file workflow, `updater.py`, the
PyInstaller build (`packaging/build.py`, which makes `dist/BinderBuilder/BinderBuilder.exe`
and a release zip), and the GitHub release workflow.

**Acceptance (all passing, 78 tests total):**
- Scenario through the session: five sections and twelve documents, sync, author edits,
  two sections reordered, sync again. Files, headers, TOC, covers and tabs are correct, and
  author body hashes are unchanged.
- GUI smoke tests on real widgets: open/close (the lock is released), toolbar moves,
  include toggle, details rename, a real drag-and-drop with the live code shown, the sync
  dialog plans and applies, settings save, and read-only mode under someone else's lock.
- Synced files open in real Word and Excel. Headers show, `Page 1 of 1` computes, and
  `=bb_` names evaluate.
- The packaged exe was built and launched: the window opened as "Binder Builder 0.2.0"
  with no errors, `--cli sync --apply` ran a full sync from the frozen build, and the
  next plan was empty.
- Not done: a hands-on test with a non-technical user (DEVIATIONS #55), and a real
  GitHub release run (nothing has been pushed).

**Deviations:** DEVIATIONS.md items 41–55.

## Phase 3: Import, claim, fill (tag `phase-3`, version 0.3.0)

**Built:** `engine/importer.py` (TOC parse, pairing, disagreements, accept, claims),
`engine/refs.py` (code scan, claims to `bb:ref` controls and `=bb_` formulas, claim invariants),
`BinderSession.fill_from_file` / `recopy` with the drift warning in the plan, the import
wizard GUI, and Fill from file / Re-copy in the binder screen and the Add document dialog.
There is also the CLI `import` command and a legacy fixture (`tests/fixtures/legacy_binder/`).

**Acceptance (all passing, 90 tests total):**
- The legacy fixture has float codes in the TOC, an inconsistent `Ref #:` label, one stale
  header code (A.05 on A.04) and one stale filename number (05 for A.06). Every file pairs
  (the stale code doesn't mis-pair), exactly those two disagreements are shown, analysis
  changes no byte on disk, and accepting writes only `.binder/`.
- After claims and the first sync, every header carries the uniform `Packet Ref #` label
  and the TOC's code (the stale code is fixed, the stale file is renamed). All 9 claimed
  references resolve and follow a later insertion, author text is identical, and the second
  plan is empty.
- Fill from file copies without touching the source, the header is taken over, a changed
  source is warned about, and re-copy clears the warning.
- The import wizard was driven end to end in the GUI test. A claimed document opens in real
  Word with both `bb:ref` controls and its sentence intact.

**Deviations:** DEVIATIONS.md items 56–67. The one with real consequences:
- Locked headers can't be inferred, so lock licensed forms before the first sync (#63).
