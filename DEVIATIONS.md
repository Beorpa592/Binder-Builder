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
