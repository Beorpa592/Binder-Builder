"""Command line: `python -m binder_builder sync <folder> --plan|--apply`, used by tests and the GUI alike."""

from __future__ import annotations

import argparse
import sys

from .engine import sync
from .engine.manifest import ManifestError
from .engine.storage import LockedError


def _cmd_sync(args) -> int:
    p = sync.plan(args.folder)
    print(p.render())
    if not args.apply:
        return 0
    if p.is_empty:
        return 0
    if not args.yes:
        answer = input("\nApply these changes? [y/N] ").strip().lower()
        if answer not in ("y", "yes"):
            print("Nothing changed.")
            return 1
    result = sync.apply(p, force_lock=args.break_lock)
    print(f"\nApplied {len(result.applied)} action(s). Backup: {result.backup or 'none needed'}")
    for f in result.failures:
        print(f"FAILED: {f}", file=sys.stderr)
    return 0 if result.ok else 2


def _cmd_import(args) -> int:
    from .engine import importer

    plan = importer.analyze(args.folder)
    print(f"TOC: {plan.toc_file or '(none found)'}\n")
    print(f"{'CODE':9} {'MATCHED BY':12} {'TITLE':40} FILE")
    for p in plan.pairs.values():
        print(f"{(p.code or ''):9} {(p.method or '-'):12} {p.title[:40]:40} {p.file or '(no file)'}")
    for f in plan.unmatched_files:
        print(f"{'':9} {'unmatched':12} {'':40} {f}")
    dis = plan.disagreements()
    print(f"\n{len(dis)} disagreement(s):")
    for d in dis:
        print(f"  {d.file or d.entity_id}: {d.describe()}")
    for w in plan.warnings:
        print(f"  warning: {w}")
    if not args.accept:
        print("\nNothing changed. Run again with --accept to write .binder/ for this pairing.")
        return 0
    importer.accept(plan)
    print(f"\nWrote {args.folder}/.binder/. Review with 'sync --plan' before the first sync.")
    if args.claim_all:
        hits = importer.claim_scan(args.folder)
        ok = [h for h in hits if h.target_id]
        r = importer.apply_claims(args.folder, ok, [h for h in hits if not h.target_id])
        print(f"Claimed {len(r.claimed)} reference(s); {len(r.skipped)} skipped.")
    return 0


def _cmd_backups(args) -> int:
    for b in sync.list_backups(args.folder):
        print(b)
    return 0


def _cmd_restore(args) -> int:
    restored = sync.restore(args.folder, args.backup)
    print(f"Restored {len(restored)} file(s) from {args.backup}.")
    return 0


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(prog="binder_builder")
    sub = ap.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("sync", help="plan or apply a sync")
    s.add_argument("folder")
    mode = s.add_mutually_exclusive_group()
    mode.add_argument("--plan", action="store_true", help="show the plan only (default)")
    mode.add_argument("--apply", action="store_true", help="show the plan, then apply it")
    s.add_argument("--yes", action="store_true", help="apply without asking")
    s.add_argument("--break-lock", action="store_true", help="break a stale structure lock")
    s.set_defaults(func=_cmd_sync)

    i = sub.add_parser("import", help="pair an existing binder folder with its TOC (changes nothing without --accept)")
    i.add_argument("folder")
    i.add_argument("--accept", action="store_true", help="write .binder/ for the proposed pairing")
    i.add_argument("--claim-all", action="store_true", help="after accepting, claim every resolvable reference")
    i.set_defaults(func=_cmd_import)

    b = sub.add_parser("backups", help="list backups")
    b.add_argument("folder")
    b.set_defaults(func=_cmd_backups)

    r = sub.add_parser("restore", help="restore files from a backup")
    r.add_argument("folder")
    r.add_argument("backup")
    r.set_defaults(func=_cmd_restore)
    return ap


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return args.func(args)
    except (ManifestError, LockedError) as e:
        print(f"error: {e}", file=sys.stderr)
        return 2
