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
