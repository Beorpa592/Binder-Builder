"""Binder stored as a plain folder on a (shared) drive, state in `.binder/`."""

from __future__ import annotations

import getpass
import json
import os
import socket
from datetime import datetime
from pathlib import Path

from ..manifest import BINDER_DIR, DEFAULT_LOCK_MAX_AGE_HOURS, Manifest
from .base import LockedError, LockInfo, Storage

LOCK_NAME = "binder.lock"
LOG_NAME = "sync.log"


def current_user() -> tuple[str, str]:
    return getpass.getuser(), socket.gethostname()


class LocalFolderStorage(Storage):
    def __init__(self, folder: Path | str):
        self.folder = Path(folder)
        self.binder_dir = self.folder / BINDER_DIR

    @property
    def lock_path(self) -> Path:
        return self.binder_dir / LOCK_NAME

    @property
    def log_path(self) -> Path:
        return self.binder_dir / LOG_NAME

    def open(self) -> Manifest:
        return Manifest.load(self.folder)

    def save(self, manifest: Manifest) -> None:
        manifest.save(self.binder_dir / "binder.yaml")

    def lock_info(self) -> LockInfo | None:
        try:
            data = json.loads(self.lock_path.read_text(encoding="utf-8"))
            return LockInfo(data["user"], data["machine"], datetime.fromisoformat(data["since"]))
        except FileNotFoundError:
            return None
        except (ValueError, KeyError):
            # unreadable lock: treat as held by unknown, dated by file mtime
            since = datetime.fromtimestamp(self.lock_path.stat().st_mtime)
            return LockInfo("unknown", "unknown", since)

    def held_by_me(self) -> bool:
        info = self.lock_info()
        return info is not None and (info.user, info.machine) == current_user()

    def lock(self, force: bool = False, max_age_hours: float = DEFAULT_LOCK_MAX_AGE_HOURS) -> bool:
        info = self.lock_info()
        if info is not None:
            if (info.user, info.machine) == current_user():
                return False
            if not (force and info.age_hours() >= max_age_hours):
                raise LockedError(info)
            # breaking a stale lock: the lock file is the app's own state, not user content
            self.lock_path.unlink()
        user, machine = current_user()
        self.binder_dir.mkdir(parents=True, exist_ok=True)
        payload = json.dumps({"user": user, "machine": machine, "since": datetime.now().isoformat(timespec="seconds")})
        try:
            fd = os.open(self.lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            raise LockedError(self.lock_info())  # lost a race
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(payload)
        return True

    def unlock(self) -> None:
        if self.held_by_me():
            self.lock_path.unlink()

    def append_log(self, entry: dict) -> None:
        self.binder_dir.mkdir(parents=True, exist_ok=True)
        with self.log_path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(entry) + "\n")

    def history(self) -> list[dict]:
        if not self.log_path.exists():
            return []
        return [json.loads(line) for line in self.log_path.read_text(encoding="utf-8").splitlines() if line.strip()]
