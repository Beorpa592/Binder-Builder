"""Storage interface (SPEC §10). A git or server backend can replace the local folder later."""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import datetime

from ..manifest import Manifest


class LockedError(Exception):
    def __init__(self, info: "LockInfo"):
        super().__init__(f"Locked by {info.user} on {info.machine} since {info.since:%Y-%m-%d %H:%M}")
        self.info = info


@dataclass
class LockInfo:
    user: str
    machine: str
    since: datetime

    def age_hours(self, now: datetime | None = None) -> float:
        return ((now or datetime.now()) - self.since).total_seconds() / 3600


class Storage(ABC):
    @abstractmethod
    def open(self) -> Manifest: ...

    @abstractmethod
    def save(self, manifest: Manifest) -> None: ...

    @abstractmethod
    def lock(self, force: bool = False) -> bool:
        """Take the structure lock. Returns True if this call acquired it, False if we already held it.

        Raises LockedError if someone else holds it (and `force` is False or the lock is not stale).
        """

    @abstractmethod
    def unlock(self) -> None: ...

    @abstractmethod
    def lock_info(self) -> LockInfo | None: ...

    @abstractmethod
    def history(self) -> list[dict]: ...
