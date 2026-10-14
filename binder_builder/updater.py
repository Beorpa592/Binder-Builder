"""Self-update from GitHub Releases (SPEC §2).

At startup the app asks GitHub for the latest release. If it is newer, the Start
screen shows the version and notes. On confirmation (packaged builds only) the
release zip is downloaded and unpacked to a temp folder, a small batch script
waits for this process to exit, copies the new files over the app folder and
restarts the app. Fallback: the user downloads the zip from the Releases page.
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import tempfile
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path

GITHUB_REPO = "Beorpa592/Binder-Builder"
LATEST_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
RELEASES_PAGE = f"https://github.com/{GITHUB_REPO}/releases"
EXE_NAME = "BinderBuilder.exe"
ASSET_PREFIX = "BinderBuilder"


@dataclass
class ReleaseInfo:
    version: str
    notes: str
    page_url: str
    asset_url: str | None


def parse_version(tag: str) -> tuple[int, ...]:
    nums = re.findall(r"\d+", tag.split("-")[0])
    return tuple(int(n) for n in nums[:3]) or (0,)


def is_newer(tag: str, current: str) -> bool:
    return parse_version(tag) > parse_version(current)


def check_for_update(current: str, timeout: float = 5.0, url: str = LATEST_URL) -> ReleaseInfo | None:
    """The latest release if it is newer than `current`; None if not, or if GitHub can't be reached."""
    req = urllib.request.Request(url, headers={"Accept": "application/vnd.github+json",
                                               "User-Agent": "BinderBuilder-updater"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except (OSError, ValueError):
        return None
    tag = data.get("tag_name") or ""
    if not tag or data.get("draft") or data.get("prerelease") or not is_newer(tag, current):
        return None
    asset = next((a["browser_download_url"] for a in data.get("assets", [])
                  if a.get("name", "").startswith(ASSET_PREFIX) and a["name"].endswith(".zip")), None)
    return ReleaseInfo(tag.lstrip("v"), data.get("body") or "", data.get("html_url") or RELEASES_PAGE, asset)


def can_self_update() -> bool:
    return bool(getattr(sys, "frozen", False))


def download_and_stage(info: ReleaseInfo) -> Path:
    """Download and unpack the release zip; returns the folder that holds the new BinderBuilder.exe."""
    if not info.asset_url or not info.asset_url.startswith("https://github.com/"):
        raise RuntimeError(f"no downloadable build for {info.version}; get it from {RELEASES_PAGE}")
    work = Path(tempfile.mkdtemp(prefix=f"BinderBuilder-{info.version}-"))
    zpath = work / "release.zip"
    req = urllib.request.Request(info.asset_url, headers={"User-Agent": "BinderBuilder-updater"})
    with urllib.request.urlopen(req, timeout=120) as resp, open(zpath, "wb") as f:
        while chunk := resp.read(1 << 16):
            f.write(chunk)
    out = work / "unpacked"
    with zipfile.ZipFile(zpath) as z:
        for member in z.namelist():
            if member.startswith("/") or ".." in Path(member).parts:
                raise RuntimeError(f"unsafe path in update archive: {member}")
        z.extractall(out)
    exe = next(out.rglob(EXE_NAME), None)
    if exe is None:
        raise RuntimeError(f"the update archive does not contain {EXE_NAME}")
    return exe.parent


def launch_swap_and_exit(staged: Path) -> None:
    """Start a detached script that waits for this process, copies the new build in, and restarts."""
    app_dir = Path(sys.executable).parent
    pid = os.getpid()
    bat = staged.parent / "apply_update.bat"
    bat.write_text(
        "@echo off\r\n"
        ":wait\r\n"
        f'tasklist /FI "PID eq {pid}" | find "{pid}" >nul && (timeout /t 1 /nobreak >nul & goto wait)\r\n'
        f'robocopy "{staged}" "{app_dir}" /E /R:5 /W:1 >nul\r\n'
        f'start "" "{app_dir / EXE_NAME}"\r\n',
        encoding="ascii",
    )
    flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen(["cmd", "/c", str(bat)], creationflags=flags, close_fds=True)
