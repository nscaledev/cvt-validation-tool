from __future__ import annotations

import os
import subprocess
from pathlib import Path

TRACKER_NEEDLE = "nscale_wc_cabling_hw_remediation_tracker"

SHAREPOINT_OPEN_URL = (
    "https://dell.sharepoint.com/:x:/r/sites/NSCALE-WARDCTYTX-Phase1/"
    "_layouts/15/Doc.aspx?sourcedoc=%7B4FB77D68-CC31-491D-B23E-6F7DFC720561%7D"
    "&file=Nscale_WC_Cabling_HW_Remediation_Tracker_16k%20-%20New.xlsx.xlsx"
    "&action=default"
)


def _is_tracker(path: Path) -> bool:
    name = path.name.lower()
    return TRACKER_NEEDLE in name and name.endswith(".xlsx") and not name.startswith("~$")


def _mdfind(root: Path) -> list[Path]:
    if not root.is_dir():
        return []
    proc = subprocess.run(
        [
            "mdfind",
            "-onlyin",
            str(root),
            f"kMDItemFSName == '*Nscale_WC_Cabling_HW_Remediation_Tracker*'c",
        ],
        capture_output=True,
        text=True,
    )
    if proc.returncode != 0:
        return []
    found: list[Path] = []
    for line in proc.stdout.splitlines():
        path = Path(line.strip())
        if path.is_file() and _is_tracker(path):
            found.append(path)
    return found


def _walk(root: Path, max_depth: int = 8) -> list[Path]:
    found: list[Path] = []
    if not root.is_dir():
        return found
    root = root.resolve()
    for dirpath, dirnames, filenames in os.walk(root):
        rel = Path(dirpath).resolve().relative_to(root)
        depth = len(rel.parts) if rel.parts != (".",) and str(rel) != "." else 0
        if depth > max_depth:
            dirnames[:] = []
            continue
        dirnames[:] = [name for name in dirnames if not name.startswith(".")]
        for name in filenames:
            path = Path(dirpath) / name
            if _is_tracker(path):
                found.append(path)
    return found


def search_roots() -> list[Path]:
    home = Path.home()
    roots = [home / "Downloads", home / "Desktop"]
    unique: list[Path] = []
    seen: set[Path] = set()
    for root in roots:
        try:
            resolved = root.resolve()
        except OSError:
            continue
        if resolved in seen or not root.is_dir():
            continue
        seen.add(resolved)
        unique.append(root)
    return unique


def discover_trackers() -> list[Path]:
    found: list[Path] = []
    seen: set[Path] = set()
    for root in search_roots():
        chunk = _mdfind(root) or _walk(root)
        for path in chunk:
            try:
                resolved = path.resolve()
            except OSError:
                continue
            if resolved in seen:
                continue
            seen.add(resolved)
            found.append(resolved)
    found.sort(key=lambda item: item.stat().st_mtime, reverse=True)
    return found
