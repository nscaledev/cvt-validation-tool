from __future__ import annotations

import os
from pathlib import Path


def load_dotenv(path: Path | None = None) -> None:
    env_path = path or Path(".env")
    if not env_path.is_file():
        return
    for raw in env_path.read_text().splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip("'").strip('"')
        os.environ.setdefault(key, value)


def upsert_dotenv(path: Path, key: str, value: str) -> None:
    """Create or replace KEY=value in a .env file. Values with spaces are quoted."""
    rendered = value
    if any(ch in value for ch in (' ', '#', '"', "'")):
        rendered = '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'
    line = f"{key}={rendered}"
    rows: list[str] = []
    replaced = False
    if path.is_file():
        for raw in path.read_text().splitlines():
            stripped = raw.strip()
            if stripped.startswith(f"{key}=") or stripped.startswith(f"{key} ="):
                rows.append(line)
                replaced = True
            else:
                rows.append(raw)
    if not replaced:
        if rows and rows[-1] != "":
            rows.append(line)
        else:
            rows.append(line)
    path.write_text("\n".join(rows).rstrip() + "\n")
    os.environ[key] = value
