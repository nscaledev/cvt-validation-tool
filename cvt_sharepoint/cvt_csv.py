"""Read CVT circuit CSV into dict rows for triage."""

from __future__ import annotations

import csv
from pathlib import Path


REQUIRED_HINTS = (
    "Protocol",
    "Exp A Node",
    "Exp A Port",
    "Exp Z Node",
    "Exp Z Port",
    "A Report",
    "Z Report",
)


def load_cvt_csv(path: Path) -> list[dict[str, str]]:
    with path.open(newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames:
            raise ValueError(f"No header row in {path}")
        rows: list[dict[str, str]] = []
        for raw in reader:
            row = {k: ("" if v is None else str(v).strip()) for k, v in raw.items() if k}
            if not any(row.values()):
                continue
            rows.append(row)
    return rows


def find_latest_cvt_csv(out_dir: Path) -> Path | None:
    candidates = sorted(
        list(out_dir.glob("circuits-*.csv")) + list(out_dir.glob("circuit-report*.csv")),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None
