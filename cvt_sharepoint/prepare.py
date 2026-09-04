"""CLI prepare flow: download prompt → triage → capped handoff HTML/Excel."""

from __future__ import annotations

import os
import sys
import time
import webbrowser
from datetime import datetime
from pathlib import Path

from cvt_sharepoint.cvt_csv import find_latest_cvt_csv, load_cvt_csv
from cvt_sharepoint.discover import SHAREPOINT_OPEN_URL
from cvt_sharepoint.handoff import ADD_LIMIT, REOPEN_LIMIT, build_handoff, write_handoff_xlsx
from cvt_sharepoint.report_html import write_handoff_html
from cvt_term import (
    Progress,
    Term,
    ask,
    banner,
    blank,
    c,
    kv,
    metric,
    ok,
    rule,
    tip,
    wait,
    warn,
)

TRACKER_PREFIX = "Nscale_WC_Cabling_HW_Remediation_Tracker_16k"
OUT_DIRNAME = "out"


def print_handoff_summary(tracker: Path, report: Path, xlsx: Path) -> None:
    """Single final summary after prepare (paths + paste tip)."""
    banner(None, "Done — handoff ready")
    tip("Local tracker was NOT modified (read-only for keys)")
    blank()
    kv("Tracker", str(tracker))
    kv("HTML", str(report), color=Term.NEON)
    kv("Excel", str(xlsx), color=Term.NEON)
    blank()
    tip("Open the HTML → copy TSV → paste Values into SharePoint CABLING_REMEDIATION")
    blank()


def project_out_dir() -> Path:
    return Path(__file__).resolve().parent.parent / OUT_DIRNAME


def _is_tracker_candidate(path: Path) -> bool:
    name = path.name
    if not name.startswith(TRACKER_PREFIX):
        return False
    if name.startswith("~$") or "_backup_" in name:
        return False
    return path.suffix.lower() in {".xlsx", ".xlsm"}


def find_trackers(out_dir: Path) -> list[Path]:
    if not out_dir.is_dir():
        return []
    found = [p for p in out_dir.iterdir() if p.is_file() and _is_tracker_candidate(p)]
    found.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return found


def prompt_download(out_dir: Path, *, open_browser: bool = True) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    banner(1, "Download tracker into out/")
    tip(f"Save here: {out_dir}")
    tip(f"Name must start with: {TRACKER_PREFIX}")
    tip("SharePoint → File → Create a copy → Download a copy (.xlsx)")
    if open_browser:
        ok("Opening SharePoint in your browser")
        webbrowser.open(SHAREPOINT_OPEN_URL)

    existing = find_trackers(out_dir)
    if existing:
        blank()
        tip("Already in out/:")
        for item in existing[:5]:
            kv("•", item.name)
        answer = ask("Use newest tracker in out/? [Y/n]").lower()
        if answer in {"", "y", "yes"}:
            return existing[0]

    blank()
    wait("Waiting for download (poll every 3s, Ctrl+C cancel)")
    known = {p.resolve() for p in find_trackers(out_dir)}
    while True:
        time.sleep(3)
        current = find_trackers(out_dir)
        for path in current:
            if path.resolve() not in known:
                ok(f"Detected: {path.name}")
                confirm = ask("Use this file? [Y/n]").lower()
                if confirm in {"", "y", "yes"}:
                    return path
                known.add(path.resolve())
        if current:
            newest = current[0]
            kv("Newest", newest.name)
            again = ask("Type y to use newest, or Enter to keep waiting").lower()
            if again in {"y", "yes"}:
                return newest


def run_prepare(
    *,
    out_dir: Path | None = None,
    cvt_csv: Path | None = None,
    tracker: Path | None = None,
    dry_run: bool = False,
    open_browser: bool = True,
    skip_prompt: bool = False,
    reopen_limit: int = REOPEN_LIMIT,
    add_limit: int = ADD_LIMIT,
) -> tuple[Path, Path, Path]:
    del dry_run  # handoff never mutates tracker remediation body
    out_dir = out_dir or project_out_dir()
    out_dir.mkdir(parents=True, exist_ok=True)

    if tracker is None:
        if skip_prompt:
            found = find_trackers(out_dir)
            if not found:
                raise FileNotFoundError(f"No tracker matching {TRACKER_PREFIX}*.xlsx in {out_dir}")
            tracker = found[0]
        else:
            tracker = prompt_download(out_dir, open_browser=open_browser)
    elif not tracker.is_file():
        raise FileNotFoundError(tracker)

    if not _is_tracker_candidate(tracker):
        raise ValueError(f"Refusing {tracker.name}. Expected prefix {TRACKER_PREFIX}")

    banner(2, "Triage")
    kv("Tracker", tracker.name)
    tip("Tracker is read-only — no local backup or edits")

    if cvt_csv is not None:
        cvt_csv = Path(cvt_csv).expanduser()
        # People sometimes paste the docs placeholder literally.
        if (not cvt_csv.is_file()) or "..." in cvt_csv.name:
            warn(f"CSV not found (or placeholder): {cvt_csv}")
            cvt_csv = None

    if cvt_csv is None:
        cvt_csv = find_latest_cvt_csv(out_dir)
    if cvt_csv is None or not cvt_csv.is_file():
        available = sorted(out_dir.glob("circuits*.csv")) + sorted(out_dir.glob("circuit-report*.csv"))
        listed = "\n".join(f"    {p.name}" for p in available) or "    (none)"
        raise FileNotFoundError(
            "No CVT circuits CSV found.\n"
            "Pass a real file, for example:\n"
            "  --csv out/circuits-fail-ethernet-dc-20260902-140237.csv\n"
            "Or omit --csv to auto-pick the newest circuits*.csv in out/.\n"
            f"CSVs currently in out/:\n{listed}\n"
            "Generate one with: ./cvt circuits circuits --out-dir out"
        )

    kv("CSV", cvt_csv.name)
    cvt_rows = load_cvt_csv(cvt_csv)
    ok(f"Loaded {len(cvt_rows)} CVT rows")
    blank()
    wait("Triage started — tracker is streamed once per run (not kept in memory afterward)")

    progress = Progress()
    phase_weights = {
        "open": (0, 35),
        "nodes": (35, 40),
        "index": (40, 75),
        "hints": (75, 82),
        "match": (82, 97),
        "batch": (97, 100),
    }
    labels = {
        "open": "stage 1/6  opening tracker (streamed read)",
        "nodes": "stage 2/6  loading Node_Table / ADMIN",
        "index": "stage 3/6  indexing remediation keys",
        "hints": "stage 4/6  NVIDIA comment hints ready",
        "match": "stage 5/6  matching CVT rows",
        "batch": "stage 6/6  building handoff batch",
    }

    def on_progress(phase: str, current: int, total: int | None) -> None:
        label = labels.get(phase, phase)
        lo, hi = phase_weights.get(phase, (0, 100))
        if phase == "open":
            progress.start_work(label)
            return
        progress.stop_work()
        if total and total > 0 and phase in {"index", "match"}:
            frac = current / total
            overall = int(lo + (hi - lo) * frac)
            progress.bar(overall, 100, f"{label} ({current}/{total})")
        else:
            progress.bar(hi, 100, label)

    try:
        result = build_handoff(
            tracker,
            cvt_rows,
            backup_path=None,
            cvt_csv=cvt_csv,
            reopen_limit=reopen_limit,
            add_limit=add_limit,
            on_progress=on_progress,
        )
    finally:
        progress.finish("Triage complete")
    ok("Triage finished")

    banner(3, "This audit batch")
    blank()
    rule()
    metric("Reopen in this batch", f"{len(result.reopens)}  (cap {reopen_limit})", alert=True)
    metric("Adds in this batch", f"{len(result.adds)}  (cap {add_limit})")
    remain_reopen = max(0, result.reopen_total - len(result.reopens))
    remain_adds = max(0, result.add_total - len(result.adds))
    metric("Remaining after run", f"reopen {remain_reopen}  ·  adds {remain_adds}")
    rule()

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    report = out_dir / f"handoff-{stamp}.html"
    xlsx = out_dir / f"handoff-{stamp}.xlsx"
    write_handoff_xlsx(xlsx, result)
    write_handoff_html(report, result, xlsx_name=xlsx.name)

    if open_browser:
        webbrowser.open(report.resolve().as_uri())
    return tracker, report, xlsx
