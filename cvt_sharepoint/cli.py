from __future__ import annotations

import argparse
import sys
import webbrowser
from pathlib import Path

from cvt_circuits.env import load_dotenv, upsert_dotenv
from cvt_sharepoint.discover import SHAREPOINT_OPEN_URL, discover_trackers
from cvt_sharepoint.local_workbook import (
    LocalWorkbookError,
    require_local_workbook,
    sheet_names,
    update_range as local_update_range,
    workbook_info,
)
from cvt_sharepoint.prepare import run_prepare
from cvt_term import stop_interrupted

ENV_FILE = Path(__file__).resolve().parent.parent / ".env"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Prepare a local CABLING_REMEDIATION working copy from a SharePoint "
            "download + CVT CSV, then open an HTML publish handoff."
        )
    )
    sub = parser.add_subparsers(dest="command", required=True)

    prepare = sub.add_parser(
        "prepare",
        help="Build capped handoff (20 reopen + 300 adds) as HTML/Excel for SharePoint paste",
    )
    prepare.add_argument(
        "--csv",
        dest="cvt_csv",
        help="CVT circuits CSV (default: newest circuits-*.csv in out/)",
    )
    prepare.add_argument(
        "--tracker",
        help="Existing tracker xlsx (skips download wait if provided)",
    )
    prepare.add_argument(
        "--dry-run",
        action="store_true",
        help="Compute reopen/add plan and HTML without modifying the working xlsx",
    )
    prepare.add_argument(
        "--no-browser",
        action="store_true",
        help="Do not open SharePoint or the HTML report automatically",
    )

    sub.add_parser("open", help="Open the SharePoint workbook in your browser")

    use = sub.add_parser("use", help="Point helper commands at a local xlsx")
    use.add_argument("--file", help="Path to xlsx")

    sub.add_parser("status", help="Show configured local workbook")
    sub.add_parser("worksheets", help="List worksheet names")

    update = sub.add_parser("update", help="Write one A1-style range (debug helper)")
    update.add_argument("--sheet", required=True)
    update.add_argument("--cell", required=True)
    update.add_argument("--value", required=True, action="append")
    update.add_argument("--dry-run", action="store_true")
    return parser


def _print_local_info(path: Path) -> None:
    info = workbook_info(path)
    print(f"Workbook: {info['name']}")
    print(f"  path: {info['path']}")
    print(f"  lastModified: {info['lastModified']}")
    print(f"  size: {info['size']} bytes")


def cmd_prepare(args: argparse.Namespace) -> int:
    from cvt_sharepoint.prepare import print_handoff_summary

    tracker, report, xlsx = run_prepare(
        cvt_csv=Path(args.cvt_csv).expanduser() if args.cvt_csv else None,
        tracker=Path(args.tracker).expanduser() if args.tracker else None,
        dry_run=args.dry_run,
        open_browser=not args.no_browser,
        skip_prompt=bool(args.tracker),
    )
    print_handoff_summary(tracker, report, xlsx)
    return 0


def cmd_open(args: argparse.Namespace) -> int:
    del args
    print("Opening SharePoint workbook…")
    webbrowser.open(SHAREPOINT_OPEN_URL)
    return 0


def _pick_download(explicit: str | None) -> Path:
    if explicit:
        path = Path(explicit).expanduser()
        if not path.is_file():
            raise LocalWorkbookError(f"File not found: {path}")
        return path.resolve()
    matches = discover_trackers()
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        listed = "\n".join(f"  {item}" for item in matches)
        raise LocalWorkbookError(
            "Several tracker downloads found. Pick one:\n"
            f"{listed}\n"
            "Then: ./cvt sharepoint use --file /full/path.xlsx"
        )
    raise LocalWorkbookError("No tracker xlsx in Downloads/Desktop.")


def cmd_use(args: argparse.Namespace) -> int:
    path = _pick_download(args.file)
    upsert_dotenv(ENV_FILE, "SHAREPOINT_LOCAL_PATH", str(path))
    _print_local_info(path)
    print(f"Saved SHAREPOINT_LOCAL_PATH in {ENV_FILE}")
    return 0


def cmd_status(args: argparse.Namespace) -> int:
    del args
    _print_local_info(require_local_workbook())
    return 0


def cmd_worksheets(args: argparse.Namespace) -> int:
    del args
    path = require_local_workbook()
    _print_local_info(path)
    print("Worksheets:")
    for name in sheet_names(path):
        print(f"  {name}")
    return 0


def cmd_update(args: argparse.Namespace) -> int:
    path = require_local_workbook()
    values = [list(args.value)]
    _print_local_info(path)
    print(f"sheet={args.sheet} cell={args.cell} values={values}")
    if args.dry_run:
        print("Dry run: no cells were written.")
        return 0
    address = local_update_range(path, args.sheet, args.cell, values)
    print(f"Updated {address} on disk.")
    return 0


def main(argv: list[str] | None = None) -> int:
    load_dotenv(ENV_FILE)
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "prepare":
            return cmd_prepare(args)
        if args.command == "open":
            return cmd_open(args)
        if args.command == "use":
            return cmd_use(args)
        if args.command == "status":
            return cmd_status(args)
        if args.command == "worksheets":
            return cmd_worksheets(args)
        if args.command == "update":
            return cmd_update(args)
        parser.error(f"unknown command {args.command}")
    except KeyboardInterrupt:
        stop_interrupted()
    except (LocalWorkbookError, FileNotFoundError, ValueError, RuntimeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
