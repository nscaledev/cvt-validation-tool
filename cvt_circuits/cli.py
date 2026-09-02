from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path

from cvt_circuits.client import CvtApiError, CvtClient
from cvt_circuits.env import load_dotenv
from cvt_circuits.export import flatten_displayed, write_displayed_csv


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Export CVT Circuits View Issues (Data Center filter) to CSV and Excel.",
    )
    parser.add_argument("--url", default=os.environ.get("CVT_URL", "https://localhost:9443"))
    parser.add_argument("--username", default=os.environ.get("CVT_USERNAME", ""))
    parser.add_argument("--password", default=os.environ.get("CVT_PASSWORD", ""))
    parser.add_argument(
        "--insecure",
        action="store_true",
        default=os.environ.get("CVT_INSECURE", "true").lower() in {"1", "true", "yes"},
    )
    parser.add_argument("--timeout", type=int, default=120)
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("stats", help="Print DC circuit counts")

    circuits = sub.add_parser("circuits", help="Download Fail+ethernet rows with Displayed Columns")
    circuits.add_argument("--filter", choices=["dc"], default="dc")
    circuits.add_argument("--status", default="Fail", help="Circuits View Status filter")
    circuits.add_argument("--protocol", default="ethernet", help="Circuits View Protocol filter")
    circuits.add_argument(
        "--a-report-not-contains",
        default="No Report",
        help="Exclude rows whose A Report contains this text (Circuits View Does Not Contain).",
    )
    circuits.add_argument("--out-dir", default="out")
    circuits.add_argument(
        "--csv",
        dest="csv_name",
        default="circuits-fail-ethernet-dc.csv",
        help="CSV name prefix; a timestamp is inserted before the extension so each run is a new file.",
    )
    return parser


def _timestamped_csv_name(name: str) -> str:
    path = Path(name)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    suffix = path.suffix or ".csv"
    return f"{path.stem}-{stamp}{suffix}"


def require_creds(args: argparse.Namespace) -> None:
    if not args.username or not args.password:
        raise SystemExit(
            "Missing CVT credentials. In this folder run: cp .env.example .env\n"
            "Then set CVT_PASSWORD, or pass --username and --password before the command."
        )


def make_client(args: argparse.Namespace) -> CvtClient:
    require_creds(args)
    client = CvtClient(
        base_url=args.url,
        username=args.username,
        password=args.password,
        insecure=args.insecure,
        timeout=args.timeout,
    )
    client.login()
    return client


def _su_number(scope: str) -> str:
    return scope.rsplit("/", 1)[-1]


def cmd_stats(args: argparse.Namespace) -> int:
    client = make_client(args)
    print(json.dumps(client.circuits_stats(context="dc"), indent=2))
    units = [_su_number(scope) for scope in client.su_scopes()]
    unique = list(dict.fromkeys(units))
    print(f"su_numbers ({len(unique)}): {', '.join(unique)}", file=sys.stderr)
    return 0


def _matches_row(circuit: dict, status: str, protocol: str, a_report_not_contains: str) -> bool:
    circuit_status = str(circuit.get("status") or "")
    circuit_protocol = str(circuit.get("protocol") or "")
    if circuit_status.lower() != status.lower() or circuit_protocol.lower() != protocol.lower():
        return False
    a_report = str(flatten_displayed(circuit).get("A Report") or "")
    needle = (a_report_not_contains or "").lower()
    if needle and needle in a_report.lower():
        return False
    return True


def _su_sort_key(su: str) -> tuple:
    if su.upper().startswith("SU") and su[2:].isdigit():
        return (0, int(su[2:]))
    return (1, su)


def _print_report_table(counts: dict[str, int], csv_path: Path) -> None:
    rows = sorted(counts.items(), key=lambda item: _su_sort_key(item[0]))
    total = sum(counts.values())
    su_w = max(len("SU"), max((len(su) for su, _ in rows), default=2))
    col = "Fail+ethernet remaining"
    n_w = max(len(col), len(str(total)))
    print()
    print(f"{'SU'.ljust(su_w)}  {col.rjust(n_w)}")
    print(f"{'-' * su_w}  {'-' * n_w}")
    for su, count in rows:
        print(f"{su.ljust(su_w)}  {str(count).rjust(n_w)}")
    print(f"{'-' * su_w}  {'-' * n_w}")
    print(f"{'Total'.ljust(su_w)}  {str(total).rjust(n_w)}")
    print()
    print(f"CSV: {csv_path}")


def cmd_circuits(args: argparse.Namespace) -> int:
    client = make_client(args)
    seen: set[str] = set()
    circuits: list[dict] = []
    counts: dict[str, int] = {}
    scopes = client.su_scopes()
    print(
        f"Walking {len(scopes)} SU numbers (Fail+{args.protocol}, skip CORE)",
        file=sys.stderr,
    )
    for scope, chunk in client.iter_su_circuits(healthy=False):
        added = 0
        for circuit in chunk:
            if not _matches_row(circuit, args.status, args.protocol, args.a_report_not_contains):
                continue
            circuit_id = str(circuit.get("circuit_id") or "")
            if circuit_id and circuit_id in seen:
                continue
            if circuit_id:
                seen.add(circuit_id)
            circuits.append(circuit)
            added += 1
        su = _su_number(scope)
        counts[su] = counts.get(su, 0) + added
        print(f"  {su}", file=sys.stderr)

    out_dir = Path(args.out_dir)
    csv_path = out_dir / _timestamped_csv_name(args.csv_name)
    write_displayed_csv(csv_path, circuits)
    _print_report_table(counts, csv_path)
    return 0


def main(argv: list[str] | None = None) -> int:
    load_dotenv(Path(__file__).resolve().parent.parent / ".env")
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "stats":
            return cmd_stats(args)
        if args.command == "circuits":
            return cmd_circuits(args)
        parser.error(f"unknown command {args.command}")
    except CvtApiError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
