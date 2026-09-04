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
    stop_interrupted,
    table,
    tip,
    warn,
)

DEFAULT_TSH_TUNNEL = (
    "tsh ssh -L 9443:192.168.10.3:9443 first_name.last_name@mobilekit-p-phy-device100"
)
DEFAULT_CVT_UI = "https://localhost:9443/cables_validation/"


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
    parser.add_argument(
        "--skip-tunnel-prompt",
        action="store_true",
        help="Skip the Teleport tunnel reminder (CI / already connected).",
    )
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


def _needs_local_tunnel(url: str) -> bool:
    host = (url or "").lower()
    return "localhost" in host or "127.0.0.1" in host


def prompt_tunnel_ready(args: argparse.Namespace, *, total_steps: int) -> None:
    """Show tsh forward command; wait until the user confirms the tunnel is up."""
    if getattr(args, "skip_tunnel_prompt", False):
        return
    if not _needs_local_tunnel(args.url):
        return

    tunnel = os.environ.get("CVT_TSH_TUNNEL", DEFAULT_TSH_TUNNEL).strip() or DEFAULT_TSH_TUNNEL
    ui = os.environ.get("CVT_UI_URL", DEFAULT_CVT_UI).strip() or DEFAULT_CVT_UI

    banner(1, "Start Teleport tunnel", total=total_steps)
    tip("Run this in another terminal (keep it open):")
    blank()
    print(c(Term.BOLD + Term.NEON, f"  {tunnel}"))
    blank()
    tip("Optional — confirm CVT loads in the browser:")
    print(c(Term.NEON, f"  {ui}"))
    blank()
    tip(f"API URL for this run: {args.url}")
    blank()
    answer = ask("Tunnel connected? Continue? [Y/n]").lower()
    if answer in {"n", "no"}:
        raise SystemExit("Stopped — start the tunnel, then re-run this command.")
    ok("Continuing with CVT login")


def _su_number(scope: str) -> str:
    return scope.rsplit("/", 1)[-1]


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
    rows_data = sorted(counts.items(), key=lambda item: _su_sort_key(item[0]))
    total = sum(counts.values())
    banner(None, "CVT Fail + ethernet (DC)")
    tip("Failed records by SU")
    blank()
    table(
        ["SU", "Failed"],
        [[su, f"{count:,}"] for su, count in rows_data]
        + [["Total", f"{total:,}"]],
        widths=[6, 10],
        align_right=False,
    )
    blank()
    ok(f"CSV written → {csv_path}")
    blank()


def cmd_stats(args: argparse.Namespace) -> int:
    total = 3
    prompt_tunnel_ready(args, total_steps=total)
    banner(2, "CVT login", total=total)
    tip(f"URL  {args.url}")
    progress = Progress()
    progress.start_work("logging in to CVT")
    try:
        client = make_client(args)
    finally:
        progress.stop_work()
        sys.stdout.write("\n")
        sys.stdout.flush()
    ok("Logged in")

    banner(3, "DC circuit stats", total=total)
    progress.start_work("fetching circuit stats + SU list")
    try:
        stats = client.circuits_stats(context="dc")
        units = [_su_number(scope) for scope in client.su_scopes()]
    finally:
        progress.finish("stats ready")
    blank()
    for key, value in stats.items():
        metric(str(key), value)
    unique = list(dict.fromkeys(units))
    blank()
    tip(f"SU numbers ({len(unique)})")
    kv("SUs", ", ".join(unique))
    blank()
    print(json.dumps(stats, indent=2))
    return 0


def cmd_circuits(args: argparse.Namespace) -> int:
    total = 4
    prompt_tunnel_ready(args, total_steps=total)
    banner(2, "CVT login", total=total)
    tip(f"URL       {args.url}")
    tip(f"Filter    {args.filter} · status={args.status} · protocol={args.protocol}")
    tip(f"Exclude   A Report contains “{args.a_report_not_contains}”")

    progress = Progress()
    progress.start_work("logging in to CVT")
    try:
        client = make_client(args)
    finally:
        progress.stop_work()
        sys.stdout.write("\n")
        sys.stdout.flush()
    ok("Logged in")

    banner(3, "Walk SU scopes", total=total)
    progress.start_work("loading SU list (data halls)")
    try:
        scopes = client.su_scopes()
    finally:
        progress.stop_work()
        sys.stdout.write("\n")
        sys.stdout.flush()
    tip(f"Walking {len(scopes)} SU numbers (Fail+{args.protocol}, skip CORE)")
    blank()

    seen: set[str] = set()
    circuits: list[dict] = []
    counts: dict[str, int] = {}
    total_scopes = max(len(scopes), 1)

    for index, scope in enumerate(scopes, start=1):
        su = _su_number(scope)
        progress.start_work(
            f"fetching {su}  ({index}/{len(scopes)})  kept {len(circuits)}"
        )
        try:
            chunk = client.circuits(
                context="su",
                items=scope,
                page="circuit",
                healthy=False,
            )
        finally:
            progress.stop_work()

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
        counts[su] = counts.get(su, 0) + added
        progress.bar(index, total_scopes, f"{su}  +{added}  (kept {len(circuits)})")

    progress.finish("SU walk complete")
    blank()
    ok(f"Matched {len(circuits)} circuits across {len(counts)} SUs")

    banner(4, "Write CSV", total=total)
    out_dir = Path(args.out_dir)
    csv_path = out_dir / _timestamped_csv_name(args.csv_name)
    progress.start_work(f"writing {csv_path.name}")
    try:
        write_displayed_csv(csv_path, circuits)
    finally:
        progress.finish("CSV written")
    blank()
    _print_report_table(counts, csv_path)
    tip("Next: ./cvt sharepoint prepare --csv " + str(csv_path))
    blank()
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
    except KeyboardInterrupt:
        stop_interrupted()
    except CvtApiError as exc:
        blank()
        warn(str(exc))
        blank()
        return 2
    except SystemExit:
        raise
    except Exception as exc:  # noqa: BLE001 — show styled error then exit
        blank()
        warn(str(exc))
        blank()
        return 1
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
