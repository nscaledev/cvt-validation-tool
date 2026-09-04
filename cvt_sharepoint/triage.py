"""Port CVT_IMPORT formula logic used to build CABLING_REMEDIATION rows."""

from __future__ import annotations

import getpass
import os
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Any

# Same delimiter used by the tracker workbook formulas.
KEY_SEP = "¦"

RESOLVED_STATUS = "confirmed resolved (nvidia)"


def _s(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _clean_upper(value: Any) -> str:
    # CLEAN() removes non-printables; approximate with stripping control chars.
    text = _s(value)
    text = "".join(ch for ch in text if ord(ch) >= 32)
    return text.upper()


def normalize_protocol(protocol: Any) -> str:
    p = _s(protocol).lower()
    if p == "ib":
        return "IB"
    if p == "ethernet":
        return "ETH"
    if p == "nvlink":
        return "NVLINK"
    return _s(protocol).upper()


def pick_su(a_su: Any, z_su: Any) -> str:
    a = _s(a_su)
    z = _s(z_su)
    if a and a.upper() != "NA":
        return a
    if z and z.upper() != "NA":
        return z
    return ""


def normalize_port(host: Any, port: Any, *, admin_map: dict[str, str] | None = None) -> str:
    """Port CVT_IMPORT AP/AV style normalization."""
    g = _s(host)
    h = _s(port)
    admin_map = admin_map or {}
    if not g:
        return ""
    if h == "enP22s2f0np0":
        return "M1"
    if h == "enP22s2f1np1":
        return "M2"
    # Ethernet swp...
    if h.lower().startswith("swp"):
        rest = h[3:]
        host_tail = g[-2:].lower()
        delim = "/1/" if host_tail in {"m1", "m2"} else "/"
        m = re.match(r"^(\d+)s(\d+)$", rest, re.I)
        if m:
            return f"{m.group(1)}{delim}{int(m.group(2)) + 1}"
        return rest
    # IB / other sw...
    if len(h) >= 2 and h[:2].lower() == "sw" and not h.lower().startswith("swp"):
        rest = h[2:]
        m = re.match(r"^(\d+)s(\d+)$", rest, re.I)
        if m:
            return f"{m.group(1)}/{m.group(2)}"
        m = re.match(r"^(\d+)p(\d+)$", rest, re.I)
        if m:
            return f"{m.group(1)}/{m.group(2)}"
        return rest
    if h in admin_map:
        return admin_map[h]
    # case-insensitive admin map
    for key, val in admin_map.items():
        if key.lower() == h.lower():
            return val
    return h


def derive_issue_type(
    a_report: Any,
    z_report: Any,
    a_remediation: Any,
    z_remediation: Any,
) -> str:
    """Port CVT_IMPORT AX Issue Type formula."""
    e = _s(a_report)
    t = _s(z_report)
    q = _s(a_remediation)
    ae = _s(z_remediation)
    if e == "Ok" and t == "Ok":
        return ""
    issue_text = e if e != "Ok" else t
    detail_text = q if e != "Ok" else ae
    if issue_text in {"Wrong-port", "Wrong-neighbor"}:
        return "Link Miswired"
    if issue_text == "RX/TX power mismatch":
        return "RX/TX Power Mismatch"
    if issue_text == "Media Unplugged":
        return "Link Down"
    if "Anomalous" in issue_text:
        if "BER" in detail_text or "Bit Error" in detail_text:
            return "High BER"
        return "Bad Signal Integrity"
    return "Bad Signal Integrity"


def connection_key(src_host: Any, src_port: Any, dst_host: Any, dst_port: Any) -> str:
    if not _s(src_host):
        return ""
    return KEY_SEP.join(
        [
            _clean_upper(src_host),
            _clean_upper(src_port),
            _clean_upper(dst_host),
            _clean_upper(dst_port),
        ]
    )


def issue_key(conn_key: str, issue_type: Any) -> str:
    if not conn_key:
        return ""
    return conn_key + KEY_SEP + _clean_upper(issue_type)


def triage_action(
    conn_key: str,
    iss_key: str,
    *,
    existing_issue_keys: dict[str, str],
    existing_conn_keys: set[str],
) -> str:
    """Port CVT_IMPORT BC triage formula.

    existing_issue_keys: issue_key -> status (last match wins / any)
    """
    if not iss_key:
        return ""
    if iss_key in existing_issue_keys:
        status = _s(existing_issue_keys[iss_key])
        if status.lower() == RESOLVED_STATUS:
            return "REOPEN EXISTING"
        return "ALREADY TRACKED"
    if conn_key in existing_conn_keys:
        return "ADD - NEW ISSUE TYPE"
    return "ADD - NEW CONNECTION"


@dataclass
class DerivedRow:
    protocol: str
    su: str
    src_host: str
    src_port: str
    src_nscale: str
    src_rack: str
    src_u: str
    dst_host: str
    dst_port: str
    dst_nscale: str
    dst_rack: str
    dst_u: str
    issue_type: str
    status_default: str
    conn_key: str
    iss_key: str
    triage: str
    source: dict[str, str]
    comments_nvidia: str = ""


def lookup_node(node_table: dict[str, tuple[str, str, str]], host: str) -> tuple[str, str, str]:
    """Return (nscale_name, rack, u) for host. Keys tried as-is and lower."""
    if not host:
        return ("", "", "")
    if host in node_table:
        return node_table[host]
    lower = {k.lower(): v for k, v in node_table.items()}
    return lower.get(host.lower(), ("", "", ""))


def derive_row(
    cvt: dict[str, str],
    *,
    node_table: dict[str, tuple[str, str, str]],
    admin_map: dict[str, str],
    existing_issue_keys: dict[str, str],
    existing_conn_keys: set[str],
) -> DerivedRow | None:
    src_host = _s(cvt.get("Exp A Node") or cvt.get("H") or "")
    if not src_host:
        return None
    dst_host = _s(cvt.get("Exp Z Node") or cvt.get("W") or "")
    src_port_raw = _s(cvt.get("Exp A Port") or cvt.get("I") or "")
    dst_port_raw = _s(cvt.get("Exp Z Port") or cvt.get("X") or "")

    src_port = normalize_port(src_host, src_port_raw, admin_map=admin_map)
    # Dest port formula uses dest host for m1/m2 three-part rule
    dst_port = normalize_port(dst_host, dst_port_raw, admin_map=admin_map)
    # Dest NIC m1/m2 in sheet uses lowercase for dest side in one branch; keep M1/M2 from normalize

    src_nscale, src_rack, src_u = lookup_node(node_table, src_host)
    dst_nscale, dst_rack, dst_u = lookup_node(node_table, dst_host)

    issue_type = derive_issue_type(
        cvt.get("A Report") or cvt.get("E"),
        cvt.get("Z Report") or cvt.get("T"),
        cvt.get("A Remediation Action") or cvt.get("F"),
        cvt.get("Z Remediation Action") or cvt.get("U"),
    )
    if not issue_type:
        return None

    conn = connection_key(src_host, src_port, dst_host, dst_port)
    iss = issue_key(conn, issue_type)
    triage = triage_action(
        conn,
        iss,
        existing_issue_keys=existing_issue_keys,
        existing_conn_keys=existing_conn_keys,
    )

    return DerivedRow(
        protocol=normalize_protocol(cvt.get("Protocol") or cvt.get("B")),
        su=pick_su(cvt.get("A SU Number") or cvt.get("D"), cvt.get("Z SU Number") or cvt.get("S")),
        src_host=src_host,
        src_port=src_port,
        src_nscale=src_nscale,
        src_rack=str(src_rack),
        src_u=str(src_u),
        dst_host=dst_host,
        dst_port=dst_port,
        dst_nscale=dst_nscale,
        dst_rack=str(dst_rack),
        dst_u=str(dst_u),
        issue_type=issue_type,
        status_default="New Issue",
        conn_key=conn,
        iss_key=iss,
        triage=triage,
        source={k: _s(v) for k, v in cvt.items()},
    )


def reporter_initials() -> str:
    """Build initials from the shell user, e.g. irfan.zulfiqar → IZ."""
    raw = os.environ.get("USER") or os.environ.get("USERNAME") or getpass.getuser() or "cvt"
    raw = raw.split("@")[0].strip().lower()
    parts = [p for p in re.split(r"[._\s-]+", raw) if p]
    if len(parts) >= 2:
        return f"{parts[0][0]}{parts[1][0]}".upper()
    if parts and len(parts[0]) >= 2:
        return parts[0][:2].upper()
    if parts:
        return parts[0][0].upper()
    return "CVT"


def reported_at(when: datetime | None = None) -> str:
    """Tracker Date Reported style: 09/04/2026 05:07 CT"""
    when = when or datetime.now()
    return when.strftime("%m/%d/%Y %H:%M") + " CT"


_REOPEN_ENTRY_START = re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}\b")
_REOPEN_STAMP = re.compile(
    r"^(\d{1,2})/(\d{1,2})/(\d{2,4})\s+(\d{1,2}):(\d{2})\s*(AM|PM)?",
    re.I,
)


def _split_reopen_entries(text: str) -> list[str]:
    """Split Re-open Notes into dated blocks (supports multiline entries)."""
    text = _s(text)
    if not text:
        return []
    entries: list[str] = []
    current: list[str] = []
    for line in text.splitlines():
        if _REOPEN_ENTRY_START.match(line.strip()) and current:
            block = "\n".join(current).strip()
            if block:
                entries.append(block)
            current = [line]
        else:
            current.append(line)
    if current:
        block = "\n".join(current).strip()
        if block:
            entries.append(block)
    return entries


def _reopen_entry_timestamp(entry: str) -> datetime:
    """Parse leading stamp for sort; unknown stamps sort last."""
    head = _s(entry).split("\n", 1)[0]
    match = _REOPEN_STAMP.match(head)
    if not match:
        return datetime.min
    month = int(match.group(1))
    day = int(match.group(2))
    year = int(match.group(3))
    if year < 100:
        year += 2000
    hour = int(match.group(4))
    minute = int(match.group(5))
    ampm = (match.group(6) or "").upper()
    if ampm == "PM" and hour < 12:
        hour += 12
    elif ampm == "AM" and hour == 12:
        hour = 0
    try:
        return datetime(year, month, day, hour, minute)
    except ValueError:
        return datetime.min


def _sort_reopen_notes(*entries: str) -> str:
    """Newest timestamp first (SharePoint shows the top/left of the cell)."""
    blocks: list[str] = []
    for raw in entries:
        blocks.extend(_split_reopen_entries(raw))
    seen: set[str] = set()
    unique: list[str] = []
    for block in blocks:
        # One physical line per entry — no Wrap Text / no in-cell line breaks.
        flat = " ".join(part.strip() for part in block.splitlines() if part.strip())
        if not flat or flat in seen:
            continue
        seen.add(flat)
        unique.append(flat)
    unique.sort(key=_reopen_entry_timestamp, reverse=True)
    return "\n".join(unique)


def reopen_note(
    existing: str,
    *,
    initials: str = "CVT",
    issue_type: str = "",
    a_report: str = "",
    z_report: str = "",
    a_remediation: str = "",
    z_remediation: str = "",
    hint: str = "",
    cvt: dict[str, str] | None = None,
    admin_map: dict[str, str] | None = None,
    when: datetime | None = None,
) -> str:
    """Build Re-open Notes with the newest stamp first (then older entries).

    Example (top of cell):
    09/04/2026 06:48 CT [IZ]: A Side: Fix miswiring Z Side: Fix miswiring Hint: ...
    09/01/2026 5:27AM CT (RL) - still miswired...
    08/31/2026 07:00 CT - still failing...
    """
    when = when or datetime.now()
    stamp = when.strftime("%m/%d/%Y %H:%M")
    initials = (_s(initials).strip("[]") or "CVT").upper()
    a_rem = _s(a_remediation)
    z_rem = _s(z_remediation)
    hint = _s(hint)

    if not hint:
        from cvt_sharepoint.comments import suggest_hint

        if cvt:
            hint = suggest_hint(issue_type, cvt=cvt, admin_map=admin_map)
        elif a_report or z_report:
            hint = suggest_hint(
                issue_type,
                a_report=a_report,
                z_report=z_report,
                admin_map=admin_map,
            )

    line = f"{stamp} CT [{initials}]: A Side: {a_rem} ; Z Side: {z_rem}"
    if hint:
        line = f"{line} ; Hint: {hint}"

    return _sort_reopen_notes(line, _s(existing))
