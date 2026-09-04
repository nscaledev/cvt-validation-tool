"""Smart Comments / Hints — NVIDIA engineer style + CVT A/Z rem.

Comments / reopen always include A and Z remediation text from CVT.
Hint matches tracker NVIDIA prose: report summary + optional port/BER facts
+ action (Correct per P2P / validate LLDP/agent / Clean reseat …).
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from cvt_sharepoint.remediation import COL_COMMENTS, COL_ISSUE_TYPE, DATA_START_ROW
from cvt_sharepoint.triage import _s, normalize_port

# Fallbacks mined from tracker Comments - NVIDIA (majority engineer phrasing).
DEFAULT_HINTS: dict[str, str] = {
    "Link Down": "Clean, reseat, or replace the optics on both A and Z sides.",
    "High BER": (
        "Please scope the links to determine the issue. "
        "Clean reseat, or replace the optics on both A and Z sides."
    ),
    "Bad Signal Integrity": "Clean, reseat, or replace the optics on both A and Z sides.",
    "Link Miswired": (
        "Correct the connection to the expected P2P ports, one cable at a time, then revalidate."
    ),
    "RX/TX Power Mismatch": "Clean/reseat both optics.",
    "OOB/iDRAC inaccessible": (
        "Investigate and restore the OOB/iDRAC cabling path; "
        "verify management-switch port and VLAN mapping, then confirm TCP/443 access."
    ),
    "Link Flapping": "Clean, reseat, or replace the optics on both A and Z sides.",
}

_SKIP_HINTS = {
    "",
    "will be addressed by above comment",
    "n/a",
    "na",
}

_SPLIT_STATUS = re.compile(
    r"\n(?=\d{1,2}/\d{1,2})"
    r"|(?<=\.)\s+(?=\d{1,2}/\d{1,2}/\d{2,4})"
)

_MEDIA_REPORTS = {
    "Media Unplugged",
    "Link Down, No signal",
    "Link Down",
}
_WRONG_REPORTS = {"Wrong-port", "Wrong-neighbor"}


@dataclass(frozen=True)
class SideAudit:
    """One side (A or Z) of a CVT circuit row."""

    side: str  # "A" or "Z"
    report: str
    rem: str
    exp_node: str
    exp_port: str
    disc_node: str
    disc_port: str
    port_status: str
    xcvr: str
    signal: str
    raw_ber: str
    eff_ber: str

    @property
    def label(self) -> str:
        return f"{self.side} side"

    @property
    def is_ok(self) -> bool:
        return self.report in {"", "Ok"}

    @property
    def is_unknown(self) -> bool:
        return "Unknown" in self.report

    @property
    def is_wrong(self) -> bool:
        return self.report in _WRONG_REPORTS

    @property
    def is_wrong_port(self) -> bool:
        return self.report == "Wrong-port"

    @property
    def is_wrong_neighbor(self) -> bool:
        return self.report == "Wrong-neighbor"

    @property
    def is_media(self) -> bool:
        return self.report in _MEDIA_REPORTS or (
            "Link Down" in self.report and "Anomalous" not in self.report
        )

    @property
    def is_anomalous(self) -> bool:
        return "Anomalous" in self.report

    @property
    def is_flap(self) -> bool:
        return "Flap" in self.report or "flap" in self.report.lower()

    @property
    def is_admin_down(self) -> bool:
        return "Admin Down" in self.report

    @property
    def is_no_report(self) -> bool:
        return self.report == "No Report"

    @property
    def link_up_with_xcvr(self) -> bool:
        return _truthy(self.port_status) and _truthy(self.xcvr)

    @property
    def disc_unknown(self) -> bool:
        return _is_unknown_disc(self.disc_node)

    @property
    def node_mismatch(self) -> bool:
        if self.disc_unknown or not self.exp_node or not self.disc_node:
            return False
        return self.disc_node.lower() != self.exp_node.lower()

    @property
    def port_mismatch(self) -> bool:
        if not self.exp_port or not self.disc_port:
            return False
        return self.disc_port.lower() != self.exp_port.lower()


def clean_engineer_comment(text: Any) -> str:
    """Keep the instructional part; drop later status-log lines."""
    raw = _s(text)
    if not raw:
        return ""
    parts = _SPLIT_STATUS.split(raw, maxsplit=1)
    cleaned = parts[0].strip()
    cleaned = re.split(r"\s+\d{1,2}/\d{1,2}/\d{2,4}\b", cleaned, maxsplit=1)[0].strip()
    if cleaned.lower() in _SKIP_HINTS:
        return ""
    if len(cleaned) < 12:
        return ""
    return cleaned


def mine_nvidia_hints(ws: Worksheet) -> dict[str, str]:
    """Scan CABLING_REMEDIATION for the most common clean Comments - NVIDIA per Issue Type."""
    tallies: dict[str, Counter[str]] = {}
    for row in ws.iter_rows(
        min_row=DATA_START_ROW,
        max_col=COL_COMMENTS,
        values_only=True,
    ):
        if not row or len(row) < COL_COMMENTS:
            continue
        issue_type = _s(row[COL_ISSUE_TYPE - 1])
        comment = clean_engineer_comment(row[COL_COMMENTS - 1])
        if not issue_type or not comment:
            continue
        tallies.setdefault(issue_type, Counter())[comment] += 1
    return finalize_nvidia_hints(tallies)


def finalize_nvidia_hints(tallies: dict[str, Counter[str]]) -> dict[str, str]:
    """Pick the best clean comment per issue type from tally counters."""
    hints: dict[str, str] = {}
    for issue_type, counter in tallies.items():
        best, count = counter.most_common(1)[0]
        if count >= 2 or (count == 1 and len(counter) == 1 and len(best) < 220):
            hints[issue_type] = best
    return hints


def _truthy(value: Any) -> bool:
    return _s(value).lower() in {"true", "yes", "1", "up"}


def _is_unknown_disc(value: Any) -> bool:
    text = _s(value)
    return (not text) or text.lower() == "unknown"


def _humanize_report(report: str) -> str:
    """Match tracker prose: Wrong-port → Wrong Port, Unknown-neighbor → Unknown Neighbor."""
    text = _s(report)
    if not text:
        return text
    # Drop long CVT parentheticals (e.g. Signal, Temperature) — keep the finding name.
    if "(" in text:
        text = text.split("(", 1)[0].strip()
    text = text.replace("-", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text.title()


def _side_from_cvt(cvt: dict[str, str], side: str) -> SideAudit:
    """Map CVT CSV columns (and letter aliases) onto one side."""
    if side == "A":
        return SideAudit(
            side="A",
            report=_s(cvt.get("A Report") or cvt.get("E") or ""),
            rem=_s(cvt.get("A Remediation Action") or cvt.get("F") or ""),
            exp_node=_s(cvt.get("Exp A Node") or cvt.get("H") or ""),
            exp_port=_s(cvt.get("Exp A Port") or cvt.get("I") or ""),
            disc_node=_s(cvt.get("Disc A Node") or cvt.get("J") or ""),
            disc_port=_s(cvt.get("Disc A Port") or cvt.get("K") or ""),
            port_status=_s(cvt.get("A Port Status") or cvt.get("M") or ""),
            xcvr=_s(cvt.get("A Transceiver Connected") or cvt.get("L") or ""),
            signal=_s(cvt.get("A Signal Stats") or ""),
            raw_ber=_s(cvt.get("A Raw BER") or ""),
            eff_ber=_s(cvt.get("A Eff BER") or ""),
        )
    return SideAudit(
        side="Z",
        report=_s(cvt.get("Z Report") or cvt.get("T") or ""),
        rem=_s(cvt.get("Z Remediation Action") or cvt.get("U") or ""),
        exp_node=_s(cvt.get("Exp Z Node") or cvt.get("W") or ""),
        exp_port=_s(cvt.get("Exp Z Port") or cvt.get("X") or ""),
        disc_node=_s(cvt.get("Disc Z Node") or cvt.get("Y") or ""),
        disc_port=_s(cvt.get("Disc Z Port") or ""),
        port_status=_s(cvt.get("Z Port Status") or ""),
        xcvr=_s(cvt.get("Z Transceiver Connected") or ""),
        signal=_s(cvt.get("Z Signal Stats") or ""),
        raw_ber=_s(cvt.get("Z Raw BER") or ""),
        eff_ber=_s(cvt.get("Z Eff BER") or ""),
    )


def _summarize_reports(a: SideAudit, z: SideAudit) -> str:
    """Tracker-style report sentence naming which side has which finding."""
    a_bad = not a.is_ok
    z_bad = not z.is_ok
    ah = _humanize_report(a.report)
    zh = _humanize_report(z.report)

    if a_bad and z_bad:
        if a.report == z.report:
            return f"Both sides report {ah}"
        return f"A side reports {ah} and Z side reports {zh}"
    if a_bad:
        if z.report == "Ok":
            return f"A side reports {ah} (Z side Ok)"
        return f"A side reports {ah}"
    if z_bad:
        if a.report == "Ok":
            return f"Z side reports {zh} (A side Ok)"
        return f"Z side reports {zh}"
    return ""


def _norm_port(
    host: str,
    port: str,
    *,
    admin_map: dict[str, str] | None,
) -> str:
    """CVT_IMPORT port display (swp22s1 → 22/2). Empty if no port."""
    if not _s(port):
        return ""
    # normalize_port needs a host for m1/m2 slash rules; fall back to a dummy ETH host.
    return normalize_port(host or "host-t0", port, admin_map=admin_map)


def _useful_disc_host(host: str) -> bool:
    h = _s(host).lower()
    return bool(h) and h not in {"unknown", "localhost", "localhost.localdomain"}


def _useful_disc_port(port: str) -> bool:
    """Skip agent/NIC noise (pf0hpf) — keep switch-style ports only."""
    p = _s(port).lower()
    if not p:
        return False
    if "hpf" in p or p.startswith("pf"):
        return False
    return True


def _discovery_bits(
    side: SideAudit,
    *,
    admin_map: dict[str, str] | None,
) -> list[str]:
    """Expected vs discovered host/port — only real mismatches (ports normalized)."""
    bits: list[str] = []
    # Host mismatch (Wrong Neighbor) — skip localhost / Unknown noise unless expected is useful.
    if side.node_mismatch and _useful_disc_host(side.disc_node):
        bits.append(
            f"{side.side} discovered {side.disc_node} (expected {side.exp_node})"
        )
    elif side.is_wrong_neighbor and side.disc_unknown and side.exp_node:
        bits.append(f"{side.side} discovered Unknown (expected {side.exp_node})")

    # Port swap only when Disc and Exp both exist and differ (normalized).
    if side.exp_port and side.disc_port:
        exp_host = side.exp_node
        disc_host = side.disc_node if not side.disc_unknown else side.exp_node
        exp_p = _norm_port(exp_host, side.exp_port, admin_map=admin_map)
        disc_p = _norm_port(disc_host, side.disc_port, admin_map=admin_map)
        if (
            exp_p
            and disc_p
            and exp_p != disc_p
            and _useful_disc_port(disc_p)
            and _useful_disc_port(exp_p)
        ):
            bits.append(f"{side.side} port {disc_p} (expected {exp_p})")
    return bits


def _discovery_evidence(
    a: SideAudit,
    z: SideAudit,
    *,
    admin_map: dict[str, str] | None,
) -> str:
    # Miswire rows, or peer port-swap next to Wrong-port / Wrong-neighbor.
    if not (a.is_wrong or z.is_wrong):
        return ""
    bits = _discovery_bits(a, admin_map=admin_map) + _discovery_bits(z, admin_map=admin_map)
    return "; ".join(bits)


def _side_has_port_swap(side: SideAudit, admin_map: dict[str, str] | None = None) -> bool:
    if not side.exp_port or not side.disc_port:
        return False
    exp_host = side.exp_node
    disc_host = side.disc_node if not side.disc_unknown else side.exp_node
    exp_p = _norm_port(exp_host, side.exp_port, admin_map=admin_map)
    disc_p = _norm_port(disc_host, side.disc_port, admin_map=admin_map)
    return bool(exp_p and disc_p and exp_p != disc_p)


def _useful_ber(ber: str) -> bool:
    text = _s(ber)
    if not text or text.upper() == "N/A":
        return False
    try:
        value = float(text)
    except ValueError:
        return False
    # Ignore zero / CVT sentinel readings.
    return value > 0 and value >= 1e-200


def _fmt_ber(ber: str) -> str:
    """Compact BER for comments (3.000000e-09 → 3e-09)."""
    text = _s(ber)
    try:
        value = float(text)
    except ValueError:
        return text
    formatted = f"{value:.0e}"
    return formatted.replace("e-0", "e-").replace("e+0", "e+")


def _ber_evidence(a: SideAudit, z: SideAudit, issue_type: str) -> str:
    """Which side has BER trouble — only when signal/BER is relevant."""
    care = (
        a.is_anomalous
        or z.is_anomalous
        or issue_type in {"High BER", "Bad Signal Integrity", "RX/TX Power Mismatch"}
    )
    if not care:
        return ""

    bits: list[str] = []
    for side in (a, z):
        show = side.is_anomalous or (
            issue_type in {"High BER", "RX/TX Power Mismatch"} and _useful_ber(side.raw_ber)
        )
        if not show and issue_type == "Bad Signal Integrity":
            show = side.is_anomalous and _useful_ber(side.raw_ber)
        if show and _useful_ber(side.raw_ber):
            bits.append(f"{side.side} BER={_fmt_ber(side.raw_ber)}")

    # Anomalous side without usable BER still gets named by the report summary.
    if not bits:
        return ""
    # If one side is bad and peer has a clearly better BER, show peer for contrast.
    if len(bits) == 1:
        bad = a if bits[0].startswith("A ") else z
        peer = z if bad.side == "A" else a
        if peer.is_ok and _useful_ber(peer.raw_ber):
            bits.append(f"{peer.side} BER={_fmt_ber(peer.raw_ber)}")
    return "; ".join(bits)


def _priority_when_rem_present(
    a: SideAudit,
    z: SideAudit,
    issue_type: str,
    *,
    admin_map: dict[str, str] | None = None,
) -> str:
    """When A/Z rem already states the fix, Hint adds priority only (no restated action)."""
    a_swap = _side_has_port_swap(a, admin_map)
    z_swap = _side_has_port_swap(z, admin_map)

    # Miswire: rem is usually just "Fix miswiring" — keep full P2P/LLDP guidance.
    if a.is_wrong or z.is_wrong:
        return _action_guidance(a, z, issue_type, admin_map=admin_map)

    # Media Unplugged is primary; peer Link Down is secondary (rem already says insert/reseat).
    if a.report == "Media Unplugged" and z.is_media:
        return "Start on A; Z Link Down is likely secondary."
    if z.report == "Media Unplugged" and a.is_media:
        return "Start on Z; A Link Down is likely secondary."

    if a.is_media and z.is_flap:
        return "Physical fix A first; then Network Admin for Z ErrDisable/flap."
    if z.is_media and a.is_flap:
        return "Physical fix Z first; then Network Admin for A ErrDisable/flap."

    if a.is_media and z.is_unknown:
        return "Start on A link/optics; then validate Z LLDP/agent."
    if z.is_media and a.is_unknown:
        return "Start on Z link/optics; then validate A LLDP/agent."

    # Unknown Neighbor: rem already says Check LLDP — don't restate.
    if a.is_unknown or z.is_unknown:
        return ""

    # Clean/reseat already in rem — summary (+ BER) is enough.
    if a.is_media or z.is_media or a.is_anomalous or z.is_anomalous:
        return ""

    if a.is_flap or z.is_flap:
        side = "A" if a.is_flap else "Z"
        return f"{side} ErrDisable/flap — Network Admin if it remains after clean/reseat."
    if a.is_admin_down or z.is_admin_down:
        side = "A" if a.is_admin_down else "Z"
        return f"{side} Admin Down — contact Network Admin after verifying cabling."
    if "NIC" in a.report or "NIC" in z.report:
        return "Complete NIC provisioning, then revalidate."
    if a.is_no_report or z.is_no_report:
        side = "A" if a.is_no_report else "Z"
        return f"{side} No Report — verify agent reachability."

    # Port-swap nuance when rem didn't cover both ends.
    if a_swap and z_swap:
        return "Both ports look swapped — correct one cable at a time, then revalidate."
    return ""


def _action_guidance(
    a: SideAudit,
    z: SideAudit,
    issue_type: str,
    *,
    admin_map: dict[str, str] | None = None,
) -> str:
    """Action clause matching tracker NVIDIA engineer phrasing."""
    a_swap = _side_has_port_swap(a, admin_map)
    z_swap = _side_has_port_swap(z, admin_map)

    # --- Miswire / neighbor (top engineer templates) ---
    if a.is_wrong and z.is_wrong:
        return (
            "Correct the connection to the expected P2P ports, "
            "one cable at a time, then revalidate."
        )
    if a.is_wrong and z.is_unknown:
        return (
            "Correct the A-side connection per P2P and validate Z-side LLDP/agent reporting."
        )
    if a.is_unknown and z.is_wrong:
        return (
            "Correct the Z-side connection per P2P and validate A-side LLDP/agent reporting."
        )
    if a.is_wrong and z.is_ok:
        return "Correct the A-side connection to the expected P2P port, then revalidate."
    if z.is_wrong and a.is_ok:
        return "Correct the Z-side connection to the expected P2P port, then revalidate."

    # Wrong + link-down/media on peer.
    if a.is_wrong and z.is_media:
        if z_swap:
            return (
                "Correct both ends to the expected P2P ports; "
                "then insert/reseat the Z-side cable/transceiver."
            )
        return (
            "Correct the A-side connection per P2P, then revalidate the link."
        )
    if z.is_wrong and a.is_media:
        if a_swap:
            return (
                "Correct both ends to the expected P2P ports; "
                "then insert/reseat the A-side cable/transceiver."
            )
        return (
            "Correct the Z-side connection per P2P, then revalidate the link."
        )
    if a.is_wrong and z.is_flap:
        return "Correct the A-side connection per P2P, then revalidate the link."
    if z.is_wrong and a.is_flap:
        return "Correct the Z-side connection per P2P, then revalidate the link."
    if a.is_wrong_neighbor and z.is_media:
        return (
            "Correct the connection per P2P and insert/reseat the Z-side cable/transceiver."
        )
    if z.is_wrong_neighbor and a.is_media:
        return (
            "Correct the connection per P2P and insert/reseat the A-side cable/transceiver."
        )

    # Media + Unknown Neighbor.
    if a.is_media and z.is_unknown:
        return (
            "Insert, reseat, or replace the A-side cable/transceiver; "
            "validate Z-side LLDP/agent reporting."
        )
    if z.is_media and a.is_unknown:
        return (
            "Insert, reseat, or replace the Z-side cable/transceiver; "
            "validate A-side LLDP/agent reporting."
        )

    # Media + ErrDisable flap.
    if a.is_media and z.is_flap:
        return (
            "Clean, reseat, or replace the A-side optics; "
            "contact Network Admin for Z-side ErrDisable/flap after physical fix."
        )
    if z.is_media and a.is_flap:
        return (
            "Clean, reseat, or replace the Z-side optics; "
            "contact Network Admin for A-side ErrDisable/flap after physical fix."
        )

    # Unknown Neighbor (LLDP/agent — engineer preferred wording).
    if a.is_unknown and z.is_unknown:
        return (
            "Validate LLDP/agent reporting on both A and Z; "
            "confirm both peers are fully provisioned and reachable."
        )
    if a.is_unknown and z.is_ok:
        return (
            "Validate A-side LLDP/agent reporting; "
            "confirm the expected peer is provisioned and reachable."
        )
    if z.is_unknown and a.is_ok:
        return (
            "Validate Z-side LLDP/agent reporting; "
            "confirm the expected peer is provisioned and reachable."
        )
    if a.is_unknown:
        return "Validate A-side LLDP/agent reporting and peer reachability."
    if z.is_unknown:
        return "Validate Z-side LLDP/agent reporting and peer reachability."

    # Media Unplugged primary (tracker BSI/Link Down pattern).
    if a.report == "Media Unplugged" and z.is_media:
        return (
            "Insert, reseat, or replace the cable/transceiver on the A side. "
            "Z-side Link Down is likely secondary."
        )
    if z.report == "Media Unplugged" and a.is_media:
        return (
            "Insert, reseat, or replace the cable/transceiver on the Z side. "
            "A-side Link Down is likely secondary."
        )
    if a.is_media and z.is_media:
        if issue_type == "Link Down":
            return "Clean, reseat, or replace the optics on both A and Z sides."
        return "Clean, reseat, or replace the optics/cable on both A and Z sides."
    if a.is_media and z.is_ok:
        return "Clean, reseat, or replace the A-side optics."
    if z.is_media and a.is_ok:
        return "Clean, reseat, or replace the Z-side optics."
    if a.is_media:
        return "Clean, reseat, or replace the A-side optics."
    if z.is_media:
        return "Clean, reseat, or replace the Z-side optics."

    # High BER / anomalous — match engineer side-specific clean/reseat.
    if a.is_anomalous and z.is_anomalous:
        if issue_type == "High BER":
            return (
                "Please scope the links to determine the issue. "
                "Clean reseat, or replace the optics on both A and Z sides."
            )
        return "Clean, reseat, or replace the optics on both A and Z sides."
    if a.is_anomalous and (z.is_ok or z.is_no_report):
        return "Clean and reseat the A-side optics."
    if z.is_anomalous and (a.is_ok or a.is_no_report):
        return "Clean and reseat the Z-side optics."
    if a.is_anomalous:
        return "Clean and reseat the A-side optics."
    if z.is_anomalous:
        return "Clean and reseat the Z-side optics."

    if a.is_flap or z.is_flap:
        side = "A" if a.is_flap else "Z"
        return (
            f"Clean, reseat, or replace the {side}-side optics; "
            "if ErrDisable remains, contact Network Admin."
        )
    if a.is_admin_down or z.is_admin_down:
        side = "A" if a.is_admin_down else "Z"
        return (
            f"{side} side is Admin Down — contact Network Admin after verifying cabling."
        )
    if "NIC" in a.report or "NIC" in z.report:
        return (
            "Complete NIC provisioning so interface names match expected, then revalidate."
        )
    if a.is_no_report or z.is_no_report:
        side = "A" if a.is_no_report else "Z"
        return (
            f"{side} side has No Report — verify agent reachability and peer provisioning."
        )

    return DEFAULT_HINTS.get(issue_type, "Revalidate in CVT after remediation.")


def build_smart_hint(
    cvt: dict[str, str],
    issue_type: str = "",
    *,
    admin_map: dict[str, str] | None = None,
) -> str:
    """NVIDIA-style Hint: reports + port/BER evidence + action/priority.

    When A/Z rem text is present (Comments / reopen), skip restating the same
    clean/insert action — only add finding summary, facts, and which-side-first.
    """
    a = _side_from_cvt(cvt, "A")
    z = _side_from_cvt(cvt, "Z")
    issue_type = _s(issue_type)
    has_rem = bool(a.rem or z.rem)

    parts: list[str] = []
    summary = _summarize_reports(a, z)
    if summary:
        parts.append(summary)

    evidence = _discovery_evidence(a, z, admin_map=admin_map)
    if evidence:
        parts.append(evidence)

    ber = _ber_evidence(a, z, issue_type)
    if ber:
        parts.append(ber)

    if has_rem:
        action = _priority_when_rem_present(a, z, issue_type, admin_map=admin_map)
    else:
        action = _action_guidance(a, z, issue_type, admin_map=admin_map)

    body = ". ".join(p.rstrip(".").strip() for p in parts if p)
    if body and action:
        return f"{body}. {action}"
    if action:
        return action
    return body or ("" if has_rem else DEFAULT_HINTS.get(issue_type, ""))


def suggest_hint(
    issue_type: str,
    *,
    a_report: Any = "",
    z_report: Any = "",
    mined: dict[str, str] | None = None,
    cvt: dict[str, str] | None = None,
    admin_map: dict[str, str] | None = None,
) -> str:
    """Prefer row-aware CVT hint; mined/default only as last resort."""
    if cvt:
        return build_smart_hint(cvt, issue_type, admin_map=admin_map)
    a_report = _s(a_report)
    z_report = _s(z_report)
    if a_report or z_report:
        return build_smart_hint(
            {"A Report": a_report, "Z Report": z_report},
            issue_type,
            admin_map=admin_map,
        )
    mined = mined or {}
    if issue_type in mined:
        return mined[issue_type]
    return DEFAULT_HINTS.get(issue_type, "")


def format_nvidia_comments(a_remediation: Any, z_remediation: Any, hint: str) -> str:
    """A/Z rem from CVT + NVIDIA-style Hint (single line)."""
    a = _s(a_remediation).replace("\t", " ").replace("\n", " ").replace("\r", " ")
    z = _s(z_remediation).replace("\t", " ").replace("\n", " ").replace("\r", " ")
    hint = _s(hint).replace("\t", " ").replace("\n", " ").replace("\r", " ")
    chunks: list[str] = []
    # Always show both rem slots so field techs see A and Z (blank if CVT empty).
    chunks.append(f"A side: {a}" if a else "A side:")
    chunks.append(f"Z Side: {z}" if z else "Z Side:")
    if hint:
        chunks.append(f"Hint: {hint}")
    return " ; ".join(chunks)


def comments_for_cvt_row(
    cvt: dict[str, str],
    issue_type: str,
    *,
    mined: dict[str, str] | None = None,
    admin_map: dict[str, str] | None = None,
) -> str:
    """Full Comments - NVIDIA string for adds / handoff paste."""
    a_rem = cvt.get("A Remediation Action") or cvt.get("F") or ""
    z_rem = cvt.get("Z Remediation Action") or cvt.get("U") or ""
    hint = suggest_hint(issue_type, mined=mined, cvt=cvt, admin_map=admin_map)
    return format_nvidia_comments(a_rem, z_rem, hint)
