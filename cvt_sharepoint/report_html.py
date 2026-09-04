"""Fancy HTML handoff with TSV copy buttons (reopens + adds).

HTML tables and TSV match Excel handoff / CABLING_REMEDIATION columns A–Z.
"""

from __future__ import annotations

import html
import json
from datetime import datetime
from pathlib import Path
from typing import Any

from cvt_sharepoint.handoff import REMEDIATION_HEADERS, HandoffResult


def _e(value: Any) -> str:
    return html.escape("" if value is None else str(value))


def write_handoff_html(path: Path, result: HandoffResult, *, xlsx_name: str = "") -> Path:
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    headers = list(REMEDIATION_HEADERS)

    reopen_payload = [
        {
            "issue_no": r["issue_no"],
            "excel_row": r["excel_row"],
            "paste_at": r["paste_at"],
            "tsv": r["tsv"],
            "values": [("" if v is None else str(v)) for v in r["values"]],
            "new_status": r["new_status"],
            "new_count": r["new_count"],
            "old_status": r["old_status"],
            "src_host": r["src_host"],
            "src_port": r["src_port"],
            "dst_host": r["dst_host"],
            "dst_port": r["dst_port"],
            "issue_type": r["issue_type"],
        }
        for r in result.reopens
    ]
    adds_payload = [
        {
            "issue_no": r["issue_no"],
            "values": [("" if v is None else str(v)) for v in r["values"]],
            "src_host": r["src_host"],
            "src_port": r["src_port"],
            "dst_host": r["dst_host"],
            "dst_port": r["dst_port"],
            "issue_type": r["issue_type"],
            "comments": r.get("comments", ""),
        }
        for r in result.adds
    ]

    reopen_left = max(0, result.reopen_total - len(result.reopens))
    add_left = max(0, result.add_total - len(result.adds))
    header_cells = "".join(f"<th>{_e(h)}</th>" for h in headers)

    doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>CVT audit handoff</title>
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin/>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet"/>
<style>
:root {{
  --bg0: #05070d;
  --bg1: #0b1020;
  --card: #121826;
  --line: #2a354d;
  --text: #f2f5ff;
  --muted: #9aa8c7;
  --accent: #d6a4ff;
  --accent2: #ffb4d9;
  --danger: #ff8fab;
  --ok: #9dffb0;
  --shadow: 0 18px 50px rgba(0,0,0,.45);
}}
* {{ box-sizing: border-box; }}
body {{
  margin: 0;
  font-family: "DM Sans", system-ui, sans-serif;
  color: var(--text);
  background: radial-gradient(900px 500px at 0% 0%, #1a1030 0%, transparent 55%),
              radial-gradient(800px 480px at 100% 0%, #2a1020 0%, transparent 50%),
              linear-gradient(180deg, var(--bg0), var(--bg1));
  min-height: 100vh;
}}
.wrap {{ max-width: 1400px; margin: 0 auto; padding: 1.5rem 1.1rem 4rem; }}
.hero {{
  display: flex; flex-wrap: wrap; gap: 1rem; justify-content: space-between; align-items: end;
  margin-bottom: 1.25rem;
}}
h1 {{ margin: 0; font-size: clamp(1.6rem, 3vw, 2.2rem); letter-spacing: -0.03em; }}
.sub {{ color: var(--muted); margin-top: .35rem; }}
.stats {{ display: flex; flex-wrap: wrap; gap: .5rem; }}
.stat {{
  background: rgba(214,164,255,.12);
  border: 1px solid rgba(214,164,255,.35);
  color: var(--accent);
  padding: .35rem .7rem;
  border-radius: 999px;
  font-size: .85rem;
  font-weight: 600;
}}
.stat.warn {{ background: rgba(255,180,217,.12); border-color: rgba(255,180,217,.4); color: var(--accent2); }}
.card {{
  background: linear-gradient(180deg, rgba(255,255,255,.03), transparent), var(--card);
  border: 1px solid var(--line);
  border-radius: 16px;
  box-shadow: var(--shadow);
  padding: 1rem 1.1rem;
  margin: 1rem 0 1.4rem;
}}
.card h2 {{
  margin: 0 0 .75rem;
  font-size: 1.15rem;
  display: flex; align-items: center; justify-content: space-between; gap: .75rem; flex-wrap: wrap;
}}
.hint {{ color: var(--muted); font-size: .92rem; margin: 0 0 1rem; }}
.btn {{
  appearance: none; border: 0; cursor: pointer;
  font: inherit; font-weight: 700;
  border-radius: 10px;
  padding: .45rem .75rem;
  background: var(--accent);
  color: #1a0b28;
  transition: transform .12s ease, filter .12s ease;
  white-space: nowrap;
}}
.btn:hover {{ filter: brightness(1.08); transform: translateY(-1px); }}
.btn.done {{ background: var(--ok); color: #062816; }}
.btn.ghost {{ background: rgba(255,255,255,.06); color: var(--text); border: 1px solid var(--line); }}
.btn.secondary {{ background: transparent; color: var(--text); border: 1px solid var(--line); }}
.progress {{
  height: 8px; border-radius: 999px; background: #0d162b; overflow: hidden; margin: .4rem 0 .8rem;
}}
.progress > span {{ display: block; height: 100%; width: 0; background: linear-gradient(90deg, var(--accent), var(--accent2)); transition: width .25s ease; }}
.scroll {{ max-height: 520px; overflow: auto; border: 1px solid var(--line); border-radius: 10px; }}
table {{ width: max-content; min-width: 100%; border-collapse: collapse; font-size: .72rem; }}
th, td {{ border-bottom: 1px solid var(--line); padding: .35rem .4rem; text-align: left; vertical-align: top; max-width: 220px; }}
th {{ color: var(--muted); font-weight: 600; position: sticky; top: 0; background: #161e30; z-index: 2; }}
td {{ font-family: "JetBrains Mono", monospace; color: #dce6ff; word-break: break-word; }}
td.sticky, th.sticky {{ position: sticky; left: 0; background: #161e30; z-index: 3; }}
td.sticky {{ background: #121826; }}
th.guide, td.guide {{ color: var(--accent2); font-style: italic; max-width: 90px; }}
.toast {{
  position: fixed; right: 1rem; bottom: 1rem; z-index: 50;
  background: #0e1b14; color: var(--ok);
  border: 1px solid rgba(157,255,176,.45);
  padding: .7rem 1rem; border-radius: 10px;
  transform: translateY(120%); opacity: 0;
  transition: .2s ease;
  font-weight: 600;
}}
.toast.show {{ transform: translateY(0); opacity: 1; }}
.topbar {{
  position: sticky; top: 0; z-index: 20;
  backdrop-filter: blur(10px);
  background: rgba(5,7,13,.88);
  border-bottom: 1px solid rgba(42,53,77,.8);
  margin: -1.5rem -1.1rem 1.2rem;
  padding: .75rem 1.1rem;
  display: flex; justify-content: space-between; gap: .75rem; flex-wrap: wrap; align-items: center;
}}
.actions {{ display: flex; flex-wrap: wrap; gap: .45rem; }}
a {{ color: var(--accent); }}
</style>
</head>
<body>
<div class="wrap">
  <div class="topbar">
    <div>
      <strong>CVT → SharePoint handoff</strong>
      <div class="sub" style="margin:0">{_e(stamp)}</div>
    </div>
    <div class="actions">
      <a class="btn secondary" id="btn-excel" href="{_e(xlsx_name)}" download>Download Excel</a>
      <button class="btn ghost" id="btn-reset" type="button">Reset copied</button>
    </div>
  </div>

  <div class="hero">
    <div>
      <h1>Paste into SharePoint</h1>
    </div>
    <div class="stats">
      <div class="stat">Reopen {len(result.reopens)} / {result.reopen_total}</div>
      <div class="stat">Adds {len(result.adds)} / {result.add_total}</div>
      <div class="stat warn">Next run left: {reopen_left} reopen · {add_left} adds</div>
    </div>
  </div>

  <section class="card" id="reopen-section">
    <h2>
      <span>1 · Update existing rows (reopen) <span id="reopen-count">0/{len(result.reopens)}</span></span>
    </h2>
    <p class="hint">Copy button excludes &quot;Excel row&quot; and &quot;Paste at&quot; column. In Sharepoint, only update data from A-Z column leave AA AB columns as it is</p>
    <div class="progress"><span id="reopen-bar"></span></div>
    <div class="scroll">
      <table>
        <thead>
          <tr>
            <th class="sticky">Copy</th>
            <th class="guide">Excel row</th>
            <th class="guide">Paste at</th>
            {header_cells}
          </tr>
        </thead>
        <tbody id="reopen-body"></tbody>
      </table>
    </div>
  </section>

  <section class="card" id="adds-section">
    <h2>
      <span>2 · New rows (adds) ({len(result.adds)})</span>
      <button class="btn" id="copy-adds" type="button">Copy all adds (TSV)</button>
    </h2>
    <div class="scroll">
      <table>
        <thead>
          <tr>
            <th class="sticky">Copy</th>
            {header_cells}
          </tr>
        </thead>
        <tbody id="adds-body"></tbody>
      </table>
    </div>
  </section>
</div>
<div class="toast" id="toast">Copied</div>
<script>
const HEADERS = {json.dumps(headers)};
const REOPENS = {json.dumps(reopen_payload)};
const ADDS = {json.dumps(adds_payload)};
const EXCEL_NAME = {json.dumps(xlsx_name)};
const KEY = "cvt-handoff-copied:" + (EXCEL_NAME || location.pathname);

const toastEl = document.getElementById("toast");
function toast(msg) {{
  toastEl.textContent = msg;
  toastEl.classList.add("show");
  setTimeout(() => toastEl.classList.remove("show"), 1600);
}}

async function copyText(text) {{
  try {{
    await navigator.clipboard.writeText(text);
    return true;
  }} catch (err) {{
    const ta = document.createElement("textarea");
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    const ok = document.execCommand("copy");
    ta.remove();
    return ok;
  }}
}}

function loadCopied() {{
  try {{ return new Set(JSON.parse(localStorage.getItem(KEY) || "[]")); }}
  catch {{ return new Set(); }}
}}
function saveCopied(set) {{
  localStorage.setItem(KEY, JSON.stringify([...set]));
}}

let copied = loadCopied();

function esc(s) {{
  return String(s ?? "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}}

function cellHtml(v, idx) {{
  const title = esc(String(v ?? "").replace(/\\r?\\n/g, " "));
  return `<td title="${{title}}">${{esc(v)}}</td>`;
}}

function plainCell(v) {{
  // Flatten newlines so SharePoint paste stays in one cell (no Wrap Text).
  return String(v ?? "").replace(/\\t/g, " ").replace(/\\r?\\n/g, " | ");
}}

function pasteTsv(row) {{
  return (row.values || []).map(plainCell).join("\\t");
}}

function htmlTable(rows) {{
  const body = rows.map(row => {{
    const tds = (row.values || []).map(v => {{
      const raw = plainCell(v);
      return `<td>${{esc(raw)}}</td>`;
    }}).join("");
    return `<tr>${{tds}}</tr>`;
  }}).join("");
  return `<html><body><table>${{body}}</table></body></html>`;
}}

async function copyRows(rows) {{
  if (!rows.length) return false;
  const html = htmlTable(rows);
  const plain = rows.map(pasteTsv).join("\\n");
  try {{
    if (window.ClipboardItem && navigator.clipboard.write) {{
      await navigator.clipboard.write([
        new ClipboardItem({{
          "text/html": new Blob([html], {{ type: "text/html" }}),
          "text/plain": new Blob([plain], {{ type: "text/plain" }}),
        }}),
      ]);
      return true;
    }}
  }} catch (err) {{
    /* fall through to plain */
  }}
  return copyText(plain);
}}

function renderReopens() {{
  const body = document.getElementById("reopen-body");
  body.innerHTML = REOPENS.map((row, idx) => {{
    const id = String(row.issue_no) + ":" + String(row.excel_row);
    const done = copied.has(id);
    const cells = (row.values || []).map((v, i) => cellHtml(v, i)).join("");
    return `<tr class="${{done ? "copied" : ""}}">
      <td class="sticky"><button class="btn ${{done ? "done" : ""}}" data-reopen="${{idx}}" type="button">${{done ? "Copied ✓" : "Copy row"}}</button></td>
      <td class="guide">${{esc(row.excel_row)}}</td>
      <td class="guide">${{esc(row.paste_at)}}</td>
      ${{cells}}
    </tr>`;
  }}).join("");

  body.querySelectorAll("[data-reopen]").forEach(btn => {{
    btn.addEventListener("click", async () => {{
      const idx = Number(btn.getAttribute("data-reopen"));
      const row = REOPENS[idx];
      const ok = await copyRows([row]);
      if (!ok) {{ toast("Copy failed"); return; }}
      const id = String(row.issue_no) + ":" + String(row.excel_row);
      copied.add(id);
      saveCopied(copied);
      renderReopens();
      updateProgress();
      toast(`Copied → paste on ${{row.paste_at}}`);
    }});
  }});
  updateProgress();
}}

function updateProgress() {{
  const total = REOPENS.length || 1;
  const n = REOPENS.filter(r => copied.has(String(r.issue_no) + ":" + String(r.excel_row))).length;
  document.getElementById("reopen-count").textContent = n + "/" + REOPENS.length;
  document.getElementById("reopen-bar").style.width = ((n / total) * 100) + "%";
}}

function addId(row) {{
  return "add:" + String(row.issue_no);
}}

function renderAdds() {{
  const body = document.getElementById("adds-body");
  body.innerHTML = ADDS.map((row, idx) => {{
    const id = addId(row);
    const done = copied.has(id);
    const cells = (row.values || []).map((v, i) => cellHtml(v, i)).join("");
    return `<tr class="${{done ? "copied" : ""}}">
      <td class="sticky"><button class="btn ${{done ? "done" : "ghost"}}" data-add="${{idx}}" type="button">${{done ? "Copied ✓" : "Copy row"}}</button></td>
      ${{cells}}
    </tr>`;
  }}).join("");

  body.querySelectorAll("[data-add]").forEach(btn => {{
    btn.addEventListener("click", async () => {{
      const idx = Number(btn.getAttribute("data-add"));
      const row = ADDS[idx];
      const ok = await copyRows([row]);
      if (!ok) {{ toast("Copy failed"); return; }}
      copied.add(addId(row));
      saveCopied(copied);
      renderAdds();
      toast("Copied");
    }});
  }});
}}

document.getElementById("copy-adds").addEventListener("click", async () => {{
  if (!ADDS.length) {{ toast("No adds in this batch"); return; }}
  const ok = await copyRows(ADDS);
  if (!ok) {{ toast("Copy failed"); return; }}
  ADDS.forEach(row => copied.add(addId(row)));
  saveCopied(copied);
  renderAdds();
  toast(`Copied ${{ADDS.length}} rows`);
}});

document.getElementById("btn-reset").addEventListener("click", () => {{
  copied = new Set();
  saveCopied(copied);
  renderReopens();
  renderAdds();
  toast("Reset copied state");
}});

renderReopens();
renderAdds();
</script>
</body>
</html>
"""
    path.write_text(doc, encoding="utf-8")
    return path
