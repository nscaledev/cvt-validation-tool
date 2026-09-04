cvt-validation-tool
===================

Exports NVIDIA Cable Validation Tool **Circuits View** rows to CSV.

Setup
-----

Dependencies install themselves. The first run creates a virtualenv named
``.venv`` in this folder and ``pip install``s ``requirements.txt`` (msal,
openpyxl, …). Later runs reuse ``.venv`` and skip pip unless
``requirements.txt`` changed. Do not brew-install those Python libraries.

Team members only need Python 3.9+ on the PATH. If ``python3`` is missing
and Homebrew is installed, ``./cvt`` will run ``brew install python``.

1. In this folder (next to `README.md`), copy the example env file and set
   your CVT UI password:

```bash
cp .env.example .env
vim .env
```

Paste your CVT UI password on the `CVT_PASSWORD=` line:

```
CVT_URL=https://localhost:9443
CVT_USERNAME=nscale
CVT_PASSWORD=your-cvt-ui-password
CVT_INSECURE=true
```

Save and quit vim with `:wq!` (type that, then press Enter).

`.env` is gitignored. `git pull` will not overwrite it. Do not commit it.

2. Keep the Teleport tunnel up so `https://localhost:9443` reaches CVT:

```bash
tsh ssh -L 9443:192.168.10.3:9443 first_name.last_name@mobilekit-p-phy-device100
```

Default export filters
----------------------

`./cvt circuits circuits` applies all of these (same as Circuits View):

| Filter | Value | Flag |
|---|---|---|
| Resource | Data Center | `--filter dc` |
| Status | Fail | `--status Fail` |
| Protocol | ethernet | `--protocol ethernet` |
| A Report | Does **not** contain `No Report` | `--a-report-not-contains "No Report"` |
| Issues only | unhealthy circuits | built-in (`healthy=false`) |
| SU number | skip `CORE` | built-in |

There is no extra hall / SU / location flag. Data Center coverage is built by
walking every SU number except `CORE` (compute halls).

Run
---

With correct filter for ethernet switches:

```bash
./cvt circuits circuits \
  --filter dc \
  --status Fail \
  --protocol ethernet \
  --a-report-not-contains "No Report" \
  --out-dir out \
  --csv circuits-fail-ethernet-dc.csv
```

`python3 -m cvt_circuits ...` also works; it uses the same ``.venv``.

Credentials come from `.env` (`CVT_USERNAME` / `CVT_PASSWORD`). Do not put the
password on the command line.

Each run writes a **new** file with a timestamp, so the previous CSV is not
replaced:

`out/circuits-fail-ethernet-dc-YYYYMMDD-HHMMSS.csv`

Example: `out/circuits-fail-ethernet-dc-20260902-124605.csv`

When the run finishes it prints a table of remaining Fail+ethernet rows per SU
(not the JSON dump). Progress lines only list each SU name so the counts are
not repeated.

Other examples
--------------

Include A Report `No Report` rows:

```bash
./cvt circuits circuits \
  --filter dc \
  --status Fail \
  --protocol ethernet \
  --a-report-not-contains "" \
  --csv circuits-fail-ethernet-including-no-report.csv
```

DC counts only (no CSV; Data Center filter only):

```bash
./cvt circuits stats
```

```bash
./cvt circuits circuits --help
```

Notes
-----

On the WC TX 16K collector a single `context=dc` circuits request times
out (~353k circuits). This script discovers SU numbers per data hall
(`GET /cablevalidation/resources/scalable_units?data_hall=<hall>`), then
walks each SU (`context=su&items=<hall>/<su>`) and de-duplicates
`circuit_id`.

Docs: [Rest APIs 2.0.1](https://networking-docs.nvidia.com/cablevalidationtool/2.0.1/rest-apis),
[Reports APIs](https://networking-docs.nvidia.com/cablevalidationtool/2.0.1/reports-apis),
[Resource Filter](https://networking-docs.nvidia.com/cablevalidationtool/2.0.1/resource-filter).

SharePoint tracker prepare (local working copy + HTML handoff)
--------------------------------------------------------------

Each audit run builds a **capped handoff** for manual SharePoint paste
(no full-tab replace, formulas untouched):

- up to **20** ``REOPEN EXISTING`` rows (copy each TSV → paste on ``U&lt;row&gt;``)
- up to **300** add rows (one TSV block → paste at bottom starting at column A)

```bash
./cvt circuits circuits --out-dir out --csv circuits-fail-ethernet-dc.csv
./cvt sharepoint prepare
```

Outputs in ``out/``:

- ``handoff-*.html`` (preferred — copy buttons)
- ``handoff-*.xlsx``

Tracker download in ``out/`` is read-only (not modified, no backup copies).

Re-run prepare on later audits to drain remaining reopens/adds 20 + 300 at a time.

Optional:

```bash
./cvt sharepoint prepare --tracker "out/Nscale_WC_Cabling_HW_Remediation_Tracker_16k - New.xlsx" --csv out/circuits-fail-ethernet-dc-20260902-140237.csv --no-browser
```

Or omit ``--csv`` to auto-use the newest ``circuits*.csv`` in ``out/``.

