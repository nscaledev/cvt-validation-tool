cvt-validation-tool
===================

NVIDIA CVT Circuits View export.

Pulls Circuits View data from the Cable Validation Tool REST API with the
**Data Center** resource filter applied.

The UI at `https://localhost:9443/cables_validation/#/default-circuits-view/_`
is a browser app. This repo talks to the collector API under `/cablevalidation`.

Why it walks data halls
-----------------------

On the WC TX 16K collector there are ~353k circuits. A single

`GET /cablevalidation/report/circuits?context=dc`

times out. The script applies the same DC scope by requesting each data hall
(`context=dh&items=<hall>`) and de-duplicating `circuit_id`.

Docs: [Rest APIs 2.0.1](https://networking-docs.nvidia.com/cablevalidationtool/2.0.1/rest-apis),
[Reports APIs](https://networking-docs.nvidia.com/cablevalidationtool/2.0.1/reports-apis),
[Resource Filter](https://networking-docs.nvidia.com/cablevalidationtool/2.0.1/resource-filter).

Prerequisites
-------------

Keep the Teleport tunnel up:

```bash
tsh ssh -L 9443:192.168.10.3:9443 firstname.lastname.ct@mobilekit-p-phy-device100
```

Copy env and set the CVT UI password (do not commit `.env`):

```bash
cp .env.example .env
```

Usage
-----

Python 3.9+, no extra packages.

```bash
# DC circuit counts (safe)
python3 -m cvt_circuits stats

# Circuits View default: Data Center filter, Fail rows only
python3 -m cvt_circuits circuits --filter dc --status fail

# Every circuit in the data center (large; skip the JSON array)
python3 -m cvt_circuits circuits --filter dc --status all --skip-json --timeout 180
```

Outputs land in `out/`:

- `circuits.jsonl` — one JSON object per circuit
- `circuits.csv` — flattened A/Z columns for tracker paste
- `circuits.json` — JSON array (omit with `--skip-json` on full dumps)

Auth
----

Session cookie login, same as the UI:

`POST /cablevalidation/login` with `httpd_username` / `httpd_password`.

Credentials come from `CVT_USERNAME` / `CVT_PASSWORD` or `--username` / `--password`.
The collector uses a self-signed cert; `--insecure` is on by default.
