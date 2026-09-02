cvt-validation-tool
===================

Exports NVIDIA Cable Validation Tool **Circuits View** rows (Data Center
filter) to CSV.

Setup
-----

1. Clone the repo and go into that folder:

```bash
git clone git@github.com:nscaledev/cvt-validation-tool.git
cd cvt-validation-tool
```

2. Keep the Teleport tunnel up so `https://localhost:9443` reaches CVT:

```bash
tsh ssh -L 9443:192.168.10.3:9443 firstname.lastname.ct@mobilekit-p-phy-device100
```

3. Create a local `.env` **in this same folder** (next to `README.md`).
   Copy the example file, then put your CVT UI password in it:

```bash
cp .env.example .env
```

Edit `.env`:

```
CVT_URL=https://localhost:9443
CVT_USERNAME=nscale
CVT_PASSWORD=your-cvt-ui-password
CVT_INSECURE=true
```

`.env` is gitignored. Do not commit it.

Run
---

From the repo folder (`cvt-validation-tool`):

```bash
python3 -m cvt_circuits circuits
```

That pulls Fail + ethernet rows and writes:

`out/circuits-fail-ethernet-dc.csv`

If you do not want a `.env` file, pass credentials on the command line.
Flags go **before** `circuits`:

```bash
python3 -m cvt_circuits --username nscale --password 'your-cvt-ui-password' circuits
```

Other commands
--------------

```bash
python3 -m cvt_circuits stats
python3 -m cvt_circuits circuits --help
```

Notes
-----

On the WC TX 16K collector a single `context=dc` circuits request times
out (~353k circuits). This script walks each data hall
(`context=dh&items=<hall>`) and de-duplicates `circuit_id`.

Docs: [Rest APIs 2.0.1](https://networking-docs.nvidia.com/cablevalidationtool/2.0.1/rest-apis),
[Reports APIs](https://networking-docs.nvidia.com/cablevalidationtool/2.0.1/reports-apis),
[Resource Filter](https://networking-docs.nvidia.com/cablevalidationtool/2.0.1/resource-filter).
