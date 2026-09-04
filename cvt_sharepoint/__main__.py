from pathlib import Path
import sys

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from venv_bootstrap import ensure_venv

ensure_venv("cvt_sharepoint")

from cvt_sharepoint.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
