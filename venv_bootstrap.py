"""Create .venv and install requirements.txt before the real script runs.

Python packages (msal, openpyxl, …) are installed with pip into .venv.
Homebrew is not used for those libraries. The venv directory name is always
``.venv`` in this repo root. Later runs reuse it and skip pip when
requirements.txt has not changed.
"""

from __future__ import annotations

import hashlib
import os
import subprocess
import sys
import venv
from pathlib import Path
from typing import Optional

VENV_DIRNAME = ".venv"
IN_VENV_FLAG = "CVT_IN_VENV"
STAMP_NAME = ".requirements.sha256"
MIN_PY = (3, 9)


def project_root() -> Path:
    return Path(__file__).resolve().parent


def venv_dir(root: Optional[Path] = None) -> Path:
    return (root or project_root()) / VENV_DIRNAME


def venv_python(root: Optional[Path] = None) -> Path:
    venv_path = venv_dir(root)
    if os.name == "nt":
        return venv_path / "Scripts" / "python.exe"
    return venv_path / "bin" / "python"


def _requirements_hash(root: Path) -> str:
    return hashlib.sha256((root / "requirements.txt").read_bytes()).hexdigest()


def _stamp_path(root: Path) -> Path:
    return venv_dir(root) / STAMP_NAME


def _running_in_venv(root: Path) -> bool:
    if os.environ.get(IN_VENV_FLAG) == "1":
        return True
    # On macOS .venv/bin/python is often a symlink to the system interpreter,
    # so compare prefixes instead of resolved executables.
    try:
        return Path(sys.prefix).resolve() == venv_dir(root).resolve()
    except OSError:
        return False


def _create_venv(root: Path) -> None:
    path = venv_dir(root)
    print(f"[cvt] Creating virtualenv {path} …", file=sys.stderr)
    builder = venv.EnvBuilder(with_pip=True, clear=False)
    try:
        builder.create(path)
    except Exception:
        # Some distros ship Python without ensurepip; retry after bootstrap.
        subprocess.check_call([sys.executable, "-m", "ensurepip", "--upgrade"])
        builder.create(path)
    print(f"[cvt] Virtualenv ready: {path}", file=sys.stderr)


def _pip(root: Path, *args: str) -> None:
    py = venv_python(root)
    subprocess.check_call(
        [str(py), "-m", "pip", "--disable-pip-version-check", *args],
        stdout=sys.stderr,
    )


def _install_requirements(root: Path) -> None:
    req = root / "requirements.txt"
    print(f"[cvt] Installing packages from requirements.txt into {venv_dir(root)} …", file=sys.stderr)
    # Keep pip itself usable on fresh venvs.
    _pip(root, "install", "--upgrade", "pip", "setuptools", "wheel")
    _pip(root, "install", "-r", str(req))
    _stamp_path(root).write_text(_requirements_hash(root) + "\n", encoding="utf-8")
    print("[cvt] Dependencies installed (reuse .venv on next run).", file=sys.stderr)


def _deps_are_current(root: Path) -> bool:
    stamp = _stamp_path(root)
    if not venv_python(root).is_file() or not stamp.is_file():
        return False
    try:
        return stamp.read_text(encoding="utf-8").strip() == _requirements_hash(root)
    except OSError:
        return False


def _reexec(root: Path, module: str) -> None:
    py = str(venv_python(root))
    env = os.environ.copy()
    env[IN_VENV_FLAG] = "1"
    existing = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = os.pathsep.join(p for p in (str(root), existing) if p)
    os.execve(py, [py, "-m", module, *sys.argv[1:]], env)


def prepare_venv(root: Optional[Path] = None) -> Path:
    """Create/update ``.venv`` and install requirements. Return venv python path.

    Idempotent: later runs reuse the same ``.venv`` and skip pip when
    ``requirements.txt`` is unchanged. Does not re-exec the process.
    """
    if sys.version_info < MIN_PY:
        needed = ".".join(str(part) for part in MIN_PY)
        raise SystemExit(
            f"Python {needed}+ is required (found {sys.version.split()[0]}).\n"
            "On macOS install it with:  brew install python\n"
            "or:  xcode-select --install"
        )

    root = root or project_root()
    req = root / "requirements.txt"
    if not req.is_file():
        raise SystemExit(f"missing {req}")

    if not venv_python(root).is_file():
        try:
            _create_venv(root)
        except Exception as exc:
            raise SystemExit(
                f"Could not create {venv_dir(root)}: {exc}\n"
                "On Debian/Ubuntu you may need: sudo apt install python3-venv python3-pip"
            ) from exc

    if not _deps_are_current(root):
        try:
            _install_requirements(root)
        except subprocess.CalledProcessError as exc:
            raise SystemExit(
                f"pip install failed (exit {exc.returncode}). "
                "Check network access to PyPI, or set HTTPS_PROXY if you are behind a proxy."
            ) from exc

    py = venv_python(root)
    if not py.is_file():
        raise SystemExit(f"venv python missing after setup: {py}")
    return py


def ensure_venv(module: str) -> None:
    """Prepare ``.venv`` then re-exec into it when not already running there."""
    root = project_root()
    prepare_venv(root)
    if not _running_in_venv(root):
        _reexec(root, module)
