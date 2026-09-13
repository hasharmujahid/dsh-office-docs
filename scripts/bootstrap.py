#!/usr/bin/env python3
"""Create (or repair) the shared Python environment used by the office-docs skill.

Idempotent: running it when the environment is already complete only verifies
imports and exits 0. The environment lives OUTSIDE any workspace so that every
DSH session, in every working directory, can use the same one.

Usage:
    python3 scripts/bootstrap.py            # create/verify
    python3 scripts/bootstrap.py --check    # verify only, never install
    python3 scripts/bootstrap.py --force    # rebuild from scratch
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

# Where the environment lives. Override with OFFICE_DOCS_VENV.
DEFAULT_ROOT = Path(os.environ.get("OFFICE_DOCS_HOME", Path.home() / ".dsh" / "office-docs"))
DEFAULT_VENV = Path(os.environ.get("OFFICE_DOCS_VENV", DEFAULT_ROOT / "venv"))

# Pinned to major versions: these APIs are stable, and a floating install is
# what makes document pipelines fail six months after they were written.
REQUIREMENTS = [
    "python-docx>=1.1,<2",
    "openpyxl>=3.1,<4",
    "python-pptx>=1.0,<2",
    "lxml>=5,<7",
    "pillow>=10",
]

# Import name -> pip requirement, so verification can name the real culprit.
IMPORT_CHECKS = [
    ("docx", "python-docx"),
    ("openpyxl", "openpyxl"),
    ("pptx", "python-pptx"),
    ("lxml", "lxml"),
    ("PIL", "pillow"),
]

VERIFY_SNIPPET = r"""
import importlib, json, sys
report = {}
for mod in ["docx", "openpyxl", "pptx", "lxml", "PIL"]:
    try:
        m = importlib.import_module(mod)
        report[mod] = getattr(m, "__version__", "ok")
    except Exception as exc:  # noqa: BLE001 - report, never crash
        report[mod] = "MISSING: %s" % exc
print(json.dumps(report))
sys.exit(0 if all(not str(v).startswith("MISSING") for v in report.values()) else 1)
"""


def venv_python(venv: Path) -> Path:
    """Interpreter inside `venv`, on either platform."""
    if os.name == "nt":
        return venv / "Scripts" / "python.exe"
    return venv / "bin" / "python"


def run(cmd: list[str], **kwargs) -> subprocess.CompletedProcess:
    printable = " ".join(cmd)
    print(f"  $ {printable}", flush=True)
    return subprocess.run(cmd, text=True, **kwargs)


def bootstrap(venv: Path, check_only: bool, force: bool) -> int:
    if force and venv.exists():
        print(f"Removing {venv}")
        shutil.rmtree(venv)

    py = venv_python(venv)
    if not py.exists():
        if check_only:
            print(f"ERROR: no environment at {venv}", file=sys.stderr)
            print(f"Run: python3 {Path(__file__).resolve()}", file=sys.stderr)
            return 1
        print(f"Creating virtual environment at {venv}")
        venv.parent.mkdir(parents=True, exist_ok=True)
        # --without-pip is faster but we need pip next.
        rc = run([sys.executable, "-m", "venv", str(venv)])
        if rc.returncode != 0:
            print("ERROR: venv creation failed", file=sys.stderr)
            return rc.returncode
    else:
        print(f"Environment present: {venv}")

    # Verify before installing: the common case is "already fine".
    probe = subprocess.run([str(py), "-c", VERIFY_SNIPPET], text=True, capture_output=True)
    if probe.returncode == 0:
        print(f"All document libraries import cleanly: {probe.stdout.strip()}")
        return 0

    if check_only:
        print(f"ERROR: environment incomplete: {probe.stdout.strip()}", file=sys.stderr)
        return 1

    print("Installing/repairing packages (network required, one time):")
    rc = run([str(py), "-m", "pip", "install", "--upgrade", "pip", "--quiet"])
    if rc.returncode != 0:
        print("WARNING: pip self-upgrade failed; continuing", file=sys.stderr)
    rc = run([str(py), "-m", "pip", "install", "--quiet", *REQUIREMENTS])
    if rc.returncode != 0:
        print("ERROR: package install failed", file=sys.stderr)
        return rc.returncode

    probe = subprocess.run([str(py), "-c", VERIFY_SNIPPET], text=True, capture_output=True)
    print(f"Verify: {probe.stdout.strip()}")
    if probe.returncode != 0:
        print("ERROR: install completed but imports still fail", file=sys.stderr)
        return 1
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--check", action="store_true", help="verify only; never install")
    ap.add_argument("--force", action="store_true", help="delete and rebuild the environment")
    ap.add_argument("--venv", default=str(DEFAULT_VENV), help=f"environment path (default: {DEFAULT_VENV})")
    args = ap.parse_args()

    if shutil.which("soffice") is None and shutil.which("libreoffice") is None:
        print("WARNING: LibreOffice (soffice) not on PATH: rendering/PDF conversion will be unavailable.", file=sys.stderr)

    rc = bootstrap(Path(args.venv), args.check, args.force)
    if rc == 0:
        print(f"\nReady. Use it with:\n  {venv_python(Path(args.venv))} scripts/office.py --help")
    return rc


if __name__ == "__main__":
    raise SystemExit(main())
