#!/usr/bin/env python3
"""verify_reference_code - execute every runnable example in reference/*.md.

Documentation drifts. A recipe that was true for the version installed when it
was written silently becomes wrong after an upgrade, and the failure surfaces as
a confusing error inside a user's document task. This extracts the Python blocks
from the reference guides, runs the ones that are complete examples, and reports
which blocks no longer work.

    ~/.dsh/office-docs/venv/bin/python scripts/verify_reference_code.py

A block counts as a complete example when it both imports something and calls
`.save(...)`. Fragments (a table alone, a chart alone) are listed but not run,
because they depend on a document created by an earlier block.

Exit code 0 means every complete example ran without error.
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
import tempfile
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
REFERENCE_DIR = SCRIPT_DIR.parent / "reference"
FENCE = re.compile(r"^```python\s*$(.*?)^```\s*$", re.MULTILINE | re.DOTALL)

# Blocks that are complete examples but need a file on disk to operate on.
SETUP = {
    "docx.md": "",
    "xlsx.md": "",
    "pptx.md": "",
}


def iter_blocks(md_path: Path):
    text = md_path.read_text(encoding="utf-8")
    return list(FENCE.finditer(text))


def is_complete(code: str) -> bool:
    return "import" in code and ".save(" in code


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--keep", action="store_true", help="keep generated artifacts")
    args = ap.parse_args()

    failures: list[str] = []
    ran = 0
    skipped = 0

    for md_path in sorted(REFERENCE_DIR.glob("*.md")):
        blocks = iter_blocks(md_path)
        print(f"\n{md_path.name}: {len(blocks)} python block(s)")
        for index, match in enumerate(blocks, start=1):
            code = match.group(1)
            label = f"{md_path.name} block {index}"
            if not is_complete(code):
                skipped += 1
                first_line = next((l.strip() for l in code.splitlines() if l.strip()), "")
                print(f"  skip  {label}: fragment ({first_line[:60]})")
                continue

            tmp = Path(tempfile.mkdtemp(prefix="office-docs-refcheck-"))
            script = tmp / "example.py"
            script.write_text(code, encoding="utf-8")
            proc = subprocess.run(
                [sys.executable, str(script)],
                capture_output=True,
                text=True,
                cwd=str(tmp),
                timeout=180,
            )
            produced = sorted(p.name for p in tmp.iterdir() if p.suffix in {".docx", ".xlsx", ".pptx"})
            if proc.returncode != 0:
                failures.append(label)
                print(f"  FAIL  {label}")
                print("        " + (proc.stderr or proc.stdout).strip().replace("\n", "\n        ")[:1200])
            else:
                ran += 1
                print(f"  ok    {label}: produced {', '.join(produced) or 'nothing'}")
                if not produced:
                    failures.append(f"{label} (ran but saved no file)")
            if not args.keep:
                import shutil

                shutil.rmtree(tmp, ignore_errors=True)
            else:
                print(f"        artifacts: {tmp}")

    print(f"\n{ran} example(s) ran, {skipped} fragment(s) skipped, {len(failures)} failure(s)")
    for name in failures:
        print(f"  FAILED: {name}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
