#!/usr/bin/env python3
"""Run all domain suites without relying on shell interpolation."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SUITES = (
    ("docx", ROOT / "packages" / "docx", ["-m", "unittest", "discover", "-s", "tests", "-p", "test_*.py", "-v"]),
    ("xlsx", ROOT / "packages" / "xlsx", ["-m", "unittest", "discover", "-s", "tests", "-p", "test_*.py", "-v"]),
    ("pptx-editor", ROOT / "packages" / "pptx-editor", ["-m", "unittest", "discover", "-s", "tests", "-p", "test_*.py", "-v"]),
    ("pptx-composer", ROOT / "packages" / "pptx-composer", ["-m", "unittest", "discover", "-s", "tests", "-t", ".", "-v"]),
    ("application-witness", ROOT / "packages" / "application-witness", ["-m", "unittest", "discover", "-s", "tests", "-p", "test_*.py", "-v"]),
    ("repository-integration", ROOT, ["-m", "unittest", "discover", "-s", "tests", "-p", "test_*.py", "-v"]),
)


def main() -> int:
    failures: list[str] = []
    for name, cwd, args in SUITES:
        print(f"\n=== {name} ===", flush=True)
        completed = subprocess.run([sys.executable, *args], cwd=cwd, check=False)
        if completed.returncode:
            failures.append(name)
    if failures:
        print(f"FAILED: {', '.join(failures)}", file=sys.stderr)
        return 1
    print("\nALL DOMAIN SUITES PASSED", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
