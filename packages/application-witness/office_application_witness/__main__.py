from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .api import ApplicationWitness, _refusal


def main() -> int:
    try:
        parser = argparse.ArgumentParser(prog="office-witness")
        parser.add_argument("--workdir", required=True)
        parser.add_argument("--executable", required=True)
        args = parser.parse_args()
        payload = json.load(sys.stdin)
        if not isinstance(payload, dict) or set(payload) - {"source", "artifact_type", "timeout_seconds"}:
            raise ValueError
        witness = ApplicationWitness(args.workdir, executable=args.executable)
        result = witness.observe(
            Path(payload["source"]),
            payload["artifact_type"],
            timeout_seconds=payload.get("timeout_seconds", 60),
        )
    except (KeyError, OSError, TypeError, ValueError, json.JSONDecodeError):
        result = _refusal("validation_failure", "invalid witness request")
    print(json.dumps(result, ensure_ascii=False, separators=(",", ":")))
    return 0 if result.get("status") == "ok" else 2


if __name__ == "__main__":
    raise SystemExit(main())
