"""Run a bounded real-LibreOffice witness matrix from a closed manifest."""

from __future__ import annotations

import argparse
import json
import re
import stat
from pathlib import Path, PurePosixPath
from typing import Any

from office_application_witness import ApplicationWitness

_ID = re.compile(r"[A-Za-z][A-Za-z0-9_.-]{0,79}\Z")
_GIT_ID = re.compile(r"[0-9a-f]{40}\Z")
_ARTIFACTS = {"docx": ".docx", "xlsx": ".xlsx", "pptx": ".pptx"}
_MAX_MANIFEST_BYTES = 1024 * 1024


class MatrixRefusal(ValueError):
    pass


def _load_manifest(path: Path) -> list[dict[str, str]]:
    info = path.lstat()
    if stat.S_ISLNK(info.st_mode) or not stat.S_ISREG(info.st_mode) or not 0 < info.st_size <= _MAX_MANIFEST_BYTES:
        raise MatrixRefusal
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict) or set(payload) != {"schema_version", "fixtures"}:
        raise MatrixRefusal
    fixtures = payload.get("fixtures")
    if payload.get("schema_version") != 1 or not isinstance(fixtures, list) or not 1 <= len(fixtures) <= 100:
        raise MatrixRefusal
    base = path.parent.resolve(strict=True)
    seen: set[str] = set()
    result: list[dict[str, str]] = []
    for item in fixtures:
        if not isinstance(item, dict) or set(item) != {"id", "artifact_type", "path"}:
            raise MatrixRefusal
        identifier = item.get("id")
        artifact_type = item.get("artifact_type")
        relative = item.get("path")
        if (
            not isinstance(identifier, str)
            or not _ID.fullmatch(identifier)
            or identifier in seen
            or artifact_type not in _ARTIFACTS
            or not isinstance(relative, str)
        ):
            raise MatrixRefusal
        pure = PurePosixPath(relative)
        if pure.is_absolute() or not pure.parts or ".." in pure.parts or "\\" in relative:
            raise MatrixRefusal
        source = (base / Path(*pure.parts)).resolve(strict=True)
        if source.parent != base and base not in source.parents:
            raise MatrixRefusal
        source_info = source.lstat()
        if stat.S_ISLNK(source_info.st_mode) or not stat.S_ISREG(source_info.st_mode) or source.suffix.lower() != _ARTIFACTS[artifact_type]:
            raise MatrixRefusal
        seen.add(identifier)
        result.append({"id": identifier, "artifact_type": artifact_type, "path": str(source)})
    return result


def run_matrix(args: argparse.Namespace) -> dict[str, Any]:
    fixtures = _load_manifest(Path(args.manifest))
    for value in (args.source_commit, args.source_tree):
        if value != "not_observed" and not _GIT_ID.fullmatch(value):
            raise MatrixRefusal
    identity = {
        "application_version": args.runtime_version,
        "image_digest": args.runtime_image_digest,
    }
    witness = ApplicationWitness(
        Path(args.workdir),
        executable=Path(args.executable),
        runtime_identity=identity,
    )
    reports = []
    all_ok = True
    for fixture in fixtures:
        report = witness.observe(fixture["path"], fixture["artifact_type"], timeout_seconds=args.timeout_seconds)
        reports.append({"id": fixture["id"], "artifact_type": fixture["artifact_type"], "report": report})
        all_ok = all_ok and report.get("status") == "ok"
    return {
        "schema_version": 1,
        "status": "ok" if all_ok else "refused",
        "reason": "fixture_failure" if not all_ok else None,
        "source_commit": args.source_commit,
        "source_tree": args.source_tree,
        "runtime_identity": identity,
        "fixtures": reports,
        "microsoft_office_equivalence_claimed": False,
        "pixel_level_fidelity_claimed": False,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--manifest", required=True)
    parser.add_argument("--workdir", required=True)
    parser.add_argument("--executable", required=True)
    parser.add_argument("--runtime-version", required=True)
    parser.add_argument("--runtime-image-digest", required=True)
    parser.add_argument("--source-commit", default="not_observed")
    parser.add_argument("--source-tree", default="not_observed")
    parser.add_argument("--timeout-seconds", type=float, default=60.0)
    args = parser.parse_args()
    try:
        report = run_matrix(args)
    except (MatrixRefusal, OSError, ValueError, TypeError, KeyError, json.JSONDecodeError):
        report = {"schema_version": 1, "status": "refused", "reason": "invalid_manifest"}
    print(json.dumps(report, ensure_ascii=False, separators=(",", ":")))
    return 0 if report["status"] == "ok" else 2


if __name__ == "__main__":
    raise SystemExit(main())
