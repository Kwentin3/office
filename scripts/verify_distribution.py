#!/usr/bin/env python3
"""Fail closed when a built wheel or sdist diverges from the release inventory."""

from __future__ import annotations

import base64
import configparser
import csv
import email.parser
import hashlib
import io
import sys
import tarfile
import zipfile
from pathlib import Path, PurePosixPath

WHEEL_RUNTIME = {
    "office_artifact_tool/__init__.py",
    "office_artifact_tool/__main__.py",
    "office_artifact_tool/api.py",
    "office_artifact_tool/core/__init__.py",
    "office_artifact_tool/core/contracts.py",
    "office_artifact_tool/core/errors.py",
    "office_artifact_tool/core/hashes.py",
    "office_artifact_tool/core/plans.py",
    "office_artifact_tool/core/semantic.py",
    "office_artifact_tool/core/transaction.py",
    "office_artifact_tool/core/validation.py",
    "office_artifact_tool/docx/__init__.py",
    "office_artifact_tool/docx/inspect.py",
    "office_artifact_tool/docx/inventory.py",
    "office_artifact_tool/docx/mutation.py",
    "office_artifact_tool/docx/preview.py",
    "office_artifact_tool/docx/presentation.py",
    "office_artifact_tool/docx/renderer.py",
    "office_artifact_tool/docx/review_contract.py",
    "office_artifact_tool/docx/template.py",
    "office_artifact_tool/resources/AGENT_SKILL.md",
    "office_artifact_tool/resources/create.schema.json",
    "office_artifact_tool/resources/inventory.schema.json",
    "office_artifact_tool/resources/plan.schema.json",
    "xlsx_artifact_tool/__init__.py",
    "xlsx_artifact_tool/__main__.py",
    "xlsx_artifact_tool/api.py",
    "xlsx_artifact_tool/inventory.py",
    "xlsx_artifact_tool/preservation.py",
    "xlsx_artifact_tool/preview.py",
    "xlsx_artifact_tool/resources/AGENT_SKILL.md",
    "xlsx_artifact_tool/resources/CONTRACT.md",
    "xlsx_artifact_tool/resources/create.schema.json",
    "xlsx_artifact_tool/resources/inventory.schema.json",
    "xlsx_artifact_tool/resources/plan.schema.json",
    "xlsx_artifact_tool/review_contract.py",
    "xlsx_artifact_tool/template.py",
    "pptx_artifact_tool/__init__.py",
    "pptx_artifact_tool/__main__.py",
    "pptx_artifact_tool/api.py",
    "pptx_artifact_tool/inventory.py",
    "pptx_artifact_tool/resources/AGENT_SKILL.md",
    "pptx_artifact_tool/resources/CONTRACT.md",
    "pptx_artifact_tool/resources/inventory.schema.json",
    "pptx_artifact_tool/template.py",
    "pptx_ai_composer/__init__.py",
    "pptx_ai_composer/__main__.py",
    "pptx_ai_composer/assets/NotoSans-Regular.ttf",
    "pptx_ai_composer/assets/OFL.txt",
    "pptx_ai_composer/asset_admission.py",
    "pptx_ai_composer/compiler.py",
    "pptx_ai_composer/contracts.py",
    "pptx_ai_composer/library.py",
    "pptx_ai_composer/preview.py",
    "pptx_ai_composer/renderer.py",
    "pptx_ai_composer/review_contract.py",
    "pptx_ai_composer/resources/DOMAIN_CONTRACTS.md",
    "pptx_ai_composer/resources/MANAGED_LIBRARY.md",
    "pptx_ai_composer/scene_contract.py",
    "pptx_ai_composer/validator.py",
    "office_application_witness/__init__.py",
    "office_application_witness/__main__.py",
    "office_application_witness/api.py",
    "office_application_witness/contracts.py",
}
DIST_INFO = "kwentin_office-0.4.0.dist-info"
WHEEL_FILENAME = "kwentin_office-0.4.0-py3-none-any.whl"
SDIST_FILENAME = "kwentin_office-0.4.0.tar.gz"
WHEEL_METADATA = {
    f"{DIST_INFO}/METADATA",
    f"{DIST_INFO}/WHEEL",
    f"{DIST_INFO}/entry_points.txt",
    f"{DIST_INFO}/top_level.txt",
    f"{DIST_INFO}/RECORD",
    f"{DIST_INFO}/licenses/LICENSE",
    f"{DIST_INFO}/licenses/NOTICE",
}
WHEEL_REQUIRED = WHEEL_RUNTIME | WHEEL_METADATA
REQUIRED_CONSOLE_SCRIPTS = {
    "office-docx": "office_artifact_tool.__main__:main",
    "office-xlsx": "xlsx_artifact_tool.__main__:main",
    "office-pptx-edit": "pptx_artifact_tool.__main__:main",
    "office-pptx-compose": "pptx_ai_composer.__main__:main",
    "office-witness": "office_application_witness.__main__:main",
}
REQUIRED_TOP_LEVEL = {
    "office_application_witness",
    "office_artifact_tool",
    "pptx_ai_composer",
    "pptx_artifact_tool",
    "xlsx_artifact_tool",
}
FORBIDDEN_PARTS = {"__pycache__", ".git", ".venv", "build", "dist"}
FORBIDDEN_SUFFIXES = {".pyc", ".pyo", ".log"}
ROOT_RELEASE_FILES = {
    "CHANGELOG.md",
    "CONTRIBUTING.md",
    "LICENSE",
    "MANIFEST.in",
    "NOTICE",
    "README.md",
    "SECURITY.md",
    "pyproject.toml",
}
SDIST_GENERATED = {
    "PKG-INFO",
    "setup.cfg",
    "kwentin_office.egg-info/PKG-INFO",
    "kwentin_office.egg-info/SOURCES.txt",
    "kwentin_office.egg-info/dependency_links.txt",
    "kwentin_office.egg-info/entry_points.txt",
    "kwentin_office.egg-info/requires.txt",
    "kwentin_office.egg-info/top_level.txt",
}


def release_source_inventory(root: Path) -> set[str]:
    inventory = {name for name in ROOT_RELEASE_FILES if (root / name).is_file()}
    rules = {
        "docs": {".md"},
        "examples": {".py", ".json"},
        "packages": {".py", ".md", ".json", ".ttf", ".txt"},
        "tests": {".py"},
        "scripts": {".py"},
    }
    for directory, suffixes in rules.items():
        base = root / directory
        if not base.is_dir():
            continue
        for path in base.rglob("*"):
            if (
                path.is_file()
                and path.suffix in suffixes
                and not any(part in FORBIDDEN_PARTS for part in path.relative_to(root).parts)
            ):
                inventory.add(path.relative_to(root).as_posix())
    return inventory


SDIST_REQUIRED = release_source_inventory(Path(__file__).resolve().parents[1]) | SDIST_GENERATED


def unsafe(names: list[str]) -> list[str]:
    bad: list[str] = []
    for name in names:
        path = PurePosixPath(name)
        if (
            not name
            or name.startswith("/")
            or "\\" in name
            or ".." in path.parts
            or any(part in FORBIDDEN_PARTS for part in path.parts)
            or path.suffix in FORBIDDEN_SUFFIXES
        ):
            bad.append(name)
    return bad


def parse_console_scripts(payload: bytes) -> dict[str, str] | None:
    parser = configparser.ConfigParser(interpolation=None)
    try:
        parser.read_string(payload.decode("utf-8"))
    except (UnicodeDecodeError, configparser.Error):
        return None
    if not parser.has_section("console_scripts"):
        return None
    return dict(parser.items("console_scripts"))


def metadata_is_valid(payload: bytes) -> bool:
    try:
        metadata = email.parser.BytesParser().parsebytes(payload)
    except (TypeError, ValueError):
        return False
    return not metadata.defects and all(
        metadata.get_all(name) == [value]
        for name, value in (
            ("Name", "kwentin-office"),
            ("Version", "0.4.0"),
            ("Requires-Python", ">=3.11"),
        )
    )


def top_level_is_valid(payload: bytes) -> bool:
    try:
        names = {line.strip() for line in payload.decode("utf-8").splitlines() if line.strip()}
    except UnicodeDecodeError:
        return False
    return names == REQUIRED_TOP_LEVEL


def wheel_metadata_is_valid(payload: bytes) -> bool:
    try:
        metadata = email.parser.BytesParser().parsebytes(payload)
    except (TypeError, ValueError):
        return False
    return not metadata.defects and all(
        metadata.get_all(name) == [value]
        for name, value in (
            ("Wheel-Version", "1.0"),
            ("Root-Is-Purelib", "true"),
            ("Tag", "py3-none-any"),
        )
    )


def record_is_valid(payload: bytes, archive_payloads: dict[str, bytes]) -> bool:
    try:
        rows = list(csv.reader(io.StringIO(payload.decode("utf-8"), newline="")))
    except (UnicodeDecodeError, csv.Error):
        return False
    if any(len(row) != 3 for row in rows):
        return False
    entries = {row[0]: (row[1], row[2]) for row in rows}
    if len(entries) != len(rows) or set(entries) != set(archive_payloads):
        return False
    record_name = f"{DIST_INFO}/RECORD"
    for name, data in archive_payloads.items():
        digest, size = entries[name]
        if name == record_name:
            if digest or size:
                return False
            continue
        expected = base64.urlsafe_b64encode(hashlib.sha256(data).digest()).rstrip(b"=").decode("ascii")
        if digest != f"sha256={expected}" or size != str(len(data)):
            return False
    return True


def strip_sdist_root(names: list[str]) -> tuple[str | None, set[str]]:
    roots = {PurePosixPath(name).parts[0] for name in names if PurePosixPath(name).parts}
    if len(roots) != 1:
        return None, set()
    root = roots.pop()
    stripped = {
        PurePosixPath(*PurePosixPath(name).parts[1:]).as_posix() for name in names if len(PurePosixPath(name).parts) > 1
    }
    return root, stripped


def main() -> int:
    dist = Path(sys.argv[1] if len(sys.argv) > 1 else "dist")
    wheels = sorted(dist.glob("*.whl"))
    sdists = sorted(dist.glob("*.tar.gz"))
    if len(wheels) != 1 or len(sdists) != 1:
        raise SystemExit("expected exactly one wheel and one sdist")
    if wheels[0].name != WHEEL_FILENAME or sdists[0].name != SDIST_FILENAME:
        raise SystemExit("artifact filename does not match project name, version, and wheel tag")
    with zipfile.ZipFile(wheels[0]) as archive:
        wheel_names = archive.namelist()
        wheel_archive_payloads = {name: archive.read(name) for name in wheel_names}
        wheel_payload = {name: archive.read(name) for name in WHEEL_METADATA if name in wheel_names}
    with tarfile.open(sdists[0], "r:gz") as archive:
        members = archive.getmembers()
        sdist_names = [member.name for member in members]
        sdist_file_names = [member.name for member in members if member.isfile()]
        non_files = [member.name for member in members if not (member.isfile() or member.isdir())]
    sdist_root, _ = strip_sdist_root(sdist_names)
    _, sdist_stripped = strip_sdist_root(sdist_file_names)
    wheel_set = set(wheel_names)
    missing_wheel = sorted(WHEEL_REQUIRED - wheel_set)
    unexpected_wheel = sorted(wheel_set - WHEEL_REQUIRED)
    missing_sdist = sorted(SDIST_REQUIRED - sdist_stripped)
    unexpected_sdist = sorted(sdist_stripped - SDIST_REQUIRED)
    duplicate_names = sorted(
        {name for name in wheel_names if wheel_names.count(name) > 1}
        | {name for name in sdist_names if sdist_names.count(name) > 1}
    )
    console_scripts = parse_console_scripts(wheel_payload.get(f"{DIST_INFO}/entry_points.txt", b""))
    invalid_console_scripts = console_scripts != REQUIRED_CONSOLE_SCRIPTS
    invalid_metadata = not metadata_is_valid(wheel_payload.get(f"{DIST_INFO}/METADATA", b""))
    invalid_wheel_metadata = not wheel_metadata_is_valid(wheel_payload.get(f"{DIST_INFO}/WHEEL", b""))
    invalid_record = not record_is_valid(
        wheel_payload.get(f"{DIST_INFO}/RECORD", b""), wheel_archive_payloads
    )
    invalid_top_level = not top_level_is_valid(wheel_payload.get(f"{DIST_INFO}/top_level.txt", b""))
    bad = unsafe(wheel_names + sdist_names)
    if (
        missing_wheel
        or unexpected_wheel
        or missing_sdist
        or unexpected_sdist
        or sdist_root != "kwentin_office-0.4.0"
        or duplicate_names
        or invalid_console_scripts
        or invalid_metadata
        or invalid_wheel_metadata
        or invalid_record
        or invalid_top_level
        or bad
        or non_files
    ):
        raise SystemExit(
            "distribution verification failed: "
            f"missing_wheel={missing_wheel}, unexpected_wheel={unexpected_wheel}, "
            f"missing_sdist={missing_sdist}, unexpected_sdist={unexpected_sdist}, "
            f"invalid_sdist_root={sdist_root != 'kwentin_office-0.4.0'}, duplicates={duplicate_names}, "
            f"invalid_console_scripts={invalid_console_scripts}, invalid_metadata={invalid_metadata}, "
            f"invalid_wheel_metadata={invalid_wheel_metadata}, invalid_record={invalid_record}, "
            f"invalid_top_level={invalid_top_level}, unsafe={bad}, non_files={non_files}"
        )
    print(f"distribution verified: wheel_members={len(wheel_names)}, sdist_members={len(sdist_names)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
