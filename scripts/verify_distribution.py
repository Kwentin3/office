#!/usr/bin/env python3
"""Fail closed when a built wheel or sdist omits required reusable assets."""

from __future__ import annotations

import configparser
import sys
import tarfile
import zipfile
from pathlib import Path, PurePosixPath

WHEEL_REQUIRED = {
    "office_artifact_tool/__init__.py",
    "office_artifact_tool/__main__.py",
    "office_artifact_tool/resources/AGENT_SKILL.md",
    "office_artifact_tool/resources/create.schema.json",
    "office_artifact_tool/resources/plan.schema.json",
    "xlsx_artifact_tool/__init__.py",
    "xlsx_artifact_tool/__main__.py",
    "xlsx_artifact_tool/resources/AGENT_SKILL.md",
    "xlsx_artifact_tool/resources/CONTRACT.md",
    "xlsx_artifact_tool/resources/create.schema.json",
    "xlsx_artifact_tool/resources/plan.schema.json",
    "pptx_artifact_tool/__init__.py",
    "pptx_artifact_tool/__main__.py",
    "pptx_artifact_tool/resources/AGENT_SKILL.md",
    "pptx_artifact_tool/resources/CONTRACT.md",
    "pptx_ai_composer/__init__.py",
    "pptx_ai_composer/__main__.py",
    "pptx_ai_composer/assets/NotoSans-Regular.ttf",
    "pptx_ai_composer/assets/OFL.txt",
    "pptx_ai_composer/resources/DOMAIN_CONTRACTS.md",
    "pptx_ai_composer/resources/MANAGED_LIBRARY.md",
}
WHEEL_REQUIRED_SUFFIXES = {
    ".dist-info/entry_points.txt",
    ".dist-info/licenses/LICENSE",
    ".dist-info/licenses/NOTICE",
}
REQUIRED_CONSOLE_SCRIPTS = {
    "office-docx": "office_artifact_tool.__main__:main",
    "office-xlsx": "xlsx_artifact_tool.__main__:main",
    "office-pptx-edit": "pptx_artifact_tool.__main__:main",
    "office-pptx-compose": "pptx_ai_composer.__main__:main",
}
SDIST_SUFFIXES = {
    "README.md",
    "LICENSE",
    "NOTICE",
    "pyproject.toml",
    "docs/openwebui-integration.md",
    "examples/openwebui_backend/office_service.py",
    "packages/docx/tests/test_mvp_api.py",
    "packages/xlsx/tests/test_safety.py",
    "packages/pptx-editor/tests/test_safety.py",
    "packages/pptx-composer/tests/test_renderer.py",
    "packages/pptx-composer/pptx_ai_composer/assets/NotoSans-Regular.ttf",
    "packages/pptx-composer/pptx_ai_composer/assets/OFL.txt",
    "tests/test_openwebui_adapter.py",
    "tests/test_distribution_verifier.py",
    "scripts/test_all.py",
    "scripts/verify_distribution.py",
}
FORBIDDEN_PARTS = {"__pycache__", ".git", ".venv", "build", "dist"}
FORBIDDEN_SUFFIXES = {".pyc", ".pyo", ".log"}


def unsafe(names: list[str]) -> list[str]:
    bad: list[str] = []
    for name in names:
        path = PurePosixPath(name)
        if name.startswith("/") or ".." in path.parts:
            bad.append(name)
        elif any(part in FORBIDDEN_PARTS for part in path.parts):
            bad.append(name)
        elif path.suffix in FORBIDDEN_SUFFIXES:
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


def main() -> int:
    dist = Path(sys.argv[1] if len(sys.argv) > 1 else "dist")
    wheels = sorted(dist.glob("*.whl"))
    sdists = sorted(dist.glob("*.tar.gz"))
    if len(wheels) != 1 or len(sdists) != 1:
        raise SystemExit("expected exactly one wheel and one sdist")
    with zipfile.ZipFile(wheels[0]) as archive:
        wheel_names = archive.namelist()
        entry_point_names = [name for name in wheel_names if name.endswith(".dist-info/entry_points.txt")]
        entry_points_payload = archive.read(entry_point_names[0]) if len(entry_point_names) == 1 else b""
    missing_wheel = sorted(WHEEL_REQUIRED - set(wheel_names))
    missing_wheel_suffixes = sorted(
        suffix for suffix in WHEEL_REQUIRED_SUFFIXES if not any(name.endswith(suffix) for name in wheel_names)
    )
    with tarfile.open(sdists[0], "r:gz") as archive:
        members = archive.getmembers()
        sdist_names = [member.name for member in members]
        non_files = [member.name for member in members if not (member.isfile() or member.isdir())]
    missing_sdist = sorted(
        suffix for suffix in SDIST_SUFFIXES if not any(name.endswith("/" + suffix) for name in sdist_names)
    )
    console_scripts = parse_console_scripts(entry_points_payload)
    invalid_console_scripts = console_scripts != REQUIRED_CONSOLE_SCRIPTS
    bad = unsafe(wheel_names + sdist_names)
    if missing_wheel or missing_wheel_suffixes or missing_sdist or invalid_console_scripts or bad or non_files:
        raise SystemExit(
            f"distribution verification failed: missing_wheel={missing_wheel}, "
            f"missing_wheel_suffixes={missing_wheel_suffixes}, missing_sdist={missing_sdist}, "
            f"invalid_console_scripts={invalid_console_scripts}, unsafe={bad}, non_files={non_files}"
        )
    print(
        f"distribution verified: wheel_members={len(wheel_names)}, "
        f"sdist_members={len(sdist_names)}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
