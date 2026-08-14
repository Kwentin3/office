from __future__ import annotations

import subprocess
import sys
import tarfile
import tempfile
import unittest
import zipfile
from io import BytesIO
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
VERIFIER = ROOT / "scripts/verify_distribution.py"
DIST_INFO = "kwentin_office-0.1.0.dist-info"

WHEEL_MEMBERS = {
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
    f"{DIST_INFO}/entry_points.txt",
    f"{DIST_INFO}/licenses/LICENSE",
    f"{DIST_INFO}/licenses/NOTICE",
}

SDIST_MEMBERS = {
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


class DistributionVerifierTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.dist = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def write_archives(
        self,
        *,
        omit_wheel: str | None = None,
        omit_sdist: str | None = None,
        entry_points: bytes | None = None,
    ) -> None:
        wheel = self.dist / "kwentin_office-0.1.0-py3-none-any.whl"
        valid_entry_points = (
            b"[console_scripts]\n"
            b"office-docx = office_artifact_tool.__main__:main\n"
            b"office-xlsx = xlsx_artifact_tool.__main__:main\n"
            b"office-pptx-edit = pptx_artifact_tool.__main__:main\n"
            b"office-pptx-compose = pptx_ai_composer.__main__:main\n"
        )
        with zipfile.ZipFile(wheel, "w") as archive:
            for name in sorted(WHEEL_MEMBERS - ({omit_wheel} if omit_wheel else set())):
                if name.endswith(".dist-info/entry_points.txt"):
                    payload = entry_points if entry_points is not None else valid_entry_points
                else:
                    payload = b"x"
                archive.writestr(name, payload)
        sdist = self.dist / "kwentin_office-0.1.0.tar.gz"
        with tarfile.open(sdist, "w:gz") as archive:
            for suffix in sorted(SDIST_MEMBERS - ({omit_sdist} if omit_sdist else set())):
                payload = b"x"
                info = tarfile.TarInfo(f"kwentin_office-0.1.0/{suffix}")
                info.size = len(payload)
                archive.addfile(info, BytesIO(payload))

    def verify(self) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [sys.executable, str(VERIFIER), str(self.dist)],
            text=True,
            capture_output=True,
            check=False,
        )

    def test_complete_distribution_passes(self) -> None:
        self.write_archives()
        self.assertEqual(self.verify().returncode, 0)

    def test_each_required_wheel_member_is_fail_closed(self) -> None:
        for member in sorted(WHEEL_MEMBERS):
            with self.subTest(member=member):
                self.write_archives(omit_wheel=member)
                result = self.verify()
                self.assertNotEqual(result.returncode, 0, member)

    def test_each_required_sdist_member_is_fail_closed(self) -> None:
        for member in sorted(SDIST_MEMBERS):
            with self.subTest(member=member):
                self.write_archives(omit_sdist=member)
                result = self.verify()
                self.assertNotEqual(result.returncode, 0, member)

    def test_all_four_console_scripts_are_required(self) -> None:
        incomplete = (
            b"[console_scripts]\n"
            b"office-docx = office_artifact_tool.__main__:main\n"
        )
        self.write_archives(entry_points=incomplete)
        self.assertNotEqual(self.verify().returncode, 0)


if __name__ == "__main__":
    unittest.main()
