from __future__ import annotations

import importlib.util
import subprocess
import sys
import tarfile
import tempfile
import unittest
import zipfile
from io import BytesIO
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
VERIFIER_PATH = ROOT / "scripts/verify_distribution.py"
SPEC = importlib.util.spec_from_file_location("distribution_verifier", VERIFIER_PATH)
assert SPEC is not None and SPEC.loader is not None
VERIFIER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(VERIFIER)

DIST_INFO = VERIFIER.DIST_INFO
WHEEL_MEMBERS = set(VERIFIER.WHEEL_REQUIRED)
SDIST_MEMBERS = set(VERIFIER.SDIST_REQUIRED)
VALID_ENTRY_POINTS = (
    b"[console_scripts]\n"
    b"office-docx = office_artifact_tool.__main__:main\n"
    b"office-xlsx = xlsx_artifact_tool.__main__:main\n"
    b"office-pptx-edit = pptx_artifact_tool.__main__:main\n"
    b"office-pptx-compose = pptx_ai_composer.__main__:main\n"
    b"office-witness = office_application_witness.__main__:main\n"
)
VALID_METADATA = b"Metadata-Version: 2.4\nName: kwentin-office\nVersion: 0.2.0\nRequires-Python: >=3.11\n"
VALID_TOP_LEVEL = b"\n".join(name.encode() for name in sorted(VERIFIER.REQUIRED_TOP_LEVEL)) + b"\n"


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
        entry_points: bytes = VALID_ENTRY_POINTS,
        metadata: bytes = VALID_METADATA,
        top_level: bytes = VALID_TOP_LEVEL,
        extra_wheel: str | None = None,
        extra_sdist: str | None = None,
        duplicate_wheel: str | None = None,
    ) -> None:
        wheel = self.dist / "kwentin_office-0.2.0-py3-none-any.whl"
        with zipfile.ZipFile(wheel, "w") as archive:
            for name in sorted(WHEEL_MEMBERS - ({omit_wheel} if omit_wheel else set())):
                if name == f"{DIST_INFO}/entry_points.txt":
                    payload = entry_points
                elif name == f"{DIST_INFO}/METADATA":
                    payload = metadata
                elif name == f"{DIST_INFO}/top_level.txt":
                    payload = top_level
                else:
                    payload = b"x"
                archive.writestr(name, payload)
            if extra_wheel:
                archive.writestr(extra_wheel, b"x")
            if duplicate_wheel:
                archive.writestr(duplicate_wheel, b"duplicate")
        sdist = self.dist / "kwentin_office-0.2.0.tar.gz"
        with tarfile.open(sdist, "w:gz") as archive:
            for suffix in sorted(SDIST_MEMBERS - ({omit_sdist} if omit_sdist else set())):
                payload = b"x"
                info = tarfile.TarInfo(f"kwentin_office-0.2.0/{suffix}")
                info.size = len(payload)
                archive.addfile(info, BytesIO(payload))
            if extra_sdist:
                payload = b"x"
                info = tarfile.TarInfo(f"kwentin_office-0.2.0/{extra_sdist}")
                info.size = len(payload)
                archive.addfile(info, BytesIO(payload))

    def verify(self) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [sys.executable, str(VERIFIER_PATH), str(self.dist)],
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
                self.assertNotEqual(self.verify().returncode, 0, member)

    def test_each_required_sdist_member_is_fail_closed(self) -> None:
        for member in sorted(SDIST_MEMBERS):
            with self.subTest(member=member):
                self.write_archives(omit_sdist=member)
                self.assertNotEqual(self.verify().returncode, 0, member)

    def test_all_five_console_scripts_are_required(self) -> None:
        self.write_archives(entry_points=b"[console_scripts]\noffice-docx = office_artifact_tool.__main__:main\n")
        self.assertNotEqual(self.verify().returncode, 0)

    def test_runtime_module_subtraction_is_refused(self) -> None:
        self.write_archives(omit_wheel="pptx_ai_composer/preview.py")
        self.assertNotEqual(self.verify().returncode, 0)

    def test_metadata_and_top_level_contents_are_closed(self) -> None:
        for kwargs in (
            {"metadata": VALID_METADATA.replace(b"Version: 0.2.0", b"Version: 9.9.9")},
            {"top_level": b"office_artifact_tool\n"},
        ):
            with self.subTest(kwargs=kwargs):
                self.write_archives(**kwargs)
                self.assertNotEqual(self.verify().returncode, 0)

    def test_unexpected_wheel_runtime_member_is_refused(self) -> None:
        self.write_archives(extra_wheel="pptx_ai_composer/private_oracle.py")
        self.assertNotEqual(self.verify().returncode, 0)

    def test_unexpected_sdist_member_is_refused(self) -> None:
        self.write_archives(extra_sdist="private-release-note.txt")
        self.assertNotEqual(self.verify().returncode, 0)

    def test_duplicate_archive_member_is_refused(self) -> None:
        self.write_archives(duplicate_wheel="pptx_ai_composer/preview.py")
        self.assertNotEqual(self.verify().returncode, 0)


if __name__ == "__main__":
    unittest.main()
