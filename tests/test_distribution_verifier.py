from __future__ import annotations

import base64
import csv
import hashlib
import importlib.util
import io
import subprocess
import sys
import tarfile
import tempfile
import tomllib
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
VALID_METADATA = b"Metadata-Version: 2.4\nName: kwentin-office\nVersion: 0.4.0\nRequires-Python: >=3.11\n"
VALID_WHEEL = b"Wheel-Version: 1.0\nGenerator: kwentin-tests\nRoot-Is-Purelib: true\nTag: py3-none-any\n"
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
        wheel_metadata: bytes = VALID_WHEEL,
        record_override: bytes | None = None,
        top_level: bytes = VALID_TOP_LEVEL,
        extra_wheel: str | None = None,
        extra_sdist: str | None = None,
        duplicate_wheel: str | None = None,
        wheel_name: str = "kwentin_office-0.4.0-py3-none-any.whl",
        sdist_name: str = "kwentin_office-0.4.0.tar.gz",
    ) -> None:
        wheel = self.dist / wheel_name
        members = sorted(WHEEL_MEMBERS - ({omit_wheel} if omit_wheel else set()))
        payloads: dict[str, bytes] = {}
        for name in members:
            if name == f"{DIST_INFO}/RECORD":
                continue
            if name == f"{DIST_INFO}/entry_points.txt":
                payload = entry_points
            elif name == f"{DIST_INFO}/METADATA":
                payload = metadata
            elif name == f"{DIST_INFO}/WHEEL":
                payload = wheel_metadata
            elif name == f"{DIST_INFO}/top_level.txt":
                payload = top_level
            else:
                payload = b"x"
            payloads[name] = payload
        if extra_wheel:
            payloads[extra_wheel] = b"x"
        record_name = f"{DIST_INFO}/RECORD"
        if record_name in members:
            if record_override is None:
                stream = io.StringIO(newline="")
                writer = csv.writer(stream, lineterminator="\n")
                for name, payload in sorted(payloads.items()):
                    digest = base64.urlsafe_b64encode(hashlib.sha256(payload).digest()).rstrip(b"=").decode("ascii")
                    writer.writerow((name, f"sha256={digest}", str(len(payload))))
                writer.writerow((record_name, "", ""))
                payloads[record_name] = stream.getvalue().encode("utf-8")
            else:
                payloads[record_name] = record_override
        with zipfile.ZipFile(wheel, "w") as archive:
            for name in members:
                if name in payloads:
                    archive.writestr(name, payloads[name])
            if duplicate_wheel:
                archive.writestr(duplicate_wheel, b"duplicate")
        sdist = self.dist / sdist_name
        with tarfile.open(sdist, "w:gz") as archive:
            for suffix in sorted(SDIST_MEMBERS - ({omit_sdist} if omit_sdist else set())):
                payload = b"x"
                info = tarfile.TarInfo(f"kwentin_office-0.4.0/{suffix}")
                info.size = len(payload)
                archive.addfile(info, BytesIO(payload))
            if extra_sdist:
                payload = b"x"
                info = tarfile.TarInfo(f"kwentin_office-0.4.0/{extra_sdist}")
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
            {"metadata": VALID_METADATA.replace(b"Version: 0.4.0", b"Version: 9.9.9")},
            {"top_level": b"office_artifact_tool\n"},
        ):
            with self.subTest(kwargs=kwargs):
                self.write_archives(**kwargs)
                self.assertNotEqual(self.verify().returncode, 0)

    def test_wheel_metadata_and_record_contents_are_verified(self) -> None:
        for kwargs in (
            {"wheel_metadata": b"x"},
            {"record_override": b"x"},
        ):
            with self.subTest(kwargs=kwargs):
                self.write_archives(**kwargs)
                self.assertNotEqual(self.verify().returncode, 0)

    def test_wheel_compatibility_headers_are_exact_singletons(self) -> None:
        for wheel_metadata in (
            VALID_WHEEL + b"Tag: cp312-cp312-manylinux_2_17_x86_64\n",
            VALID_WHEEL + b"Root-Is-Purelib: false\n",
            VALID_WHEEL + b"Wheel-Version: 2.0\n",
        ):
            with self.subTest(wheel_metadata=wheel_metadata):
                self.write_archives(wheel_metadata=wheel_metadata)
                self.assertNotEqual(self.verify().returncode, 0)

    def test_artifact_filenames_are_bound_to_name_version_and_tag(self) -> None:
        for kwargs in (
            {"wheel_name": "renamed-0.4.0-py3-none-any.whl"},
            {"wheel_name": "kwentin_office-0.4.0-cp312-cp312-manylinux_2_17_x86_64.whl"},
            {"sdist_name": "renamed-0.4.0.tar.gz"},
            {"sdist_name": "kwentin_office-9.9.9.tar.gz"},
        ):
            with self.subTest(kwargs=kwargs):
                for artifact in self.dist.iterdir():
                    artifact.unlink()
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

    def test_ci_clean_install_asserts_current_project_version(self) -> None:
        with (ROOT / "pyproject.toml").open("rb") as stream:
            version = tomllib.load(stream)["project"]["version"]
        workflow = (ROOT / ".github" / "workflows" / "ci.yml").read_text(encoding="utf-8")
        self.assertIn(f"m.version('kwentin-office') == '{version}'", workflow)

    def test_ci_installs_and_runs_authoritative_pytest_suite(self) -> None:
        workflow = (ROOT / ".github" / "workflows" / "ci.yml").read_text(encoding="utf-8")
        self.assertIn("pip install --upgrade pip build pytest", workflow)
        self.assertIn("python -m pytest -q --import-mode=importlib", workflow)


if __name__ == "__main__":
    unittest.main()
