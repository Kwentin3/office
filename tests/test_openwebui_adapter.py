from __future__ import annotations

import importlib.util
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

MODULE_PATH = Path(__file__).resolve().parents[1] / "examples/openwebui_backend/office_service.py"
SPEC = importlib.util.spec_from_file_location("office_service_example", MODULE_PATH)
assert SPEC is not None and SPEC.loader is not None
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)
OfficeService = MODULE.OfficeService


class OfficeServicePathSafetyTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_refuses_symlink_workspace_root(self) -> None:
        real = self.root / "real"
        real.mkdir()
        link = self.root / "workspace-link"
        link.symlink_to(real, target_is_directory=True)
        with self.assertRaisesRegex(ValueError, "symlink"):
            OfficeService(link)

    def test_refuses_symlink_request_directories(self) -> None:
        service = OfficeService(self.root / "service")
        request = service.request_root("req_1")
        outside = self.root / "outside"
        outside.mkdir()
        upload = self.root / "upload.docx"
        upload.write_bytes(b"opaque upload bytes")

        (request / "input").symlink_to(outside, target_is_directory=True)
        with self.assertRaisesRegex(ValueError, "symlink"):
            service.stage_upload("req_1", upload, ".docx")
        self.assertFalse((outside / "source.docx").exists())

        (request / "input").unlink()
        (request / "output").symlink_to(outside, target_is_directory=True)
        with self.assertRaisesRegex(ValueError, "symlink"):
            service.output_path("req_1", "result.docx")

    def test_stages_exclusively_inside_request_workspace(self) -> None:
        service = OfficeService(self.root / "service")
        upload = self.root / "upload.xlsx"
        upload.write_bytes(b"opaque upload bytes")
        staged = service.stage_upload("req_2", upload, ".xlsx")
        self.assertEqual(staged.parent.parent, service.request_root("req_2"))
        self.assertEqual(staged.read_bytes(), upload.read_bytes())
        with self.assertRaises(FileExistsError):
            service.stage_upload("req_2", upload, ".xlsx")
    def test_partial_staged_upload_is_removed_on_copy_failure(self) -> None:
        service = OfficeService(self.root / "service")
        upload = self.root / "upload.pptx"
        upload.write_bytes(b"opaque upload bytes")
        with patch.object(MODULE.shutil, "copyfileobj", side_effect=OSError("copy failed")):
            with self.assertRaisesRegex(OSError, "copy failed"):
                service.stage_upload("req_3", upload, ".pptx")
        self.assertFalse((service.request_root("req_3") / "input/source.pptx").exists())

    def test_directory_swap_during_stage_cannot_escape_request_workspace(self) -> None:
        service = OfficeService(self.root / "service")
        upload = self.root / "upload.docx"
        upload.write_bytes(b"RACE")
        request = service.request_root("req_race")
        input_dir = request / "input"
        input_dir.mkdir()
        outside = self.root / "outside"
        outside.mkdir()
        real_open = MODULE.os.open
        swapped = False

        def racing_open(path, flags, mode=0o777, *, dir_fd=None):
            nonlocal swapped
            if not swapped and Path(path).name == "source.docx":
                swapped = True
                input_dir.rename(request / "input-original")
                input_dir.symlink_to(outside, target_is_directory=True)
            return real_open(path, flags, mode, dir_fd=dir_fd)

        with patch.object(MODULE.os, "open", side_effect=racing_open):
            with self.assertRaisesRegex(ValueError, "workspace changed"):
                service.stage_upload("req_race", upload, ".docx")
        self.assertFalse((outside / "source.docx").exists())
        self.assertFalse((request / "input-original/source.docx").exists())


if __name__ == "__main__":
    unittest.main()
