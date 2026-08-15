from __future__ import annotations

import hashlib
import importlib.util
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image

MODULE_PATH = Path(__file__).resolve().parents[1] / "examples/openwebui_backend/office_service.py"
DECK_FIXTURE = Path(__file__).resolve().parents[1] / "examples/pptx-composer/managed-library.deck.json"
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
        with patch.object(MODULE.shutil, "copyfileobj", side_effect=OSError("copy failed")), self.assertRaisesRegex(
            OSError, "copy failed"
        ):
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

        with patch.object(MODULE.os, "open", side_effect=racing_open), self.assertRaisesRegex(
            ValueError, "workspace changed"
        ):
            service.stage_upload("req_race", upload, ".docx")
        self.assertFalse((outside / "source.docx").exists())
        self.assertFalse((request / "input-original/source.docx").exists())

    def test_application_witness_requires_absolute_host_executable(self) -> None:
        service = OfficeService(self.root / "service")
        with self.assertRaisesRegex(ValueError, "absolute"):
            service.application_witness("req_witness", "soffice")
        witness = service.application_witness("req_witness", Path("/bin/false"))
        self.assertEqual(Path(witness.executable), Path("/bin/false"))

    def test_docx_chat_review_is_isolated_and_returns_webui_media_artifact(self) -> None:
        service = OfficeService(self.root / "service")
        model = {
            "document_id": "invitation",
            "metadata": {"title": "Invitation"},
            "blocks": [
                {"block_id": "title", "type": "heading", "level": 1, "text": "Приглашение"},
                {"block_id": "body", "type": "paragraph", "text": "Будем рады видеть вас."},
            ],
        }

        result = service.docx_chat_review("req_docx_preview", model)

        request_root = service.request_root("req_docx_preview")
        review_path = Path(result["review_contract"])
        self.assertTrue(review_path.is_relative_to(request_root / "internal/docx-preview"))
        self.assertEqual(result["interaction"], "chat_only")
        self.assertRegex(result["revision"], r"^[0-9a-f]{64}$")
        self.assertEqual(len(result["display_artifacts"]), 1)
        artifact = Path(result["display_artifacts"][0])
        self.assertTrue(artifact.is_file())
        self.assertEqual(artifact.suffix, ".html")
        self.assertTrue(artifact.is_relative_to(request_root / "internal/docx-preview"))
        html = artifact.read_text(encoding="utf-8")
        self.assertIn('data-block-id="title"', html)
        self.assertNotIn("<script", html.lower())
        self.assertNotIn("http://", html.lower())
        self.assertNotIn("https://", html.lower())
        review = json.loads(review_path.read_text(encoding="utf-8"))
        self.assertEqual(review["revision"], result["revision"])
        self.assertEqual(review["interaction"], "chat_only")
        self.assertEqual(review["kind"], "docx_chat_review")

    def test_directory_swap_during_chat_review_cannot_publish_outside_request_workspace(self) -> None:
        service = OfficeService(self.root / "service")
        request = service.request_root("req_preview_race")
        outside = self.root / "outside-preview"
        outside.mkdir()
        domain = request / "internal/docx-preview"
        real_render = MODULE.render_docx_preview
        swapped = False

        def swap_then_render(model, output):
            nonlocal swapped
            if not swapped:
                swapped = True
                domain.rename(request / "internal/docx-preview-original")
                domain.symlink_to(outside, target_is_directory=True)
            return real_render(model, output)

        model = {"blocks": [{"type": "paragraph", "text": "safe"}]}
        with patch.object(MODULE, "render_docx_preview", side_effect=swap_then_render):
            with self.assertRaisesRegex(ValueError, "workspace changed"):
                service.docx_chat_review("req_preview_race", model)

        self.assertFalse((outside / "review/review.json").exists())
        self.assertFalse((outside / "review/document.html").exists())

    def test_request_directory_swap_during_each_chat_review_is_refused(self) -> None:
        deck = json.loads(DECK_FIXTURE.read_text(encoding="utf-8"))
        cases = (
            ("docx", MODULE.render_docx_preview, {"blocks": [{"type": "paragraph", "text": "safe"}]}),
            ("xlsx", MODULE.render_xlsx_preview, {"sheets": [{"name": "Data", "cells": {"A1": {"value": "safe"}}}]}),
            ("pptx", MODULE.render_preview, deck),
        )

        for domain, real_render, model in cases:
            with self.subTest(domain=domain):
                service = OfficeService(self.root / f"service-{domain}")
                request_id = f"req-ancestor-{domain}"
                request = service.request_root(request_id)
                renamed = service.root / f"{request_id}-original"
                replacement = self.root / f"replacement-{domain}"
                replacement.mkdir()

                def swap_request_then_render(candidate_model, output):
                    request.rename(renamed)
                    request.symlink_to(replacement, target_is_directory=True)
                    return real_render(candidate_model, output)

                method = getattr(service, f"{domain}_chat_review")
                patch_name = "render_preview" if domain == "pptx" else f"render_{domain}_preview"
                with patch.object(MODULE, patch_name, side_effect=swap_request_then_render):
                    with self.assertRaisesRegex(ValueError, "workspace changed"):
                        method(request_id, model)

                self.assertFalse((replacement / "internal").exists())

    def test_xlsx_chat_review_is_isolated_and_returns_visible_sheet_media(self) -> None:
        service = OfficeService(self.root / "service")
        model = {
            "workbook_id": "budget",
            "sheets": [
                {
                    "name": "Summary",
                    "cells": {
                        "A1": {"value": "Показатель", "style": "header"},
                        "B1": {"value": "Значение", "style": "header"},
                        "A2": {"value": "Выручка"},
                        "B2": {"value": 125000, "style": "currency"},
                    },
                },
                {"name": "Internal", "state": "hidden", "cells": {"A1": {"value": "secret"}}},
            ],
        }

        result = service.xlsx_chat_review("req_xlsx_preview", model)

        request_root = service.request_root("req_xlsx_preview")
        review_path = Path(result["review_contract"])
        self.assertTrue(review_path.is_relative_to(request_root / "internal/xlsx-preview"))
        self.assertEqual(result["interaction"], "chat_only")
        self.assertRegex(result["revision"], r"^[0-9a-f]{64}$")
        self.assertEqual(len(result["display_artifacts"]), 1)
        artifact = Path(result["display_artifacts"][0])
        self.assertTrue(artifact.is_file())
        self.assertEqual(artifact.suffix, ".html")
        self.assertTrue(artifact.is_relative_to(request_root / "internal/xlsx-preview"))
        html = artifact.read_text(encoding="utf-8")
        self.assertIn('data-sheet="Summary"', html)
        self.assertIn('data-cell="B2"', html)
        self.assertNotIn("<script", html.lower())
        self.assertNotIn("http://", html.lower())
        self.assertNotIn("https://", html.lower())
        review = json.loads(review_path.read_text(encoding="utf-8"))
        self.assertEqual(review["revision"], result["revision"])
        self.assertEqual(review["interaction"], "chat_only")
        self.assertEqual(review["kind"], "xlsx_chat_review")
        self.assertEqual([sheet["sheet"] for sheet in review["sheets"]], ["Summary"])

    def test_pptx_chat_review_is_isolated_and_returns_webui_media_artifacts(self) -> None:
        service = OfficeService(self.root / "service")
        deck = json.loads(DECK_FIXTURE.read_text(encoding="utf-8"))

        result = service.pptx_chat_review("req_preview", deck)

        request_root = service.request_root("req_preview")
        review_path = Path(result["review_contract"])
        self.assertTrue(review_path.is_relative_to(request_root / "internal/pptx-composer-preview"))
        self.assertEqual(result["interaction"], "chat_only")
        self.assertRegex(result["revision"], r"^[0-9a-f]{64}$")
        self.assertTrue(result["display_artifacts"])
        for artifact in result["display_artifacts"]:
            path = Path(artifact)
            self.assertTrue(path.is_file())
            self.assertTrue(path.is_relative_to(request_root / "internal/pptx-composer-preview"))
        review = json.loads(review_path.read_text(encoding="utf-8"))
        self.assertEqual(review["revision"], result["revision"])
        self.assertEqual(review["interaction"], "chat_only")

    def test_pptx_chat_review_requires_host_approved_assets(self) -> None:
        service = OfficeService(self.root / "service")
        deck = json.loads(DECK_FIXTURE.read_text(encoding="utf-8"))
        asset = self.root / "hero.png"
        Image.new("RGB", (32, 32), "navy").save(asset)
        deck["assets"] = [{
            "asset_id": "hero",
            "kind": "png",
            "path": str(asset),
            "sha256": hashlib.sha256(asset.read_bytes()).hexdigest(),
            "alt_text": "Approved hero",
        }]
        deck["slides"] = [{
            "slide_id": "cover",
            "archetype": "cover",
            "title": "Approved asset",
            "subtitle": "Host-owned preview input",
            "asset_id": "hero",
            "source_ids": [],
        }]

        with self.assertRaisesRegex(ValueError, "host-approved"):
            service.pptx_chat_review("req_unapproved", deck)

        result = service.pptx_chat_review("req_approved", deck, allowed_asset_paths=(asset,))
        self.assertEqual(result["status"], "previewed")

    def test_pptx_chat_review_refuses_approved_asset_with_wrong_hash(self) -> None:
        service = OfficeService(self.root / "service")
        deck = json.loads(DECK_FIXTURE.read_text(encoding="utf-8"))
        asset = self.root / "wrong-hash.png"
        Image.new("RGB", (32, 32), "navy").save(asset)
        deck["assets"] = [{
            "asset_id": "hero",
            "kind": "png",
            "path": str(asset),
            "sha256": "0" * 64,
            "alt_text": "Approved path with unauthenticated bytes",
        }]
        deck["slides"] = [{
            "slide_id": "cover",
            "archetype": "cover",
            "title": "Hash-bound review",
            "subtitle": "Must refuse before publication",
            "asset_id": "hero",
            "source_ids": [],
        }]

        with self.assertRaisesRegex(ValueError, "asset hash mismatch: hero"):
            service.pptx_chat_review("req_wrong_hash", deck, allowed_asset_paths=(asset,))

        review = service.request_root("req_wrong_hash") / "internal/pptx-composer-preview/review/review.json"
        self.assertFalse(review.exists())

    def test_pptx_chat_review_returns_typed_refusal_for_missing_allowed_asset(self) -> None:
        service = OfficeService(self.root / "service")
        deck = json.loads(DECK_FIXTURE.read_text(encoding="utf-8"))
        missing = self.root / "missing.png"
        deck["assets"] = [{
            "asset_id": "hero",
            "kind": "png",
            "path": str(missing),
            "sha256": "a" * 64,
            "alt_text": "Missing hero",
        }]
        deck["slides"] = [{
            "slide_id": "cover",
            "archetype": "cover",
            "title": "Missing asset",
            "subtitle": "Must refuse cleanly",
            "asset_id": "hero",
            "source_ids": [],
        }]

        with self.assertRaisesRegex(ValueError, "host-approved"):
            service.pptx_chat_review("req_missing", deck, allowed_asset_paths=(missing,))

    def test_pptx_chat_review_normalizes_missing_deck_asset_refusal(self) -> None:
        service = OfficeService(self.root / "service")
        deck = json.loads(DECK_FIXTURE.read_text(encoding="utf-8"))
        missing = self.root / "model-selected-missing.png"
        deck["assets"] = [{
            "asset_id": "hero",
            "kind": "png",
            "path": str(missing),
            "sha256": "a" * 64,
            "alt_text": "Missing model-selected hero",
        }]
        deck["slides"] = [{
            "slide_id": "cover",
            "archetype": "cover",
            "title": "Missing asset",
            "subtitle": "Must refuse deterministically",
            "asset_id": "hero",
            "source_ids": [],
        }]

        with self.assertRaisesRegex(ValueError, "host-approved"):
            service.pptx_chat_review("req_missing_model_path", deck)


if __name__ == "__main__":
    unittest.main()
