import hashlib
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from xml.etree import ElementTree

from PIL import Image
from pptx_ai_composer.compiler import compile_deck
from pptx_ai_composer.contracts import validate_deck_spec
from pptx_ai_composer.preview import PreviewError, render_preview, render_scene_preview
from pptx_ai_composer.scene_contract import validate_scene_spec

from tests.fixtures import valid_deck_spec


class FastPreviewTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tempdir.name)
        self.output = self.root / "preview"

    def tearDown(self):
        self.tempdir.cleanup()

    def test_generates_svg_per_slide_and_manifest(self):
        result = render_preview(valid_deck_spec(), self.output)
        self.assertEqual(result["status"], "previewed")
        self.assertEqual(result["slide_count"], 3)
        manifest = json.loads((self.output / "manifest.json").read_text(encoding="utf-8"))
        self.assertEqual(manifest["fidelity"], "structural_preview_not_powerpoint_render")
        self.assertEqual(len(manifest["slides"]), 3)
        for slide in manifest["slides"]:
            path = self.output / slide["file"]
            self.assertTrue(path.exists())
            root = ElementTree.fromstring(path.read_text(encoding="utf-8"))
            self.assertTrue(root.tag.endswith("svg"))
            png_path = self.output / slide["png_file"]
            self.assertTrue(png_path.exists())
            with Image.open(png_path) as image:
                self.assertEqual(image.size, (1280, 720))
        cover = (self.output / "slide-01.svg").read_text(encoding="utf-8")
        comparison = (self.output / "slide-02.svg").read_text(encoding="utf-8")
        self.assertIn("Hermes pilot", cover)
        self.assertIn("Five tools", comparison)
        self.assertNotIn("VERTICAL SLICE", cover)
        self.assertIn("Источник: Pilot workbook", comparison)
        self.assertIn("diagnostics", manifest)
        self.assertEqual(manifest["diagnostics"]["text_overflow"], [])

    def test_publishes_closed_chat_review_packet_bound_to_deck_revision(self):
        raw_deck = valid_deck_spec()
        result = render_preview(raw_deck, self.output)

        deck = validate_deck_spec(raw_deck)
        scene = validate_scene_spec(compile_deck(deck))
        encoded = json.dumps(
            {"deck": deck, "scene": scene},
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
        self.assertEqual(result["revision"], hashlib.sha256(encoded).hexdigest())

        review_path = self.output / "review.json"
        self.assertEqual(result["review_contract"], str(review_path.resolve()))
        self.assertEqual(result["interaction"], "chat_only")
        self.assertEqual(result["display_artifacts"], [
            str((self.output / "slide-01.png").resolve()),
            str((self.output / "slide-02.png").resolve()),
            str((self.output / "slide-03.png").resolve()),
        ])
        review = json.loads(review_path.read_text(encoding="utf-8"))
        self.assertEqual(set(review), {
            "contract_version", "kind", "interaction", "deck_id", "revision",
            "fidelity", "limitations", "diagnostics", "slides",
        })
        self.assertEqual(review["contract_version"], "1.0")
        self.assertEqual(review["kind"], "pptx_chat_review")
        self.assertEqual(review["interaction"], "chat_only")
        self.assertRegex(review["revision"], r"^[0-9a-f]{64}$")
        self.assertEqual(review["revision"], result["revision"])
        self.assertEqual([slide["slide_id"] for slide in review["slides"]], ["s_cover", "s_compare", "s_chart"])
        for number, slide in enumerate(review["slides"], start=1):
            self.assertEqual(set(slide), {"slide_id", "number", "png_file", "png_sha256", "svg_file", "svg_sha256"})
            self.assertEqual(slide["number"], number)
            self.assertRegex(slide["png_sha256"], r"^[0-9a-f]{64}$")
            self.assertRegex(slide["svg_sha256"], r"^[0-9a-f]{64}$")
            self.assertEqual(
                slide["png_sha256"],
                hashlib.sha256((self.output / slide["png_file"]).read_bytes()).hexdigest(),
            )
            self.assertEqual(
                slide["svg_sha256"],
                hashlib.sha256((self.output / slide["svg_file"]).read_bytes()).hexdigest(),
            )

    def test_reports_text_overflow_in_manifest_without_silent_truncation(self):
        deck = valid_deck_spec()
        deck["slides"][1]["left"]["items"] = [
            "Очень длинный пункт, который намеренно повторяет подробное объяснение для проверки диагностической границы preview backend." for _ in range(5)
        ]
        render_preview(deck, self.output)
        manifest = json.loads((self.output / "manifest.json").read_text(encoding="utf-8"))
        issues = manifest["diagnostics"]["text_overflow"]
        self.assertTrue(issues)
        self.assertEqual(issues[0]["slide_id"], "s_compare")
        self.assertIn("node_id", issues[0])
        self.assertGreater(issues[0]["required_lines"], issues[0]["available_lines"])

    def test_replaces_existing_preview_directory_atomically(self):
        self.output.mkdir()
        (self.output / "stale.txt").write_text("stale", encoding="utf-8")
        render_preview(valid_deck_spec(), self.output)
        self.assertFalse((self.output / "stale.txt").exists())
        self.assertTrue((self.output / "manifest.json").exists())

    def test_replacing_review_never_makes_published_directory_disappear(self):
        render_preview(valid_deck_spec(), self.output)
        from pptx_ai_composer import preview as preview_module

        real_exchange = preview_module._exchange_directories
        observations = []

        def observe_exchange(source, target):
            observations.append((self.output.exists(), (self.output / "review.json").exists()))
            real_exchange(source, target)
            observations.append((self.output.exists(), (self.output / "review.json").exists()))

        with patch.object(preview_module, "_exchange_directories", side_effect=observe_exchange):
            result = render_preview(valid_deck_spec(), self.output)

        self.assertEqual(result["status"], "previewed")
        self.assertTrue(observations)
        self.assertTrue(all(all(observation) for observation in observations))

    def test_failed_atomic_exchange_preserves_previous_complete_review(self):
        render_preview(valid_deck_spec(), self.output)
        previous_review = (self.output / "review.json").read_bytes()
        previous_png = (self.output / "slide-01.png").read_bytes()
        from pptx_ai_composer import preview as preview_module

        with patch.object(preview_module, "_exchange_directories", side_effect=OSError("exchange failed")):
            with self.assertRaisesRegex(OSError, "exchange failed"):
                render_preview(valid_deck_spec(), self.output)

        self.assertEqual((self.output / "review.json").read_bytes(), previous_review)
        self.assertEqual((self.output / "slide-01.png").read_bytes(), previous_png)

    def test_post_commit_cleanup_failure_returns_successful_new_review(self):
        first = render_preview(valid_deck_spec(), self.output)
        revised = valid_deck_spec()
        revised["slides"][0]["title"] = "Revised"
        from pptx_ai_composer import preview as preview_module

        original_replace = preview_module.os.replace
        backup = self.output.with_name(f".{self.output.name}.old")

        def fail_old_generation_rename(source, target):
            if Path(target) == backup:
                raise RecursionError("cleanup failed")
            original_replace(source, target)

        with patch.object(preview_module.os, "replace", side_effect=fail_old_generation_rename):
            second = render_preview(revised, self.output)

        self.assertNotEqual(first["revision"], second["revision"])
        self.assertIn("Revised", (self.output / "slide-01.svg").read_text(encoding="utf-8"))
        self.assertFalse(backup.exists())
        self.assertEqual([path for path in self.root.iterdir() if path != self.output], [])

    def test_review_discloses_that_image_nodes_are_placeholders(self):
        asset = self.root / "hero.png"
        Image.new("RGB", (20, 20), "red").save(asset)
        deck = valid_deck_spec()
        deck["assets"] = [{
            "asset_id": "hero", "kind": "png", "path": str(asset),
            "sha256": hashlib.sha256(asset.read_bytes()).hexdigest(), "alt_text": "Hero",
        }]
        deck["slides"][0]["asset_id"] = "hero"

        render_preview(deck, self.output)
        review = json.loads((self.output / "review.json").read_text(encoding="utf-8"))

        self.assertTrue(any("image" in item.lower() and "placeholder" in item.lower() for item in review["limitations"]))

    def test_chat_revision_rebuilds_review_from_changed_deck_spec(self):
        first_deck = valid_deck_spec()
        first = render_preview(first_deck, self.output)
        first_png = (self.output / "slide-02.png").read_bytes()

        revised_deck = valid_deck_spec()
        revised_deck["slides"][1]["title"] = "A shorter revised title"
        second = render_preview(revised_deck, self.output)
        review = json.loads((self.output / "review.json").read_text(encoding="utf-8"))

        self.assertNotEqual(first["revision"], second["revision"])
        self.assertEqual(second["revision"], review["revision"])
        self.assertNotEqual(first_png, (self.output / "slide-02.png").read_bytes())
        self.assertIn("A shorter revised title", (self.output / "slide-02.svg").read_text(encoding="utf-8"))

    def test_review_revision_binds_named_variant_and_slide_selection(self):
        first = render_preview(
            valid_deck_spec(),
            self.output,
            variants={"s_compare": "balanced"},
            slide_ids=["s_compare"],
        )
        second = render_preview(
            valid_deck_spec(),
            self.output,
            variants={"s_compare": "compact"},
            slide_ids=["s_compare"],
        )

        self.assertNotEqual(first["revision"], second["revision"])
        review = json.loads((self.output / "review.json").read_text(encoding="utf-8"))
        self.assertEqual(review["revision"], second["revision"])

    def test_scene_backend_renders_without_semantic_deck(self):
        scene = validate_scene_spec(compile_deck(valid_deck_spec(), slide_ids=["s_compare"]))
        result = render_scene_preview(scene, self.output)
        encoded = json.dumps(
            {"scene": scene},
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
        self.assertEqual(result["revision"], hashlib.sha256(encoded).hexdigest())
        self.assertEqual(result["slide_count"], 1)
        manifest = json.loads((self.output / "manifest.json").read_text(encoding="utf-8"))
        self.assertEqual(manifest["slides"][0]["slide_id"], "s_compare")
        svg = (self.output / "slide-01.svg").read_text(encoding="utf-8")
        self.assertIn("Five tools", svg)
        self.assertIn('data-role="comparison.left.items"', svg)

    def test_scene_backend_does_not_accept_caller_supplied_revision(self):
        scene = compile_deck(valid_deck_spec(), slide_ids=["s_compare"])

        with self.assertRaisesRegex(TypeError, "unexpected keyword argument 'revision'"):
            render_scene_preview(scene, self.output, revision="0" * 64)

    def test_preview_accepts_named_variant_and_slide_selection(self):
        result = render_preview(valid_deck_spec(), self.output, variants={"s_compare": "compact"}, slide_ids=["s_compare"])
        self.assertEqual(result["slide_count"], 1)
        manifest = json.loads((self.output / "manifest.json").read_text(encoding="utf-8"))
        self.assertEqual(manifest["slides"][0]["variant"], "compact")

    def test_refuses_output_collision_with_protected_path(self):
        with self.assertRaisesRegex(PreviewError, "protected"):
            render_preview(valid_deck_spec(), self.output, protected_paths=[self.output])

    def test_refuses_output_directory_that_contains_protected_input(self):
        spec = self.output / "deck.json"
        self.output.mkdir()
        spec.write_text("{}", encoding="utf-8")
        with self.assertRaisesRegex(PreviewError, "contain a protected path"):
            render_preview(valid_deck_spec(), self.output, protected_paths=[spec])
        self.assertTrue(spec.exists())

    def test_refuses_existing_non_directory_output(self):
        self.output.write_text("keep", encoding="utf-8")
        with self.assertRaisesRegex(PreviewError, "must be a directory"):
            render_preview(valid_deck_spec(), self.output)
        self.assertEqual(self.output.read_text(encoding="utf-8"), "keep")

    def test_refuses_preview_directory_that_contains_asset(self):
        asset = self.output / "hero.png"
        self.output.mkdir()
        Image.new("RGB", (20, 20), "red").save(asset)
        deck = valid_deck_spec()
        deck["assets"] = [{
            "asset_id": "hero", "kind": "png", "path": str(asset),
            "sha256": hashlib.sha256(asset.read_bytes()).hexdigest(), "alt_text": "Hero",
        }]
        deck["slides"][0]["asset_id"] = "hero"
        with self.assertRaisesRegex(PreviewError, "contain a protected path"):
            render_preview(deck, self.output)
        self.assertTrue(asset.exists())

    def test_refuses_asset_hash_mismatch_without_publishing_review(self):
        asset = self.root / "hero.png"
        Image.new("RGB", (20, 20), "red").save(asset)
        deck = valid_deck_spec()
        deck["assets"] = [{
            "asset_id": "hero", "kind": "png", "path": str(asset),
            "sha256": "0" * 64, "alt_text": "Hero",
        }]
        deck["slides"][0]["asset_id"] = "hero"

        with self.assertRaisesRegex(PreviewError, "asset hash mismatch: hero"):
            render_preview(deck, self.output)

        self.assertFalse((self.output / "review.json").exists())

    def test_refuses_hash_bound_non_image_without_publishing_review(self):
        asset = self.root / "hero.png"
        asset.write_bytes(b"this is not a PNG")
        deck = valid_deck_spec()
        deck["assets"] = [{
            "asset_id": "hero", "kind": "png", "path": str(asset),
            "sha256": hashlib.sha256(asset.read_bytes()).hexdigest(), "alt_text": "Hero",
        }]
        deck["slides"][0]["asset_id"] = "hero"

        with self.assertRaisesRegex(PreviewError, "invalid raster asset: hero"):
            render_preview(deck, self.output)

        self.assertFalse((self.output / "review.json").exists())

    def test_refuses_raster_declared_as_wrong_format(self):
        asset = self.root / "hero.jpg"
        Image.new("RGB", (20, 20), "red").save(asset, "PNG")
        deck = valid_deck_spec()
        deck["assets"] = [{
            "asset_id": "hero", "kind": "jpeg", "path": str(asset),
            "sha256": hashlib.sha256(asset.read_bytes()).hexdigest(), "alt_text": "Hero",
        }]
        deck["slides"][0]["asset_id"] = "hero"

        with self.assertRaisesRegex(PreviewError, "raster format mismatch: hero"):
            render_preview(deck, self.output)

        self.assertFalse((self.output / "review.json").exists())

    def test_refuses_raster_over_dimension_limit(self):
        asset = self.root / "hero.png"
        Image.new("RGB", (8193, 1), "red").save(asset)
        deck = valid_deck_spec()
        deck["assets"] = [{
            "asset_id": "hero", "kind": "png", "path": str(asset),
            "sha256": hashlib.sha256(asset.read_bytes()).hexdigest(), "alt_text": "Hero",
        }]
        deck["slides"][0]["asset_id"] = "hero"

        with self.assertRaisesRegex(PreviewError, "raster dimensions exceed limit: hero"):
            render_preview(deck, self.output)

        self.assertFalse((self.output / "review.json").exists())

    def test_refuses_asset_over_byte_limit_before_reading(self):
        asset = self.root / "hero.png"
        with asset.open("wb") as stream:
            stream.truncate(20 * 1024 * 1024 + 1)
        deck = valid_deck_spec()
        deck["assets"] = [{
            "asset_id": "hero", "kind": "png", "path": str(asset),
            "sha256": "0" * 64, "alt_text": "Hero",
        }]
        deck["slides"][0]["asset_id"] = "hero"

        with self.assertRaisesRegex(PreviewError, "asset exceeds byte limit: hero"):
            render_preview(deck, self.output)

        self.assertFalse((self.output / "review.json").exists())

    def test_refuses_fifo_asset_without_blocking_or_publishing_review(self):
        asset = self.root / "hero.png"
        os.mkfifo(asset)
        probe = """
import sys
from pathlib import Path
from pptx_ai_composer.preview import PreviewError, render_preview
from tests.fixtures import valid_deck_spec

deck = valid_deck_spec()
deck["assets"] = [{
    "asset_id": "hero", "kind": "png", "path": sys.argv[1],
    "sha256": "0" * 64, "alt_text": "Hero",
}]
deck["slides"][0]["asset_id"] = "hero"
try:
    render_preview(deck, Path(sys.argv[2]))
except PreviewError as exc:
    if str(exc) != "asset is missing or unsafe: hero":
        raise
else:
    raise AssertionError("FIFO asset was accepted")
"""
        subprocess.run(
            [sys.executable, "-c", probe, str(asset), str(self.output)],
            cwd=Path(__file__).parent.parent,
            check=True,
            capture_output=True,
            text=True,
            timeout=3,
        )
        self.assertFalse((self.output / "review.json").exists())

    def test_refuses_svg_fallback_hash_mismatch_without_publishing_review(self):
        svg = self.root / "hero.svg"
        fallback = self.root / "hero.png"
        svg.write_text(
            '<svg xmlns="http://www.w3.org/2000/svg"><rect width="20" height="20" fill="red"/></svg>',
            encoding="utf-8",
        )
        Image.new("RGB", (20, 20), "red").save(fallback)
        deck = valid_deck_spec()
        deck["assets"] = [{
            "asset_id": "hero",
            "kind": "svg",
            "path": str(svg),
            "sha256": hashlib.sha256(svg.read_bytes()).hexdigest(),
            "fallback_path": str(fallback),
            "fallback_sha256": "0" * 64,
            "alt_text": "Hero",
        }]
        deck["slides"][0]["asset_id"] = "hero"

        with self.assertRaisesRegex(PreviewError, "asset fallback hash mismatch: hero"):
            render_preview(deck, self.output)

        self.assertFalse((self.output / "review.json").exists())

    def test_refuses_hash_bound_non_png_svg_fallback(self):
        svg = self.root / "hero.svg"
        fallback = self.root / "hero.png"
        svg.write_text(
            '<svg xmlns="http://www.w3.org/2000/svg"><rect width="20" height="20" fill="red"/></svg>',
            encoding="utf-8",
        )
        fallback.write_bytes(b"this is not a PNG")
        deck = valid_deck_spec()
        deck["assets"] = [{
            "asset_id": "hero",
            "kind": "svg",
            "path": str(svg),
            "sha256": hashlib.sha256(svg.read_bytes()).hexdigest(),
            "fallback_path": str(fallback),
            "fallback_sha256": hashlib.sha256(fallback.read_bytes()).hexdigest(),
            "alt_text": "Hero",
        }]
        deck["slides"][0]["asset_id"] = "hero"

        with self.assertRaisesRegex(PreviewError, "invalid raster asset fallback: hero"):
            render_preview(deck, self.output)

        self.assertFalse((self.output / "review.json").exists())

    def test_refuses_backup_directory_that_contains_protected_input(self):
        backup = self.root / ".preview.old"
        backup.mkdir()
        spec = backup / "deck.json"
        spec.write_text("protected", encoding="utf-8")
        with self.assertRaisesRegex(PreviewError, "backup.*protected"):
            render_preview(valid_deck_spec(), self.output, protected_paths=[spec])
        self.assertEqual(spec.read_text(encoding="utf-8"), "protected")
        self.assertFalse(self.output.exists())

    def test_refuses_symlink_output_or_backup(self):
        target = self.root / "target"
        target.mkdir()
        (target / "keep.txt").write_text("keep", encoding="utf-8")
        self.output.symlink_to(target, target_is_directory=True)
        with self.assertRaisesRegex(PreviewError, "symlink"):
            render_preview(valid_deck_spec(), self.output)
        self.output.unlink()
        backup = self.root / ".preview.old"
        backup.symlink_to(target, target_is_directory=True)
        with self.assertRaisesRegex(PreviewError, "symlink"):
            render_preview(valid_deck_spec(), self.output)
        self.assertEqual((target / "keep.txt").read_text(encoding="utf-8"), "keep")


if __name__ == "__main__":
    unittest.main()
