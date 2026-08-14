import json
import tempfile
import unittest
from pathlib import Path
from xml.etree import ElementTree

from PIL import Image
from pptx_ai_composer.compiler import compile_deck
from pptx_ai_composer.preview import PreviewError, render_preview, render_scene_preview
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

    def test_scene_backend_renders_without_semantic_deck(self):
        scene = compile_deck(valid_deck_spec(), slide_ids=["s_compare"])
        result = render_scene_preview(scene, self.output)
        self.assertEqual(result["slide_count"], 1)
        manifest = json.loads((self.output / "manifest.json").read_text(encoding="utf-8"))
        self.assertEqual(manifest["slides"][0]["slide_id"], "s_compare")
        svg = (self.output / "slide-01.svg").read_text(encoding="utf-8")
        self.assertIn("Five tools", svg)
        self.assertIn('data-role="comparison.left.items"', svg)

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
            "sha256": __import__("hashlib").sha256(asset.read_bytes()).hexdigest(), "alt_text": "Hero",
        }]
        deck["slides"][0]["asset_id"] = "hero"
        with self.assertRaisesRegex(PreviewError, "contain a protected path"):
            render_preview(deck, self.output)
        self.assertTrue(asset.exists())

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
