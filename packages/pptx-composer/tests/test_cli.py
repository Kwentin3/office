import json
import subprocess
import tempfile
import unittest
from pathlib import Path

from tests.fixtures import valid_deck_spec


class JsonCliTests(unittest.TestCase):
    def test_render_then_validate_over_json_lines(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            spec_path = root / "deck.json"
            output = root / "deck.pptx"
            spec_path.write_text(json.dumps(valid_deck_spec()), encoding="utf-8")
            render = subprocess.run(
                [
                    "/workspace/.venv-docx-study/bin/python",
                    "-m",
                    "pptx_ai_composer",
                ],
                input=json.dumps({"action": "render", "spec": str(spec_path), "output": str(output)}) + "\n",
                text=True,
                capture_output=True,
                cwd="/workspace/pptx-ai-composer",
                env={"PYTHONPATH": "."},
                check=False,
            )
            self.assertEqual(render.returncode, 0, render.stderr)
            render_result = json.loads(render.stdout)
            self.assertEqual(render_result["status"], "rendered")
            self.assertTrue(output.exists())

            validate = subprocess.run(
                ["/workspace/.venv-docx-study/bin/python", "-m", "pptx_ai_composer"],
                input=json.dumps({"action": "validate", "spec": str(spec_path), "source": str(output)}) + "\n",
                text=True,
                capture_output=True,
                cwd="/workspace/pptx-ai-composer",
                env={"PYTHONPATH": "."},
                check=False,
            )
            self.assertEqual(validate.returncode, 0, validate.stderr)
            report = json.loads(validate.stdout)
            self.assertEqual(report["status"], "valid_with_unexecuted_gates")

    def test_preview_and_catalog_over_json_lines(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            spec_path = root / "deck.json"
            preview_path = root / "preview"
            spec_path.write_text(json.dumps(valid_deck_spec()), encoding="utf-8")
            preview = subprocess.run(
                ["/workspace/.venv-docx-study/bin/python", "-m", "pptx_ai_composer"],
                input=json.dumps({"action": "preview", "spec": str(spec_path), "output": str(preview_path)}) + "\n",
                text=True,
                capture_output=True,
                cwd="/workspace/pptx-ai-composer",
                env={"PYTHONPATH": "."},
                check=False,
            )
            self.assertEqual(preview.returncode, 0, preview.stderr)
            self.assertEqual(json.loads(preview.stdout)["status"], "previewed")
            self.assertTrue((preview_path / "manifest.json").exists())
            catalog = subprocess.run(
                ["/workspace/.venv-docx-study/bin/python", "-m", "pptx_ai_composer"],
                input=json.dumps({"action": "catalog"}) + "\n",
                text=True,
                capture_output=True,
                cwd="/workspace/pptx-ai-composer",
                env={"PYTHONPATH": "."},
                check=False,
            )
            self.assertEqual(catalog.returncode, 0, catalog.stderr)
            payload = json.loads(catalog.stdout)
            self.assertIn("comparison", payload["archetypes"])
            self.assertIn("shape", payload["components"])

    def test_preview_can_select_one_slide_and_named_variant(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            spec_path = root / "deck.json"
            preview_path = root / "preview"
            spec_path.write_text(json.dumps(valid_deck_spec()), encoding="utf-8")
            result = subprocess.run(
                ["/workspace/.venv-docx-study/bin/python", "-m", "pptx_ai_composer"],
                input=json.dumps({
                    "action": "preview",
                    "spec": str(spec_path),
                    "output": str(preview_path),
                    "variants": {"s_compare": "compact"},
                    "slide_ids": ["s_compare"],
                }) + "\n",
                text=True,
                capture_output=True,
                cwd="/workspace/pptx-ai-composer",
                env={"PYTHONPATH": "."},
                check=False,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(json.loads(result.stdout)["slide_count"], 1)
            manifest = json.loads((preview_path / "manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(manifest["slides"][0]["variant"], "compact")

    def test_cli_returns_typed_error_for_unknown_action(self):
        result = subprocess.run(
            ["/workspace/.venv-docx-study/bin/python", "-m", "pptx_ai_composer"],
            input=json.dumps({"action": "explode"}) + "\n",
            text=True,
            capture_output=True,
            cwd="/workspace/pptx-ai-composer",
            env={"PYTHONPATH": "."},
            check=False,
        )
        self.assertEqual(result.returncode, 2)
        payload = json.loads(result.stdout)
        self.assertEqual(payload["status"], "error")
        self.assertEqual(payload["code"], "invalid_request")


if __name__ == "__main__":
    unittest.main()
