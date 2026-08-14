from __future__ import annotations

import copy
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from docx import Document
from lxml import etree
from office_artifact_tool import DocxArtifactTool


class StrictTemplateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.tool = DocxArtifactTool(self.root / "work")
        self.source = self.root / "template.docx"
        document = Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("Customer: {{cus")
        paragraph.add_run("tomer}}")
        table = document.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "Amount: {{amount}}"
        document.sections[0].header.paragraphs[0].text = "Contract {{number}}"
        document.save(self.source)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_strict_fill_resolves_split_runs_tables_and_headers(self) -> None:
        output = self.root / "filled.docx"
        result = self.tool.fill_template(
            self.source,
            {"customer": "Acme", "amount": "100", "number": "42"},
            output,
        )
        self.assertEqual(result["status"], "ok")
        snapshot = self.tool.inspect(output)
        texts = [item.get("text", "") for item in snapshot["elements"]]
        self.assertIn("Customer: Acme", texts)
        self.assertIn("Amount: 100", texts)
        self.assertIn("Contract 42", texts)
        self.assertFalse(any("{{" in text or "}}" in text for text in texts))

    def test_strict_fill_refuses_missing_and_unknown_values_without_output(self) -> None:
        for values in (
            {"customer": "Acme", "amount": "100"},
            {"customer": "Acme", "amount": "100", "number": "42", "extra": "x"},
        ):
            with self.subTest(values=values):
                output = self.root / "never.docx"
                result = self.tool.fill_template(self.source, values, output)
                self.assertEqual(result["status"], "refused")
                self.assertEqual(result["reason"], "validation_failure")
                self.assertFalse(output.exists())

    def test_strict_fill_preserves_non_token_runs(self) -> None:
        document = Document()
        paragraph = document.add_paragraph()
        prefix = paragraph.add_run("Prefix ")
        prefix.italic = True
        token = paragraph.add_run("{{name}}")
        token.bold = True
        paragraph = document.add_paragraph()
        prefix = paragraph.add_run("Second ")
        prefix.bold = True
        token = paragraph.add_run("{{name}}")
        token.italic = True
        cell = document.add_table(rows=1, cols=1).cell(0, 0)
        cell_paragraph = cell.paragraphs[0]
        cell_paragraph.clear()
        prefix = cell_paragraph.add_run("Cell ")
        prefix.italic = True
        token = cell_paragraph.add_run("{{name}}")
        token.bold = True
        source = self.root / "formatted.docx"
        document.save(source)
        output = self.root / "formatted-output.docx"
        result = self.tool.fill_template(source, {"name": "Acme"}, output)
        self.assertEqual(result["status"], "ok")
        rendered = Document(output).paragraphs[0]
        self.assertEqual(rendered.text, "Prefix Acme")
        self.assertTrue(rendered.runs[0].italic)
        self.assertTrue(rendered.runs[1].bold)
        rendered = Document(output).paragraphs[1]
        self.assertEqual(rendered.text, "Second Acme")
        self.assertTrue(rendered.runs[0].bold)
        self.assertTrue(rendered.runs[1].italic)
        rendered = Document(output).tables[0].cell(0, 0).paragraphs[0]
        self.assertEqual(rendered.text, "Cell Acme")
        self.assertTrue(rendered.runs[0].italic)
        self.assertTrue(rendered.runs[1].bold)

    def test_strict_fill_rejects_package_wide_marker(self) -> None:
        rewritten = self.root / "package-marker.docx"
        with zipfile.ZipFile(self.source) as package, zipfile.ZipFile(rewritten, "w") as target:
            for info in package.infolist():
                payload = package.read(info.filename)
                if info.filename == "docProps/core.xml":
                    root = etree.fromstring(payload)
                    root.set("{urn:kwentin:test}marker", "{{residual}}")
                    payload = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                target.writestr(copy.copy(info), payload)
        output = self.root / "package-marker-output.docx"
        result = self.tool.fill_template(
            rewritten,
            {"customer": "Acme", "amount": "100", "number": "42"},
            output,
        )
        self.assertEqual((result["status"], result["reason"]), ("refused", "unsupported_capability"))
        self.assertFalse(output.exists())

    def test_strict_fill_rejects_residual_marker_in_unsupported_candidate_scope(self) -> None:
        source = self.root / "residual-source.docx"
        document = Document()
        document.add_paragraph("{{name}}")
        document.save(source)
        output = self.root / "residual-output.docx"
        original_apply = self.tool.apply

        def inject_residual(source_path, plan, candidate):
            result = original_apply(source_path, plan, candidate)
            if result.get("status") == "ok":
                rewritten = self.root / "residual-candidate.docx"
                with zipfile.ZipFile(candidate) as package, zipfile.ZipFile(rewritten, "w") as target:
                    for info in package.infolist():
                        payload = package.read(info.filename)
                        if info.filename == "docProps/core.xml":
                            root = etree.fromstring(payload)
                            root.set("{urn:kwentin:test}residual", "{{residual}}")
                            payload = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                        target.writestr(copy.copy(info), payload)
                rewritten.replace(candidate)
            return result

        with mock.patch.object(self.tool, "apply", side_effect=inject_residual):
            result = self.tool.fill_template(source, {"name": "Acme"}, output)
        self.assertEqual((result["status"], result["reason"]), ("refused", "validation_failure"))
        self.assertFalse(output.exists())

    def test_strict_fill_refuses_malformed_overflow_and_complex_cell_scope(self) -> None:
        malformed = Document()
        malformed.add_paragraph("{{{name}}}")
        malformed_source = self.root / "malformed.docx"
        malformed.save(malformed_source)
        result = self.tool.fill_template(malformed_source, {"name": "Acme"}, self.root / "malformed-output.docx")
        self.assertEqual((result["status"], result["reason"]), ("refused", "validation_failure"))

        overflow = Document()
        overflow.add_paragraph("X{{name}}")
        overflow_source = self.root / "overflow.docx"
        overflow.save(overflow_source)
        result = self.tool.fill_template(overflow_source, {"name": "x" * 32767}, self.root / "overflow-output.docx")
        self.assertEqual((result["status"], result["reason"]), ("refused", "validation_failure"))

        complex_cell = Document()
        cell = complex_cell.add_table(rows=1, cols=1).cell(0, 0)
        cell.text = "Keep"
        cell.add_paragraph("{{name}}")
        complex_source = self.root / "complex-cell.docx"
        complex_cell.save(complex_source)
        result = self.tool.fill_template(complex_source, {"name": "Acme"}, self.root / "complex-output.docx")
        self.assertEqual((result["status"], result["reason"]), ("refused", "unsupported_capability"))

    def test_strict_fill_rejects_dtd_entity_declarations_package_wide(self) -> None:
        declaration = '<!DOCTYPE cp:coreProperties [<!ENTITY hidden "{{hidden}}">]>'
        for encoding in ("UTF-8", "UTF-16"):
            with self.subTest(encoding=encoding):
                rewritten = self.root / f"dtd-marker-{encoding}.docx"
                with zipfile.ZipFile(self.source) as package, zipfile.ZipFile(rewritten, "w") as target:
                    for info in package.infolist():
                        payload = package.read(info.filename)
                        if info.filename == "docProps/core.xml":
                            root = etree.fromstring(payload)
                            payload = etree.tostring(
                                root,
                                xml_declaration=True,
                                encoding=encoding,
                                doctype=declaration,
                            )
                        target.writestr(copy.copy(info), payload)
                output = self.root / f"dtd-output-{encoding}.docx"
                result = self.tool.fill_template(
                    rewritten,
                    {"customer": "Acme", "amount": "100", "number": "42"},
                    output,
                )
                self.assertEqual(result["status"], "refused")
                self.assertFalse(output.exists())

    def test_cli_fill_template_matches_api_contract(self) -> None:
        values = self.root / "values.json"
        values.write_text(json.dumps({"customer": "Acme", "amount": "100", "number": "42"}))
        output = self.root / "cli.docx"
        completed = subprocess.run(
            [
                sys.executable,
                "-m",
                "office_artifact_tool",
                "--workspace",
                str(self.root / "cli-work"),
                "fill-template",
                "--source",
                str(self.source),
                "--values",
                str(values),
                "--output",
                str(output),
            ],
            text=True,
            capture_output=True,
            check=False,
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)
        self.assertEqual(json.loads(completed.stdout)["status"], "ok")
        self.assertTrue(output.is_file())


if __name__ == "__main__":
    unittest.main()
