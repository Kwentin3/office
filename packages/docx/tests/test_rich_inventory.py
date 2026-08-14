from __future__ import annotations

import copy
import json
import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
from lxml import etree
from office_artifact_tool import DocxArtifactTool
from office_artifact_tool.docx.inventory import BLOCKER_FEATURES, WARNING_FEATURES

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DOCX_FEATURES = {
    "tracked_changes",
    "content_controls",
    "comments",
    "images",
    "fields",
    "alt_chunks",
    "text_boxes",
    "merged_cells",
    "hyperlinks",
    "external_relationships",
    "headers",
    "footers",
    "footnotes",
    "endnotes",
    "charts",
    "smartart",
    "embedded_packages",
    "ole_objects",
    "activex",
    "document_protection",
    "macros",
    "signatures",
}


class RichInventoryTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.tool = DocxArtifactTool(self.root / "work")
        self.source = self.root / "source.docx"
        document = Document()
        document.add_paragraph("Normal")
        document.save(self.source)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def rewrite_document(self, mutate) -> None:
        rewritten = self.root / "rewritten.docx"
        with zipfile.ZipFile(self.source) as source, zipfile.ZipFile(rewritten, "w") as output:
            for info in source.infolist():
                payload = source.read(info.filename)
                if info.filename == "word/document.xml":
                    root = etree.fromstring(payload)
                    mutate(root)
                    payload = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                output.writestr(copy.copy(info), payload)
        rewritten.replace(self.source)

    def test_inventory_reports_safe_baseline(self) -> None:
        result = self.tool.inspect(self.source, view="inventory")
        self.assertEqual(result["status"], "ok")
        self.assertEqual(result["schema_version"], 2)
        self.assertEqual(result["findings"], [])
        self.assertFalse(result["findings_truncated"])
        self.assertEqual(set(result["features"]), DOCX_FEATURES)
        self.assertEqual(result["mutation_policy"]["decision"], "safe")
        self.assertEqual(result["features"]["tracked_changes"], 0)
        self.assertEqual(result["features"]["comments"], 0)

    def test_inventory_schema_matches_policy_partitions(self) -> None:
        schema = json.loads(
            (Path(__file__).parents[1] / "office_artifact_tool/resources/inventory.schema.json").read_text()
        )
        policy = schema["properties"]["mutation_policy"]
        self.assertEqual(set(policy["properties"]["blockers"]["items"]["enum"]), set(BLOCKER_FEATURES))
        self.assertEqual(set(policy["properties"]["warnings"]["items"]["enum"]), set(WARNING_FEATURES))
        constraints = {
            rule["if"]["properties"]["decision"]["const"]: rule["then"]["properties"] for rule in policy["allOf"]
        }
        self.assertEqual(constraints["safe"]["blockers"]["maxItems"], 0)
        self.assertEqual(constraints["safe"]["warnings"]["maxItems"], 0)
        self.assertEqual(constraints["safe_with_warnings"]["warnings"]["minItems"], 1)
        self.assertEqual(constraints["refuse_mutation"]["blockers"]["minItems"], 1)

    def test_inventory_schema_v2_bounds_location_findings(self) -> None:
        schema = json.loads(
            (Path(__file__).parents[1] / "office_artifact_tool/resources/inventory.schema.json").read_text()
        )
        finding = schema["properties"]["findings"]
        self.assertEqual(schema["properties"]["schema_version"]["const"], 2)
        self.assertIn("findings", schema["required"])
        self.assertIn("findings_truncated", schema["required"])
        self.assertEqual(finding["maxItems"], 1000)
        self.assertEqual(finding["items"]["required"], ["feature", "scope", "part"])
        self.assertEqual(finding["items"]["properties"]["part"]["pattern"], "^(?!/)(?!.*(?:^|/)\\.\\.(?:/|$)).+$")

    def test_tracked_changes_are_inventory_blocker_and_apply_refuses(self) -> None:
        def add_tracked_change(root) -> None:
            paragraph = root.find(f".//{{{W}}}p")
            insertion = etree.SubElement(paragraph, f"{{{W}}}ins")
            run = etree.SubElement(insertion, f"{{{W}}}r")
            etree.SubElement(run, f"{{{W}}}t").text = "Tracked"

        self.rewrite_document(add_tracked_change)
        inventory = self.tool.inspect(self.source, view="inventory")
        self.assertEqual(inventory["features"]["tracked_changes"], 1)
        self.assertEqual(inventory["mutation_policy"]["decision"], "refuse_mutation")
        snapshot = self.tool.inspect(self.source)
        target = next(item for item in snapshot["elements"] if item["kind"] == "paragraph")
        planned = self.tool.plan(
            snapshot,
            {"operations": [{"type": "replace_text", "target_id": target["id"], "old": "Normal", "new": "Changed"}]},
        )
        result = self.tool.apply(self.source, planned["plan"], self.root / "out.docx")
        self.assertEqual(result["status"], "refused")
        self.assertEqual(result["reason"], "unsupported_capability")

    def test_inventory_findings_locate_tracked_change_in_exact_story_target(self) -> None:
        def add_tracked_change(root) -> None:
            paragraph = root.find(f".//{{{W}}}p")
            insertion = etree.SubElement(paragraph, f"{{{W}}}ins")
            run = etree.SubElement(insertion, f"{{{W}}}r")
            etree.SubElement(run, f"{{{W}}}t").text = "Tracked"

        self.rewrite_document(add_tracked_change)
        target = next(item for item in self.tool.inspect(self.source)["elements"] if item["kind"] == "paragraph")
        first = self.tool.inspect(self.source, view="inventory")
        second = self.tool.inspect(self.source, view="inventory")
        expected = {
            "feature": "tracked_changes",
            "scope": "story",
            "part": "word/document.xml",
            "target_id": target["id"],
        }
        self.assertEqual(first["findings"], [expected])
        self.assertEqual(second["findings"], first["findings"])

    def test_inventory_findings_locate_hyperlink_in_exact_story_target(self) -> None:
        def add_hyperlink(root) -> None:
            paragraph = root.find(f".//{{{W}}}p")
            run = paragraph.find(f"{{{W}}}r")
            hyperlink = etree.Element(f"{{{W}}}hyperlink")
            hyperlink.set(f"{{{W}}}anchor", "local")
            paragraph.replace(run, hyperlink)
            hyperlink.append(run)

        self.rewrite_document(add_hyperlink)
        target = next(item for item in self.tool.inspect(self.source)["elements"] if item["kind"] == "paragraph")
        inventory = self.tool.inspect(self.source, view="inventory")
        self.assertEqual(inventory["features"]["hyperlinks"], 1)
        self.assertIn(
            {"feature": "hyperlinks", "scope": "story", "part": "word/document.xml", "target_id": target["id"]},
            inventory["findings"],
        )

    def test_inventory_findings_locate_merged_cell_target(self) -> None:
        document = Document(self.source)
        table = document.add_table(rows=1, cols=2)
        table.cell(0, 0).text = "Merged"
        table.cell(0, 0).merge(table.cell(0, 1))
        document.save(self.source)
        target = next(
            item for item in self.tool.inspect(self.source)["elements"]
            if item["kind"] == "cell" and item["text"] == "Merged"
        )
        inventory = self.tool.inspect(self.source, view="inventory")
        self.assertGreaterEqual(inventory["features"]["merged_cells"], 1)
        self.assertIn(
            {"feature": "merged_cells", "scope": "story", "part": "word/document.xml", "target_id": target["id"]},
            inventory["findings"],
        )

    def test_apply_relaxes_story_blocker_only_outside_resolved_target_parts(self) -> None:
        document = Document(self.source)
        document.sections[0].header.paragraphs[0].text = "Header"
        document.save(self.source)
        rewritten = self.root / "header-tracked.docx"
        with zipfile.ZipFile(self.source) as source, zipfile.ZipFile(rewritten, "w") as output:
            for info in source.infolist():
                payload = source.read(info.filename)
                if info.filename.startswith("word/header") and info.filename.endswith(".xml"):
                    root = etree.fromstring(payload)
                    paragraph = root.find(f".//{{{W}}}p")
                    insertion = etree.SubElement(paragraph, f"{{{W}}}ins")
                    run = etree.SubElement(insertion, f"{{{W}}}r")
                    etree.SubElement(run, f"{{{W}}}t").text = "Tracked"
                    payload = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                output.writestr(copy.copy(info), payload)
        rewritten.replace(self.source)
        inventory = self.tool.inspect(self.source, view="inventory")
        self.assertEqual(inventory["mutation_policy"]["decision"], "refuse_mutation")
        snapshot = self.tool.inspect(self.source)
        target = next(item for item in snapshot["elements"] if item["story_part"] == "word/document.xml")
        plan = self.tool.plan(
            snapshot,
            {"operations": [{"type": "replace_text", "target_id": target["id"], "old": "Normal", "new": "Changed"}]},
        )["plan"]
        output = self.root / "location-aware.docx"
        result = self.tool.apply(self.source, plan, output)
        self.assertEqual(result["status"], "ok")
        self.assertEqual(result["validation"]["changed_members"], ["word/document.xml"])

    def test_supported_header_is_counted_without_warning(self) -> None:
        document = Document(self.source)
        document.sections[0].header.paragraphs[0].text = "Supported story"
        document.save(self.source)
        inventory = self.tool.inspect(self.source, view="inventory")
        self.assertEqual(inventory["features"]["headers"], 1)
        self.assertEqual(inventory["mutation_policy"]["decision"], "safe")

    def test_document_protection_is_blocker(self) -> None:
        rewritten = self.root / "protected.docx"
        with zipfile.ZipFile(self.source) as source, zipfile.ZipFile(rewritten, "w") as output:
            for info in source.infolist():
                payload = source.read(info.filename)
                if info.filename == "word/settings.xml":
                    root = etree.fromstring(payload)
                    etree.SubElement(root, f"{{{W}}}documentProtection")
                    payload = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                output.writestr(copy.copy(info), payload)
        rewritten.replace(self.source)
        inventory = self.tool.inspect(self.source, view="inventory")
        self.assertEqual(inventory["features"]["document_protection"], 1)
        self.assertEqual(inventory["mutation_policy"]["decision"], "refuse_mutation")


if __name__ == "__main__":
    unittest.main()
