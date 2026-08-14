from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from office_artifact_tool import DocxArtifactTool
from PIL import Image


class DocxPreservationMatrixTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.tool = DocxArtifactTool(self.root / "work")

    def tearDown(self) -> None:
        self.temp.cleanup()

    @staticmethod
    def members(path: Path) -> dict[str, bytes]:
        with zipfile.ZipFile(path) as archive:
            return {info.filename: archive.read(info.filename) for info in archive.infolist()}

    def rich_document(self) -> Path:
        source = self.root / "rich.docx"
        image = self.root / "image.png"
        Image.new("RGB", (8, 8), "red").save(image)
        document = Document()
        paragraph = document.add_paragraph()
        relationship = document.part.relate_to(
            "https://example.invalid/preserve",
            RELATIONSHIP_TYPE.HYPERLINK,
            is_external=True,
        )
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), relationship)
        run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = "Linked text"
        run.append(text)
        hyperlink.append(run)
        paragraph._p.append(hyperlink)
        document.add_picture(str(image))
        table = document.add_table(rows=1, cols=2)
        table.cell(0, 0).text = "Merged text"
        table.cell(0, 0).merge(table.cell(0, 1))
        document.save(source)
        return source

    def test_text_replacement_preserves_hyperlink_relationship_and_image_member(self) -> None:
        source = self.rich_document()
        before = self.members(source)
        snapshot = self.tool.inspect(source)
        target = next(item for item in snapshot["elements"] if item.get("text") == "Linked text")
        plan = self.tool.plan(
            snapshot,
            {"operations": [{"type": "replace_text", "target_id": target["id"], "old": "Linked text", "new": "Changed text"}]},
        )["plan"]
        output = self.root / "replaced.docx"
        result = self.tool.apply(source, plan, output)
        self.assertEqual(result["status"], "ok")
        after = self.members(output)
        self.assertEqual(set(after), set(before))
        self.assertEqual({name for name in before if before[name] != after[name]}, {"word/document.xml"})
        self.assertEqual(after["word/_rels/document.xml.rels"], before["word/_rels/document.xml.rels"])
        media = [name for name in before if name.startswith("word/media/")]
        self.assertTrue(media)
        for name in media:
            self.assertEqual(after[name], before[name])
        self.assertIn(b"Changed text", after["word/document.xml"])
        self.assertIn(b"w:hyperlink", after["word/document.xml"])

    def test_merged_cell_text_edit_preserves_merge_properties_and_out_of_scope_members(self) -> None:
        source = self.rich_document()
        before = self.members(source)
        snapshot = self.tool.inspect(source)
        target = next(item for item in snapshot["elements"] if item.get("kind") == "cell" and item.get("text") == "Merged text")
        plan = self.tool.plan(
            snapshot,
            {"operations": [{"type": "set_cell_text", "target_id": target["id"], "text": "Changed merged"}]},
        )["plan"]
        output = self.root / "merged.docx"
        result = self.tool.apply(source, plan, output)
        self.assertEqual(result["status"], "ok")
        after = self.members(output)
        self.assertEqual({name for name in before if before[name] != after[name]}, {"word/document.xml"})
        self.assertIn(b"w:gridSpan", after["word/document.xml"])
        self.assertIn(b"Changed merged", after["word/document.xml"])
        for name in before:
            if name != "word/document.xml":
                self.assertEqual(after[name], before[name], name)


if __name__ == "__main__":
    unittest.main()
