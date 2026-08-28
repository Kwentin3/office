from __future__ import annotations

import re
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from office_artifact_tool import DocxArtifactTool, render_docx_preview
from office_artifact_tool.docx.presentation import resolve_presentation
from office_artifact_tool.docx.preview import _css


class PresentationParityTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def test_preview_and_docx_share_explicit_a4_page_and_body_typography(self) -> None:
        model = {"blocks": [{"type": "paragraph", "text": "Body", "style": "Normal"}]}
        preview_dir = self.root / "review"
        output = self.root / "document.docx"

        preview = render_docx_preview(model, preview_dir)
        created = DocxArtifactTool(self.root / "work").create(model, output)

        self.assertEqual(preview["status"], "previewed")
        self.assertEqual(created["status"], "ok")
        html = (preview_dir / "document.html").read_text(encoding="utf-8")
        variables = dict(re.findall(r"--([a-z0-9-]+):([^;]+);", html))
        self.assertEqual(variables["page-width"], "210mm")
        self.assertEqual(variables["page-height"], "297mm")
        self.assertEqual(variables["margin-top"], "18mm")
        self.assertEqual(variables["margin-right"], "18mm")
        self.assertEqual(variables["margin-bottom"], "18mm")
        self.assertEqual(variables["margin-left"], "18mm")
        self.assertEqual(variables["body-font"], 'Arial,"Liberation Sans",sans-serif')
        self.assertEqual(variables["body-size"], "10.5pt")
        self.assertIn('class="document-page"', html)

        document = Document(output)
        section = document.sections[0]
        self.assertAlmostEqual(section.page_width.mm, 210, places=1)
        self.assertAlmostEqual(section.page_height.mm, 297, places=1)
        self.assertAlmostEqual(section.top_margin.mm, 18, places=1)
        self.assertAlmostEqual(section.right_margin.mm, 18, places=1)
        self.assertAlmostEqual(section.bottom_margin.mm, 18, places=1)
        self.assertAlmostEqual(section.left_margin.mm, 18, places=1)
        normal = document.styles["Normal"]
        self.assertEqual(normal.font.name, "Arial")
        self.assertAlmostEqual(normal.font.size.pt, 10.5, places=1)

    def test_preview_and_docx_share_managed_heading_and_paragraph_styles(self) -> None:
        model = {
            "blocks": [
                {"type": "paragraph", "text": "Candidate", "style": "Title"},
                {"type": "paragraph", "text": "Role", "style": "Subtitle"},
                {"type": "heading", "level": 1, "text": "Summary"},
                {"type": "heading", "level": 2, "text": "Experience"},
                {"type": "heading", "level": 3, "text": "Company"},
                {"type": "bulleted_list", "items": ["Result"]},
                {"type": "table", "style": "Table Grid", "rows": [["A", "B"], ["1", "2"]]},
            ]
        }
        preview_dir = self.root / "styled-review"
        output = self.root / "styled.docx"

        render_docx_preview(model, preview_dir)
        DocxArtifactTool(self.root / "work").create(model, output)

        html = (preview_dir / "document.html").read_text(encoding="utf-8")
        variables = dict(re.findall(r"--([a-z0-9-]+):([^;]+);", html))
        self.assertEqual(variables["title-size"], "24pt")
        self.assertEqual(variables["title-color"], "#17365D")
        self.assertEqual(variables["heading-1-size"], "20pt")
        self.assertEqual(variables["heading-1-color"], "#17365D")
        self.assertEqual(variables["heading-2-size"], "12pt")
        self.assertEqual(variables["heading-2-color"], "#365F91")
        self.assertIn('class="word-style word-style-title"', html)
        self.assertIn('class="word-style word-style-subtitle"', html)
        self.assertIn('class="word-style word-style-heading-1"', html)
        self.assertIn('class="word-style word-style-list-bullet"', html)
        self.assertRegex(html, r"\.word-style-list-bullet\{[^}]*margin-left:0mm")
        self.assertIn(
            "ul.word-style,ol.word-style{padding-left:6.35mm;margin-top:0;margin-bottom:0}",
            html,
        )
        self.assertIn('class="word-table word-table-grid"', html)

        document = Document(output)
        expected = {
            "Title": (24, "17365D", True),
            "Heading 1": (20, "17365D", True),
            "Heading 2": (12, "365F91", True),
            "Heading 3": (11, "4F81BD", True),
        }
        for name, (size, color, bold) in expected.items():
            style = document.styles[name]
            self.assertAlmostEqual(style.font.size.pt, size, places=1)
            self.assertEqual(str(style.font.color.rgb), color)
            self.assertIs(style.font.bold, bold)
        self.assertAlmostEqual(document.styles["List Bullet"].paragraph_format.left_indent.mm, 6.35, places=1)

    def test_preview_and_create_refuse_styles_outside_the_managed_presentation_contract(self) -> None:
        cases = [
            {"blocks": [{"type": "paragraph", "text": "Header", "style": "Header"}]},
            {"blocks": [{"type": "paragraph", "text": "Use the list block", "style": "List Bullet"}]},
            {"blocks": [{"type": "table", "style": "Light Shading Accent 1", "rows": [["A"]]}]},
        ]
        tool = DocxArtifactTool(self.root / "work")

        for index, model in enumerate(cases):
            preview_dir = self.root / f"unsupported-review-{index}"
            output = self.root / f"unsupported-{index}.docx"
            preview = render_docx_preview(model, preview_dir)
            created = tool.create(model, output)

            expected = {
                "status": "refused",
                "reason": "unsupported_capability",
                "details": "style is outside the managed DOCX presentation contract",
            }
            self.assertEqual(preview, expected)
            self.assertEqual(created, expected)
            self.assertFalse(preview_dir.exists())
            self.assertFalse(output.exists())

    def test_preview_and_create_share_typed_refusal_for_unstringifiable_table_integer(self) -> None:
        huge_integer = 1 << 16_610
        model = {"blocks": [{"type": "table", "rows": [[huge_integer]]}]}
        preview_dir = self.root / "huge-integer-review"
        output = self.root / "huge-integer.docx"
        expected = {
            "status": "refused",
            "reason": "validation_failure",
            "details": "table cells must be safely renderable finite XML-compatible scalar values",
        }

        self.assertEqual(render_docx_preview(model, preview_dir), expected)
        self.assertEqual(DocxArtifactTool(self.root / "work").create(model, output), expected)
        self.assertFalse(preview_dir.exists())
        self.assertFalse(output.exists())

    def test_review_identity_binds_the_managed_presentation_profile(self) -> None:
        model = {"document_id": "identity-test", "blocks": [{"type": "paragraph", "text": "Body"}]}
        preview_dir = self.root / "identity-review"

        result = render_docx_preview(model, preview_dir)

        expected_payload = {
            "model": model,
            "presentation_id": "professional-a4/v2",
        }
        expected_revision = __import__("hashlib").sha256(
            __import__("json").dumps(
                expected_payload,
                ensure_ascii=False,
                sort_keys=True,
                separators=(",", ":"),
            ).encode("utf-8")
        ).hexdigest()
        self.assertEqual(result["revision"], expected_revision)
        review = __import__("json").loads((preview_dir / "review.json").read_text(encoding="utf-8"))
        self.assertEqual(review["contract_version"], "1.1")
        self.assertEqual(review["presentation_id"], "professional-a4/v2")
        self.assertEqual(review["fidelity"], "styled_layout_proxy_not_word_render")
        self.assertIn("line wrapping and pagination remain browser approximations", review["limitations"])

    def test_preview_and_docx_share_table_grid_borders_and_cell_padding(self) -> None:
        model = {
            "blocks": [
                {"type": "table", "style": "Table Grid", "rows": [["A", "B"], ["1", "2"]]},
            ]
        }
        preview_dir = self.root / "table-review"
        output = self.root / "table.docx"

        render_docx_preview(model, preview_dir)
        DocxArtifactTool(self.root / "work").create(model, output)

        html = (preview_dir / "document.html").read_text(encoding="utf-8")
        variables = dict(re.findall(r"--([a-z0-9-]+):([^;]+);", html))
        self.assertEqual(variables["table-border-color"], "#B9C2CC")
        self.assertEqual(variables["table-border-width"], "0.75pt")
        self.assertEqual(variables["table-cell-padding"], "1.6mm")
        self.assertIn(
            '<td><p class="word-style word-style-normal">A</p></td>',
            html,
        )

        document = Document(output)
        table_properties = document.tables[0]._tbl.tblPr
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            borders = table_properties.xpath(f"./w:tblBorders/w:{edge}")
            self.assertEqual(len(borders), 1)
            self.assertEqual(borders[0].get(f"{{{namespace}}}val"), "single")
            self.assertEqual(borders[0].get(f"{{{namespace}}}sz"), "6")
            self.assertEqual(borders[0].get(f"{{{namespace}}}color"), "B9C2CC")
        for edge in ("top", "left", "bottom", "right"):
            cell_margin = table_properties.xpath(f"./w:tblCellMar/w:{edge}")
            self.assertEqual(len(cell_margin), 1)
            self.assertEqual(cell_margin[0].get(f"{{{namespace}}}w"), "91")

    def test_source_packaged_and_agent_create_schemas_expose_only_managed_styles(self) -> None:
        package_root = Path(__file__).resolve().parents[1]
        schema_paths = [
            package_root / "schemas" / "create.schema.json",
            package_root / "office_artifact_tool" / "resources" / "create.schema.json",
            package_root / "agent" / "schemas" / "create.schema.json",
        ]
        schemas = [__import__("json").loads(path.read_text(encoding="utf-8")) for path in schema_paths]

        self.assertEqual(schemas[1:], [schemas[0], schemas[0]])
        paragraph_styles = schemas[0]["$defs"]["paragraph"]["properties"]["style"]["enum"]
        table_styles = schemas[0]["$defs"]["table"]["properties"]["style"]["enum"]
        self.assertEqual(
            paragraph_styles,
            ["Heading 1", "Heading 2", "Heading 3", "Heading 4", "Heading 5", "Heading 6", "Heading 7", "Heading 8", "Heading 9", "Intense Quote", "Normal", "Quote", "Subtitle", "Title"],
        )
        self.assertEqual(table_styles, ["Table Grid"])

    def test_ragged_table_preview_pads_rows_to_the_same_shape_as_docx(self) -> None:
        model = {
            "blocks": [
                {"type": "table", "style": "Table Grid", "rows": [["A", "B"], ["1"]]},
            ]
        }
        preview_dir = self.root / "ragged-review"
        output = self.root / "ragged.docx"

        render_docx_preview(model, preview_dir)
        DocxArtifactTool(self.root / "work").create(model, output)

        html = (preview_dir / "document.html").read_text(encoding="utf-8")
        self.assertEqual(html.count("<td>"), 4)
        self.assertIn(
            '<tr><td><p class="word-style word-style-normal">1</p></td>'
            '<td><p class="word-style word-style-normal"></p></td></tr>',
            html,
        )
        self.assertIn("table-layout:fixed", html)
        self.assertIn('.word-table .word-style-normal:empty::after{content:"\\00a0"}', html)
        document = Document(output)
        self.assertEqual(len(document.tables[0].rows), 2)
        self.assertEqual(len(document.tables[0].columns), 2)
        self.assertEqual(document.tables[0].cell(1, 1).text, "")
        self.assertFalse(document.tables[0].autofit)

    def test_css_page_rule_is_derived_from_the_resolved_page_contract(self) -> None:
        presentation = resolve_presentation()
        custom_page = replace(
            presentation.page,
            width_mm=216,
            height_mm=279,
            margin_top_mm=12,
            margin_right_mm=13,
            margin_bottom_mm=14,
            margin_left_mm=15,
        )

        css = _css(replace(presentation, page=custom_page))

        self.assertIn("@page{size:216mm 279mm;margin:12mm 13mm 14mm 15mm}", css)

    def test_heading_levels_seven_to_nine_use_valid_block_markup(self) -> None:
        model = {
            "blocks": [
                {"type": "heading", "level": 7, "text": "Seven"},
                {"type": "heading", "level": 8, "text": "Eight"},
                {"type": "heading", "level": 9, "text": "Nine"},
            ]
        }
        preview_dir = self.root / "deep-headings"

        render_docx_preview(model, preview_dir)

        html = (preview_dir / "document.html").read_text(encoding="utf-8")
        for level, text in ((7, "Seven"), (8, "Eight"), (9, "Nine")):
            self.assertNotIn(f"<h{level}", html)
            self.assertIn(
                f'<p role="heading" aria-level="{level}" class="word-style word-style-heading-{level}">{text}</p>',
                html,
            )

    def test_list_item_spacing_and_table_spacing_do_not_invent_renderer_defaults(self) -> None:
        css = _css(resolve_presentation())

        self.assertIn("ul.word-style,ol.word-style{padding-left:6.35mm;margin-top:0;margin-bottom:0}", css)
        self.assertIn(".word-style-list-bullet li{margin-bottom:2pt}", css)
        self.assertIn(".word-style-list-number li{margin-bottom:2pt}", css)
        self.assertIn(".word-table{width:100%;table-layout:fixed;border-collapse:collapse;margin:0}", css)

    def test_every_managed_paragraph_style_translates_from_one_contract(self) -> None:
        presentation = resolve_presentation()
        output = self.root / "all-styles.docx"
        model = {"blocks": [{"type": "paragraph", "text": "Body"}]}

        DocxArtifactTool(self.root / "work").create(model, output)
        document = Document(output)
        css = _css(presentation)

        for configured in presentation.paragraph_styles:
            with self.subTest(style=configured.name):
                token = configured.name.casefold().replace(" ", "-")
                rule_match = re.search(rf"\.word-style-{re.escape(token)}\{{([^}}]+)\}}", css)
                self.assertIsNotNone(rule_match)
                rule = rule_match.group(1)
                css_left_indent = 0 if configured.name in {"List Bullet", "List Number"} else configured.left_indent_mm
                self.assertIn(f"font-size:var(--{token}-size)", rule)
                self.assertIn(f"color:var(--{token}-color)", rule)
                self.assertIn(f"font-weight:{700 if configured.bold else 400}", rule)
                self.assertIn(f"font-style:{'italic' if configured.italic else 'normal'}", rule)
                self.assertIn(f"margin-top:{configured.space_before_pt:g}pt", rule)
                self.assertIn(f"margin-bottom:{configured.space_after_pt:g}pt", rule)
                self.assertIn(f"margin-left:{css_left_indent:g}mm", rule)

                style = document.styles[configured.name]
                self.assertEqual(style.font.name, presentation.body.font_name)
                self.assertAlmostEqual(style.font.size.pt, configured.size_pt, places=1)
                self.assertEqual(str(style.font.color.rgb), configured.color)
                self.assertIs(style.font.bold, configured.bold)
                self.assertIs(style.font.italic, configured.italic)
                self.assertAlmostEqual(style.paragraph_format.line_spacing, presentation.body.line_height, places=2)
                self.assertAlmostEqual(style.paragraph_format.space_before.pt, configured.space_before_pt, places=1)
                self.assertAlmostEqual(style.paragraph_format.space_after.pt, configured.space_after_pt, places=1)
                self.assertAlmostEqual(style.paragraph_format.left_indent.mm, configured.left_indent_mm, places=1)
                self.assertIs(style.paragraph_format.keep_with_next, configured.keep_with_next)

    def test_managed_word_styles_drop_uncontracted_template_visual_defaults(self) -> None:
        presentation = resolve_presentation()
        output = self.root / "normalized-styles.docx"
        model = {"blocks": [{"type": "paragraph", "text": "Body"}]}

        DocxArtifactTool(self.root / "work").create(model, output)
        document = Document(output)

        for configured in presentation.paragraph_styles:
            with self.subTest(style=configured.name):
                element = document.styles[configured.name]._element
                for path in (
                    "./w:pPr/w:pBdr",
                    "./w:pPr/w:contextualSpacing",
                    "./w:pPr/w:keepLines",
                    "./w:rPr/w:spacing",
                    "./w:rPr/w:kern",
                ):
                    self.assertFalse(element.xpath(path), path)
                if configured.name not in {"List Bullet", "List Number"}:
                    self.assertFalse(element.xpath("./w:pPr/w:numPr"))

                fonts = element.xpath("./w:rPr/w:rFonts")
                self.assertEqual(len(fonts), 1)
                for script in ("ascii", "hAnsi", "eastAsia", "cs"):
                    self.assertEqual(fonts[0].get(qn(f"w:{script}")), presentation.body.font_name)
                for theme in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
                    self.assertIsNone(fonts[0].get(qn(f"w:{theme}")))

                size = str(round(configured.size_pt * 2))
                self.assertEqual(element.xpath("./w:rPr/w:sz/@w:val"), [size])
                self.assertEqual(element.xpath("./w:rPr/w:szCs/@w:val"), [size])
                self.assertEqual(element.xpath("./w:rPr/w:bCs/@w:val"), ["1" if configured.bold else "0"])
                self.assertEqual(element.xpath("./w:rPr/w:iCs/@w:val"), ["1" if configured.italic else "0"])

        intense_quote = document.styles["Intense Quote"]._element
        self.assertIsNone(intense_quote.xpath("./w:pPr/w:ind")[0].get(qn("w:right")))


if __name__ == "__main__":
    unittest.main()
