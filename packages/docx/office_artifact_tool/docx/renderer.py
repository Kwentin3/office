from __future__ import annotations

from pathlib import Path
from typing import Any

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor
from docx.table import Table

from ..core.errors import ArtifactError
from .presentation import DocxPresentation, TablePresentation, resolve_presentation

_ALLOWED = {"heading", "paragraph", "numbered_list", "bulleted_list", "table"}


def _remove_children(parent: Any, *names: str) -> None:
    for name in names:
        for child in parent.findall(qn(f"w:{name}")):
            parent.remove(child)


def _set_on_off(parent: Any, name: str, enabled: bool) -> None:
    _remove_children(parent, name)
    child = OxmlElement(f"w:{name}")
    child.set(qn("w:val"), "1" if enabled else "0")
    parent.append(child)


def _set_half_points(parent: Any, name: str, points: float) -> None:
    _remove_children(parent, name)
    child = OxmlElement(f"w:{name}")
    child.set(qn("w:val"), str(round(points * 2)))
    parent.append(child)


def _normalize_style_xml(style: Any, *, font_name: str, size_pt: float, bold: bool, italic: bool) -> None:
    element = style._element
    paragraph_properties = element.get_or_add_pPr()
    _remove_children(paragraph_properties, "pBdr", "contextualSpacing", "keepLines")
    if style.name not in {"List Bullet", "List Number"}:
        _remove_children(paragraph_properties, "numPr")
    indent = paragraph_properties.find(qn("w:ind"))
    if indent is not None:
        indent.attrib.pop(qn("w:right"), None)
        if style.name not in {"List Bullet", "List Number"}:
            indent.attrib.pop(qn("w:hanging"), None)
            indent.attrib.pop(qn("w:firstLine"), None)

    run_properties = element.get_or_add_rPr()
    _remove_children(run_properties, "spacing", "kern")
    fonts = run_properties.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        run_properties.insert(0, fonts)
    for theme in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        fonts.attrib.pop(qn(f"w:{theme}"), None)
    for script in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(f"w:{script}"), font_name)
    _set_half_points(run_properties, "szCs", size_pt)
    _set_on_off(run_properties, "bCs", bold)
    _set_on_off(run_properties, "iCs", italic)


def _apply_presentation(document: DocumentObject, presentation: DocxPresentation) -> None:
    page = presentation.page
    section = document.sections[0]
    section.page_width = Mm(page.width_mm)
    section.page_height = Mm(page.height_mm)
    section.top_margin = Mm(page.margin_top_mm)
    section.right_margin = Mm(page.margin_right_mm)
    section.bottom_margin = Mm(page.margin_bottom_mm)
    section.left_margin = Mm(page.margin_left_mm)

    for configured in presentation.paragraph_styles:
        style = document.styles[configured.name]
        style.font.name = presentation.body.font_name
        style.font.size = Pt(configured.size_pt)
        style.font.color.rgb = RGBColor.from_string(configured.color)
        style.font.bold = configured.bold
        style.font.italic = configured.italic
        style.paragraph_format.line_spacing = presentation.body.line_height
        style.paragraph_format.space_before = Pt(configured.space_before_pt)
        style.paragraph_format.space_after = Pt(configured.space_after_pt)
        style.paragraph_format.left_indent = Mm(configured.left_indent_mm)
        style.paragraph_format.keep_with_next = configured.keep_with_next
        _normalize_style_xml(
            style,
            font_name=presentation.body.font_name,
            size_pt=configured.size_pt,
            bold=configured.bold,
            italic=configured.italic,
        )


def _set_table_presentation(table: Table, presentation: TablePresentation) -> None:
    properties = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    border_size = str(round(presentation.border_width_pt * 8))
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), border_size)
        border.set(qn("w:color"), presentation.border_color)
        borders.append(border)
    properties.append(borders)

    margins = OxmlElement("w:tblCellMar")
    margin_twips = str(round(presentation.cell_padding_mm / 25.4 * 1440))
    for edge in ("top", "left", "bottom", "right"):
        margin = OxmlElement(f"w:{edge}")
        margin.set(qn("w:w"), margin_twips)
        margin.set(qn("w:type"), "dxa")
        margins.append(margin)
    properties.append(margins)


def render(model: dict[str, Any], output: Path) -> None:
    if not isinstance(model, dict) or not isinstance(model.get("blocks"), list):
        raise ArtifactError("validation_failure", "model.blocks must be a list")

    presentation = resolve_presentation()
    document = Document()
    _apply_presentation(document, presentation)
    metadata = model.get("metadata", {})
    if not isinstance(metadata, dict):
        raise ArtifactError("validation_failure", "metadata must be an object")
    for key in ("title", "subject", "author", "keywords", "comments"):
        if key in metadata and isinstance(metadata[key], str):
            setattr(document.core_properties, key, metadata[key])

    for block in model["blocks"]:
        if not isinstance(block, dict) or block.get("type") not in _ALLOWED:
            raise ArtifactError("unsupported_capability", "unsupported create block")
        kind = block["type"]
        if kind == "heading":
            text = block.get("text")
            level = block.get("level", 1)
            if not isinstance(text, str) or not isinstance(level, int) or not 1 <= level <= 9:
                raise ArtifactError("validation_failure", "invalid heading")
            document.add_heading(text, level=level)
        elif kind == "paragraph":
            text = block.get("text")
            style = block.get("style")
            if not isinstance(text, str) or (style is not None and not isinstance(style, str)):
                raise ArtifactError("validation_failure", "invalid paragraph")
            try:
                document.add_paragraph(text, style=style)
            except KeyError as error:
                raise ArtifactError("validation_failure", "unknown paragraph style") from error
        elif kind in {"numbered_list", "bulleted_list"}:
            items = block.get("items")
            style = "List Number" if kind == "numbered_list" else "List Bullet"
            if not isinstance(items, list) or not items or not all(isinstance(item, str) for item in items):
                raise ArtifactError("validation_failure", "invalid list")
            for item in items:
                document.add_paragraph(item, style=style)
        elif kind == "table":
            rows = block.get("rows")
            style = block.get("style", "Table Grid")
            if not isinstance(rows, list) or not rows or not all(isinstance(row, list) and row for row in rows):
                raise ArtifactError("validation_failure", "invalid table")
            width = max(map(len, rows))
            table = document.add_table(rows=len(rows), cols=width)
            table.autofit = False
            try:
                table.style = style
            except KeyError as error:
                raise ArtifactError("validation_failure", "unknown table style") from error
            _set_table_presentation(table, presentation.table)
            for row_index, row in enumerate(rows):
                for cell_index, value in enumerate(row):
                    table.cell(row_index, cell_index).text = str(value)

    document.save(output)
