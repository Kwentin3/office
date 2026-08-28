"""Closed presentation contract shared by DOCX create and preview.

This module owns document presentation meaning. Renderers only translate the
resolved values into their native representation; they must not invent layout
or typography defaults.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class PagePresentation:
    width_mm: float
    height_mm: float
    margin_top_mm: float
    margin_right_mm: float
    margin_bottom_mm: float
    margin_left_mm: float


@dataclass(frozen=True, slots=True)
class TextPresentation:
    font_name: str
    fallback_fonts: tuple[str, ...]
    size_pt: float
    color: str
    line_height: float


@dataclass(frozen=True, slots=True)
class ParagraphPresentation:
    name: str
    size_pt: float
    color: str
    bold: bool = False
    italic: bool = False
    space_before_pt: float = 0
    space_after_pt: float = 0
    left_indent_mm: float = 0
    keep_with_next: bool = False


@dataclass(frozen=True, slots=True)
class TablePresentation:
    name: str
    border_color: str
    border_width_pt: float
    cell_padding_mm: float


@dataclass(frozen=True, slots=True)
class DocxPresentation:
    presentation_id: str
    page: PagePresentation
    body: TextPresentation
    paragraph_styles: tuple[ParagraphPresentation, ...]
    table: TablePresentation

    def paragraph_style(self, name: str) -> ParagraphPresentation:
        for style in self.paragraph_styles:
            if style.name == name:
                return style
        raise KeyError(name)


PROFESSIONAL_A4 = DocxPresentation(
    presentation_id="professional-a4/v2",
    page=PagePresentation(
        width_mm=210,
        height_mm=297,
        margin_top_mm=18,
        margin_right_mm=18,
        margin_bottom_mm=18,
        margin_left_mm=18,
    ),
    body=TextPresentation(
        font_name="Arial",
        fallback_fonts=("Liberation Sans", "sans-serif"),
        size_pt=10.5,
        color="252B33",
        line_height=1.08,
    ),
    paragraph_styles=(
        ParagraphPresentation("Normal", 10.5, "252B33", space_after_pt=5),
        ParagraphPresentation("Title", 24, "17365D", bold=True, space_after_pt=4, keep_with_next=True),
        ParagraphPresentation("Subtitle", 11.5, "4F81BD", space_after_pt=10, keep_with_next=True),
        ParagraphPresentation("Heading 1", 20, "17365D", bold=True, space_before_pt=12, space_after_pt=6, keep_with_next=True),
        ParagraphPresentation("Heading 2", 12, "365F91", bold=True, space_before_pt=12, space_after_pt=4, keep_with_next=True),
        ParagraphPresentation("Heading 3", 11, "4F81BD", bold=True, space_before_pt=8, space_after_pt=3, keep_with_next=True),
        ParagraphPresentation("Heading 4", 10.5, "365F91", bold=True, space_before_pt=7, space_after_pt=2, keep_with_next=True),
        ParagraphPresentation("Heading 5", 10.5, "365F91", bold=True, italic=True, space_before_pt=7, space_after_pt=2, keep_with_next=True),
        ParagraphPresentation("Heading 6", 10.5, "4F81BD", bold=True, space_before_pt=6, space_after_pt=2, keep_with_next=True),
        ParagraphPresentation("Heading 7", 10.5, "4F81BD", italic=True, space_before_pt=6, space_after_pt=2, keep_with_next=True),
        ParagraphPresentation("Heading 8", 10, "59636E", bold=True, space_before_pt=5, space_after_pt=2, keep_with_next=True),
        ParagraphPresentation("Heading 9", 10, "59636E", italic=True, space_before_pt=5, space_after_pt=2, keep_with_next=True),
        ParagraphPresentation("List Bullet", 10.5, "252B33", space_after_pt=2, left_indent_mm=6.35),
        ParagraphPresentation("List Number", 10.5, "252B33", space_after_pt=2, left_indent_mm=6.35),
        ParagraphPresentation("Quote", 10.5, "59636E", italic=True, space_before_pt=4, space_after_pt=4, left_indent_mm=6.35),
        ParagraphPresentation("Intense Quote", 10.5, "365F91", italic=True, space_before_pt=5, space_after_pt=5, left_indent_mm=6.35),
    ),
    table=TablePresentation(
        name="Table Grid",
        border_color="B9C2CC",
        border_width_pt=0.75,
        cell_padding_mm=1.6,
    ),
)


MANAGED_PARAGRAPH_STYLES = frozenset(style.name for style in PROFESSIONAL_A4.paragraph_styles)
MANAGED_PARAGRAPH_BLOCK_STYLES = MANAGED_PARAGRAPH_STYLES - {"List Bullet", "List Number"}
MANAGED_TABLE_STYLES = frozenset({PROFESSIONAL_A4.table.name})


def resolve_presentation() -> DocxPresentation:
    """Return the one closed DOCX presentation profile supported by V1."""
    return PROFESSIONAL_A4
