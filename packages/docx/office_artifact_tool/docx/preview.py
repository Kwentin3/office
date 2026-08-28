"""Bounded styled-layout HTML proxy for chat-only DOCX review."""

from __future__ import annotations

import ctypes
import hashlib
import html
import json
import os
import shutil
import tempfile
from pathlib import Path
from typing import Any

from ..core.contracts import validate_create_model
from ..core.errors import ArtifactError, refusal
from .presentation import DocxPresentation, resolve_presentation
from .review_contract import validate_review_packet

MAX_PREVIEW_BLOCKS = 100
MAX_LIST_ITEMS_PER_BLOCK = 100
MAX_TABLE_ROWS_PER_BLOCK = 100
MAX_TABLE_COLUMNS_PER_ROW = 20
MAX_TEXT_CHARACTERS_PER_VALUE = 10_000


class PreviewError(ValueError):
    """A deterministic DOCX preview publication refusal."""


def _exchange_directories(source: Path, target: Path) -> None:
    """Atomically exchange two existing directories or fail closed."""
    libc = ctypes.CDLL(None, use_errno=True)
    renameat2 = getattr(libc, "renameat2", None)
    if renameat2 is None:
        raise PreviewError("atomic directory exchange is unavailable")
    renameat2.argtypes = [ctypes.c_int, ctypes.c_char_p, ctypes.c_int, ctypes.c_char_p, ctypes.c_uint]
    renameat2.restype = ctypes.c_int
    if renameat2(-100, os.fsencode(source), -100, os.fsencode(target), 2) != 0:
        error = ctypes.get_errno()
        raise OSError(error, os.strerror(error))


def _cleanup_old_generation(temporary: Path, backup: Path) -> None:
    """Best-effort cleanup after the directory exchange commit point."""
    try:
        os.replace(temporary, backup)
    except Exception:
        try:
            shutil.rmtree(temporary)
        except Exception:
            pass
        return
    try:
        shutil.rmtree(backup)
    except Exception:
        pass


def _revision(model: dict[str, Any], presentation_id: str | None = None) -> str:
    identity = {
        "model": model,
        "presentation_id": presentation_id or resolve_presentation().presentation_id,
    }
    encoded = json.dumps(identity, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def _prepare_review(model: dict[str, Any]) -> tuple[dict[str, Any], DocxPresentation, str]:
    validated = validate_create_model(model)
    presentation = resolve_presentation()
    revision = _revision(validated, presentation.presentation_id)
    return validated, presentation, revision


def prepare_review_identity(model: dict[str, Any]) -> tuple[dict[str, Any], str]:
    """Validate a model and bind its review identity to presentation V1."""
    validated, _presentation, revision = _prepare_review(model)
    return validated, revision


def canonical_review_revision(model: dict[str, Any]) -> str:
    return prepare_review_identity(model)[1]


def _escaped(value: Any) -> str:
    return html.escape(str(value), quote=True)


def _block_ids(model: dict[str, Any]) -> list[str]:
    used = {block["block_id"] for block in model["blocks"] if "block_id" in block}
    identifiers: list[str] = []
    for index, block in enumerate(model["blocks"], start=1):
        identifier = block.get("block_id")
        if identifier is None:
            identifier = f"block-{index:04d}"
            suffix = 2
            while identifier in used:
                identifier = f"block-{index:04d}-{suffix}"
                suffix += 1
            used.add(identifier)
        identifiers.append(identifier)
    return identifiers


def _bounded_text(value: Any) -> tuple[str, int]:
    text = str(value)
    omitted = max(0, len(text) - MAX_TEXT_CHARACTERS_PER_VALUE)
    return text[:MAX_TEXT_CHARACTERS_PER_VALUE], omitted


def _css(presentation: DocxPresentation) -> str:
    page = presentation.page
    body = presentation.body
    font_stack = ",".join((body.font_name, *(f'"{name}"' if " " in name else name for name in body.fallback_fonts)))
    variables = []
    rules = []
    for configured in presentation.paragraph_styles:
        token = _style_token(configured.name)
        css_left_indent = 0 if configured.name in {"List Bullet", "List Number"} else configured.left_indent_mm
        variables.extend((
            f"--{token}-size:{configured.size_pt:g}pt;",
            f"--{token}-color:#{configured.color};",
        ))
        rules.append(
            f".word-style-{token}{{font-size:var(--{token}-size);color:var(--{token}-color);"
            f"font-weight:{700 if configured.bold else 400};font-style:{'italic' if configured.italic else 'normal'};"
            f"margin-top:{configured.space_before_pt:g}pt;margin-bottom:{configured.space_after_pt:g}pt;"
            f"margin-left:{css_left_indent:g}mm;line-height:{body.line_height:g}}}"
        )
    list_bullet = presentation.paragraph_style("List Bullet")
    list_number = presentation.paragraph_style("List Number")
    list_indent = list_bullet.left_indent_mm
    return (
        ":root{"
        f"--page-width:{page.width_mm:g}mm;--page-height:{page.height_mm:g}mm;"
        f"--margin-top:{page.margin_top_mm:g}mm;--margin-right:{page.margin_right_mm:g}mm;"
        f"--margin-bottom:{page.margin_bottom_mm:g}mm;--margin-left:{page.margin_left_mm:g}mm;"
        f"--body-font:{font_stack};--body-size:{body.size_pt:g}pt;"
        f"--body-color:#{body.color};--line-height:{body.line_height:g};"
        f"--table-border-color:#{presentation.table.border_color};"
        f"--table-border-width:{presentation.table.border_width_pt:g}pt;"
        f"--table-cell-padding:{presentation.table.cell_padding_mm:g}mm;"
        f"{''.join(variables)}"
        "}"
        "*{box-sizing:border-box}"
        "html{background:#e9edf2}"
        "body{margin:0;padding:24px;font-family:var(--body-font);font-size:var(--body-size);"
        "line-height:var(--line-height);color:var(--body-color)}"
        ".document-page{width:var(--page-width);min-height:var(--page-height);margin:0 auto;"
        "padding:var(--margin-top) var(--margin-right) var(--margin-bottom) var(--margin-left);"
        "background:#fff;box-shadow:0 4px 24px rgba(31,42,55,.16)}"
        "h1,h2,h3,h4,h5,h6,p,ul,ol{margin-right:0;padding-top:0;padding-bottom:0}"
        f"{''.join(rules)}"
        f"ul.word-style,ol.word-style{{padding-left:{list_indent:g}mm;margin-top:0;margin-bottom:0}}"
        f".word-style-list-bullet li{{margin-bottom:{list_bullet.space_after_pt:g}pt}}"
        f".word-style-list-number li{{margin-bottom:{list_number.space_after_pt:g}pt}}"
        ".word-table{width:100%;table-layout:fixed;border-collapse:collapse;margin:0}"
        ".word-table td{border:var(--table-border-width) solid var(--table-border-color);"
        "padding:var(--table-cell-padding);vertical-align:top}"
        '.word-table .word-style-normal:empty::after{content:"\\00a0"}'
        f"@page{{size:{page.width_mm:g}mm {page.height_mm:g}mm;"
        f"margin:{page.margin_top_mm:g}mm {page.margin_right_mm:g}mm "
        f"{page.margin_bottom_mm:g}mm {page.margin_left_mm:g}mm}}"
    )


def _style_token(name: str) -> str:
    return name.casefold().replace(" ", "-")


def _render_html(model: dict[str, Any], document_id: str, presentation: DocxPresentation) -> tuple[bytes, dict[str, Any]]:
    blocks: list[str] = []
    truncations: list[dict[str, Any]] = []
    identifiers = _block_ids(model)
    selected = model["blocks"][:MAX_PREVIEW_BLOCKS]
    for block, block_id in zip(selected, identifiers, strict=False):
        kind = block["type"]
        content = ""
        if kind in {"heading", "paragraph"}:
            text, omitted = _bounded_text(block["text"])
            if omitted:
                truncations.append({"block_id": block_id, "content": "text_characters", "omitted": omitted})
            if kind == "heading":
                level = block.get("level", 1)
                token = _style_token(f"Heading {level}")
                if level <= 6:
                    content = f'<h{level} class="word-style word-style-{token}">{_escaped(text)}</h{level}>'
                else:
                    content = (
                        f'<p role="heading" aria-level="{level}" '
                        f'class="word-style word-style-{token}">{_escaped(text)}</p>'
                    )
            else:
                token = _style_token(block.get("style", "Normal"))
                content = f'<p class="word-style word-style-{token}">{_escaped(text)}</p>'
        elif kind in {"numbered_list", "bulleted_list"}:
            tag = "ol" if kind == "numbered_list" else "ul"
            item_texts: list[str] = []
            omitted_characters = 0
            for item in block["items"][:MAX_LIST_ITEMS_PER_BLOCK]:
                text, omitted = _bounded_text(item)
                omitted_characters += omitted
                item_texts.append(f"<li>{_escaped(text)}</li>")
            if omitted_characters:
                truncations.append(
                    {"block_id": block_id, "content": "text_characters", "omitted": omitted_characters}
                )
            omitted_items = max(0, len(block["items"]) - MAX_LIST_ITEMS_PER_BLOCK)
            if omitted_items:
                truncations.append({"block_id": block_id, "content": "list_items", "omitted": omitted_items})
            token = _style_token("List Number" if kind == "numbered_list" else "List Bullet")
            content = f'<{tag} class="word-style word-style-{token}">{"".join(item_texts)}</{tag}>'
        elif kind == "table":
            row_html: list[str] = []
            omitted_characters = 0
            retained_rows = block["rows"][:MAX_TABLE_ROWS_PER_BLOCK]
            retained_width = min(max(map(len, retained_rows)), MAX_TABLE_COLUMNS_PER_ROW)
            for row in retained_rows:
                cells: list[str] = []
                retained_values = row[:MAX_TABLE_COLUMNS_PER_ROW]
                for value in (*retained_values, *("" for _ in range(retained_width - len(retained_values)))):
                    text, omitted = _bounded_text(value)
                    omitted_characters += omitted
                    cells.append(
                        '<td><p class="word-style word-style-normal">'
                        f"{_escaped(text)}</p></td>"
                    )
                row_html.append("<tr>" + "".join(cells) + "</tr>")
            omitted_rows = max(0, len(block["rows"]) - MAX_TABLE_ROWS_PER_BLOCK)
            if omitted_rows:
                truncations.append({"block_id": block_id, "content": "table_rows", "omitted": omitted_rows})
            omitted_columns = sum(max(0, len(row) - MAX_TABLE_COLUMNS_PER_ROW) for row in retained_rows)
            if omitted_columns:
                truncations.append(
                    {"block_id": block_id, "content": "table_columns", "omitted": omitted_columns}
                )
            if omitted_characters:
                truncations.append(
                    {"block_id": block_id, "content": "text_characters", "omitted": omitted_characters}
                )
            token = _style_token(block.get("style", "Table Grid"))
            content = f'<table class="word-table word-{token}"><tbody>{"".join(row_html)}</tbody></table>'
        blocks.append(f'<section data-block-id="{_escaped(block_id)}">{content}</section>')

    raw_title = model.get("metadata", {}).get("title", "DOCX styled-layout preview")
    title, omitted_title = _bounded_text(raw_title)
    if omitted_title:
        truncations.insert(
            0,
            {"block_id": document_id, "content": "text_characters", "omitted": omitted_title},
        )
    document = (
        f'<!doctype html><html><head><meta charset="utf-8"><title>{_escaped(title)}</title>'
        f"<style>{_css(presentation)}</style></head>"
        f'<body><main class="document-page">{"".join(blocks)}</main></body></html>\n'
    ).encode()
    total = len(model["blocks"])
    rendered = len(selected)
    diagnostics = {
        "truncated": bool(total - rendered or truncations),
        "total_blocks": total,
        "rendered_blocks": rendered,
        "omitted_blocks": total - rendered,
        "limits": {
            "max_blocks": MAX_PREVIEW_BLOCKS,
            "max_list_items_per_block": MAX_LIST_ITEMS_PER_BLOCK,
            "max_table_rows_per_block": MAX_TABLE_ROWS_PER_BLOCK,
            "max_table_columns_per_row": MAX_TABLE_COLUMNS_PER_ROW,
            "max_text_characters_per_value": MAX_TEXT_CHARACTERS_PER_VALUE,
        },
        "truncations": truncations,
    }
    return document, diagnostics


def _publish(
    output: Path,
    html_bytes: bytes,
    *,
    document_id: str,
    presentation_id: str,
    revision: str,
    diagnostics: dict[str, Any],
) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = Path(tempfile.mkdtemp(prefix=f".{output.name}.", dir=output.parent))
    backup = output.with_name(f".{output.name}.old")
    try:
        (temporary / "document.html").write_bytes(html_bytes)
        review = validate_review_packet(
            {
                "contract_version": "1.1",
                "kind": "docx_chat_review",
                "interaction": "chat_only",
                "document_id": document_id,
                "presentation_id": presentation_id,
                "revision": revision,
                "fidelity": "styled_layout_proxy_not_word_render",
                "limitations": [
                    "page geometry, managed styles, colors, and spacing mirror the DOCX presentation contract",
                    "line wrapping and pagination remain browser approximations",
                    "fonts, headers, footers, and fields are not Word-rendered",
                    "only a bounded subset of create-model content is displayed",
                ],
                "diagnostics": diagnostics,
                "artifact": {
                    "file": "document.html",
                    "sha256": hashlib.sha256(html_bytes).hexdigest(),
                },
            }
        )
        (temporary / "review.json").write_text(
            json.dumps(review, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        if backup.exists():
            shutil.rmtree(backup)
        if output.exists():
            _exchange_directories(temporary, output)
            _cleanup_old_generation(temporary, backup)
        else:
            os.replace(temporary, output)
    except Exception:
        shutil.rmtree(temporary, ignore_errors=True)
        raise


def render_docx_preview(model: dict[str, Any], output_dir: str | Path) -> dict[str, Any]:
    """Validate a create model and identify its chat-review revision."""
    try:
        validated, presentation, revision = _prepare_review(model)
    except ArtifactError as exc:
        return refusal(exc.reason, exc.details)
    raw_output = Path(output_dir).expanduser().absolute()
    raw_backup = raw_output.with_name(f".{raw_output.name}.old")
    if raw_output.is_symlink() or raw_backup.is_symlink():
        raise PreviewError("preview output and backup must not be symlinks")
    output = raw_output.resolve(strict=False)
    if output.exists() and not output.is_dir():
        raise PreviewError("existing preview output must be a directory")
    if raw_backup.exists() and not raw_backup.is_dir():
        raise PreviewError("preview backup path must be a directory")
    document_id = validated.get("document_id", "document")
    html_bytes, diagnostics = _render_html(validated, document_id, presentation)
    _publish(
        output,
        html_bytes,
        document_id=document_id,
        presentation_id=presentation.presentation_id,
        revision=revision,
        diagnostics=diagnostics,
    )
    return {
        "status": "previewed",
        "document_id": document_id,
        "interaction": "chat_only",
        "revision": revision,
        "review_contract": str((output / "review.json").resolve()),
        "display_artifacts": [str((output / "document.html").resolve())],
        "output": str(output.resolve()),
    }
