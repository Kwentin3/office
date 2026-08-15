"""Closed contract from the PPTX preview domain to chat orchestration.

This artifact contains review evidence only. It never carries prompts, renderer
coordinates, Office objects, callbacks, or mutation instructions.
"""

from __future__ import annotations

import copy
import re
from typing import Any


class ReviewContractError(ValueError):
    """The preview review packet violates its closed boundary."""


_ID = re.compile(r"^[A-Za-z][A-Za-z0-9_.-]{0,63}$")
_SHA256 = re.compile(r"^[0-9a-f]{64}$")
_PACKET_FIELDS = {
    "contract_version",
    "kind",
    "interaction",
    "deck_id",
    "revision",
    "fidelity",
    "limitations",
    "diagnostics",
    "slides",
}
_SLIDE_FIELDS = {"slide_id", "number", "png_file", "png_sha256", "svg_file", "svg_sha256"}
_DIAGNOSTIC_FIELDS = {"text_overflow"}
_OVERFLOW_FIELDS = {"slide_id", "node_id", "role", "required_lines", "available_lines"}


def _object(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ReviewContractError(f"{path} must be an object")
    return value


def _closed(value: dict[str, Any], fields: set[str], path: str) -> None:
    unknown = sorted(set(value) - fields)
    missing = sorted(fields - set(value))
    if unknown:
        raise ReviewContractError(f"{path} unknown field: {unknown[0]}")
    if missing:
        raise ReviewContractError(f"{path} missing field: {missing[0]}")


def _identifier(value: Any, path: str) -> str:
    if not isinstance(value, str) or not _ID.fullmatch(value):
        raise ReviewContractError(f"{path} must be a stable identifier")
    return value


def _digest(value: Any, path: str) -> str:
    if not isinstance(value, str) or not _SHA256.fullmatch(value):
        raise ReviewContractError(f"{path} must be lowercase SHA-256")
    return value


def _text(value: Any, path: str, maximum: int) -> str:
    if not isinstance(value, str) or not value.strip() or len(value) > maximum:
        raise ReviewContractError(f"{path} must be non-empty text up to {maximum} characters")
    return value


def _array(value: Any, path: str, maximum: int, minimum: int = 0) -> list[Any]:
    if not isinstance(value, list) or not minimum <= len(value) <= maximum:
        raise ReviewContractError(f"{path} cardinality must be between {minimum} and {maximum}")
    return value


def _artifact_filename(value: Any, path: str, suffix: str) -> str:
    text = _text(value, path, 100)
    if "/" in text or "\\" in text or text in {".", ".."} or not text.endswith(suffix):
        raise ReviewContractError(f"{path} must be a local {suffix} basename")
    return text


def validate_review_packet(raw: Any) -> dict[str, Any]:
    """Validate and defensively copy a V1 chat-only review packet."""
    packet = _object(raw, "review")
    _closed(packet, _PACKET_FIELDS, "review")
    if packet["contract_version"] != "1.0":
        raise ReviewContractError("review.contract_version must be 1.0")
    if packet["kind"] != "pptx_chat_review":
        raise ReviewContractError("review.kind must be pptx_chat_review")
    if packet["interaction"] != "chat_only":
        raise ReviewContractError("review.interaction must be chat_only")
    _identifier(packet["deck_id"], "review.deck_id")
    _digest(packet["revision"], "review.revision")
    if packet["fidelity"] != "structural_preview_not_powerpoint_render":
        raise ReviewContractError("review.fidelity is unsupported")

    limitations = _array(packet["limitations"], "review.limitations", 20)
    for index, limitation in enumerate(limitations):
        _text(limitation, f"review.limitations[{index}]", 500)

    diagnostics = _object(packet["diagnostics"], "review.diagnostics")
    _closed(diagnostics, _DIAGNOSTIC_FIELDS, "review.diagnostics")
    overflow = _array(diagnostics["text_overflow"], "review.diagnostics.text_overflow", 1600)
    for index, raw_issue in enumerate(overflow):
        path = f"review.diagnostics.text_overflow[{index}]"
        issue = _object(raw_issue, path)
        _closed(issue, _OVERFLOW_FIELDS, path)
        _identifier(issue["slide_id"], f"{path}.slide_id")
        _identifier(issue["node_id"], f"{path}.node_id")
        _text(issue["role"], f"{path}.role", 80)
        for field in ("required_lines", "available_lines"):
            if isinstance(issue[field], bool) or not isinstance(issue[field], int) or issue[field] < 1:
                raise ReviewContractError(f"{path}.{field} must be a positive integer")

    slides = _array(packet["slides"], "review.slides", 20, minimum=1)
    slide_ids: set[str] = set()
    filenames: set[str] = set()
    for index, raw_slide in enumerate(slides):
        path = f"review.slides[{index}]"
        slide = _object(raw_slide, path)
        _closed(slide, _SLIDE_FIELDS, path)
        slide_id = _identifier(slide["slide_id"], f"{path}.slide_id")
        if slide_id in slide_ids:
            raise ReviewContractError(f"duplicate review slide_id: {slide_id}")
        slide_ids.add(slide_id)
        number = slide["number"]
        if isinstance(number, bool) or not isinstance(number, int) or number < 1:
            raise ReviewContractError(f"{path}.number must be a positive integer")
        if number != index + 1:
            raise ReviewContractError(f"{path}.number must match review order")
        png_name = _artifact_filename(slide["png_file"], f"{path}.png_file", ".png")
        svg_name = _artifact_filename(slide["svg_file"], f"{path}.svg_file", ".svg")
        if png_name in filenames or svg_name in filenames:
            raise ReviewContractError(f"{path} artifact filename must be unique")
        filenames.update({png_name, svg_name})
        _digest(slide["png_sha256"], f"{path}.png_sha256")
        _digest(slide["svg_sha256"], f"{path}.svg_sha256")

    return copy.deepcopy(packet)
