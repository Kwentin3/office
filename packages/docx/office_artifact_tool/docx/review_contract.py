"""Closed contract for DOCX styled-layout preview evidence returned to chat."""

from __future__ import annotations

import copy
import re
from typing import Any


class ReviewContractError(ValueError):
    """The DOCX review packet violates its closed boundary."""


_ID = re.compile(r"^[A-Za-z][A-Za-z0-9_.-]{0,63}$")
_SHA256 = re.compile(r"^[0-9a-f]{64}$")
_PACKET_FIELDS = {
    "contract_version",
    "kind",
    "interaction",
    "document_id",
    "presentation_id",
    "revision",
    "fidelity",
    "limitations",
    "diagnostics",
    "artifact",
}
_ARTIFACT_FIELDS = {"file", "sha256"}
_DIAGNOSTIC_FIELDS = {
    "truncated",
    "total_blocks",
    "rendered_blocks",
    "omitted_blocks",
    "limits",
    "truncations",
}
_LIMIT_FIELDS = {
    "max_blocks",
    "max_list_items_per_block",
    "max_table_rows_per_block",
    "max_table_columns_per_row",
    "max_text_characters_per_value",
}
_TRUNCATION_FIELDS = {"block_id", "content", "omitted"}
_CONTENT_KINDS = {"text_characters", "list_items", "table_rows", "table_columns"}


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


def _integer(value: Any, path: str, *, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        raise ReviewContractError(f"{path} must be an integer of at least {minimum}")
    return value


def validate_review_packet(raw: Any) -> dict[str, Any]:
    """Validate and defensively copy a recursively closed V1 review packet."""
    packet = _object(raw, "review")
    _closed(packet, _PACKET_FIELDS, "review")
    if packet["contract_version"] != "1.1":
        raise ReviewContractError("review.contract_version must be 1.1")
    if packet["kind"] != "docx_chat_review":
        raise ReviewContractError("review.kind must be docx_chat_review")
    if packet["interaction"] != "chat_only":
        raise ReviewContractError("review.interaction must be chat_only")
    _identifier(packet["document_id"], "review.document_id")
    if packet["presentation_id"] != "professional-a4/v2":
        raise ReviewContractError("review.presentation_id is unsupported")
    _digest(packet["revision"], "review.revision")
    if packet["fidelity"] != "styled_layout_proxy_not_word_render":
        raise ReviewContractError("review.fidelity is unsupported")

    limitations = packet["limitations"]
    if not isinstance(limitations, list) or len(limitations) > 20:
        raise ReviewContractError("review.limitations must be an array of at most 20 items")
    if not all(isinstance(item, str) and item and len(item) <= 500 for item in limitations):
        raise ReviewContractError("review.limitations items must be non-empty text")

    artifact = _object(packet["artifact"], "review.artifact")
    _closed(artifact, _ARTIFACT_FIELDS, "review.artifact")
    if artifact["file"] != "document.html":
        raise ReviewContractError("review.artifact.file must be document.html")
    _digest(artifact["sha256"], "review.artifact.sha256")

    diagnostics = _object(packet["diagnostics"], "review.diagnostics")
    _closed(diagnostics, _DIAGNOSTIC_FIELDS, "review.diagnostics")
    if not isinstance(diagnostics["truncated"], bool):
        raise ReviewContractError("review.diagnostics.truncated must be boolean")
    total = _integer(diagnostics["total_blocks"], "review.diagnostics.total_blocks")
    rendered = _integer(diagnostics["rendered_blocks"], "review.diagnostics.rendered_blocks")
    omitted = _integer(diagnostics["omitted_blocks"], "review.diagnostics.omitted_blocks")
    if rendered + omitted != total:
        raise ReviewContractError("review.diagnostics block counts are inconsistent")

    limits = _object(diagnostics["limits"], "review.diagnostics.limits")
    _closed(limits, _LIMIT_FIELDS, "review.diagnostics.limits")
    for field in _LIMIT_FIELDS:
        _integer(limits[field], f"review.diagnostics.limits.{field}", minimum=1)

    truncations = diagnostics["truncations"]
    if not isinstance(truncations, list) or len(truncations) > 500:
        raise ReviewContractError("review.diagnostics.truncations must be a bounded array")
    for index, raw_issue in enumerate(truncations):
        path = f"review.diagnostics.truncations[{index}]"
        issue = _object(raw_issue, path)
        _closed(issue, _TRUNCATION_FIELDS, path)
        _identifier(issue["block_id"], f"{path}.block_id")
        if issue["content"] not in _CONTENT_KINDS:
            raise ReviewContractError(f"{path}.content is unsupported")
        _integer(issue["omitted"], f"{path}.omitted", minimum=1)
    if diagnostics["truncated"] != bool(omitted or truncations):
        raise ReviewContractError("review.diagnostics.truncated is inconsistent")

    return copy.deepcopy(packet)
