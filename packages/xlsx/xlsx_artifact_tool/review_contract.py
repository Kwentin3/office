"""Closed XLSX preview evidence contract for chat orchestration only."""

from __future__ import annotations

import copy
import re
from typing import Any


class ReviewContractError(ValueError):
    """The XLSX review packet violates its closed boundary."""


_ID = re.compile(r"^[A-Za-z][A-Za-z0-9_.-]{0,63}$")
_SHA256 = re.compile(r"^[0-9a-f]{64}$")
_PACKET_FIELDS = {
    "contract_version",
    "kind",
    "interaction",
    "workbook_id",
    "revision",
    "fidelity",
    "limitations",
    "diagnostics",
    "sheets",
}
_DIAGNOSTIC_FIELDS = {"truncated_sheets", "omitted_sheets"}
_TRUNCATION_FIELDS = {"sheet", "rows_shown", "columns_shown", "omitted_cells"}
_OMISSION_FIELDS = {"sheet", "reason"}
_SHEET_FIELDS = {"sheet", "number", "html_file", "html_sha256"}
_OMISSION_REASONS = {"hidden", "very_hidden", "artifact_limit"}


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


def _text(value: Any, path: str, maximum: int) -> str:
    if not isinstance(value, str) or not value.strip() or len(value) > maximum:
        raise ReviewContractError(f"{path} must be non-empty text up to {maximum} characters")
    return value


def _identifier(value: Any, path: str) -> str:
    if not isinstance(value, str) or not _ID.fullmatch(value):
        raise ReviewContractError(f"{path} must be a stable identifier")
    return value


def _digest(value: Any, path: str) -> str:
    if not isinstance(value, str) or not _SHA256.fullmatch(value):
        raise ReviewContractError(f"{path} must be lowercase SHA-256")
    return value


def _array(value: Any, path: str, maximum: int, minimum: int = 0) -> list[Any]:
    if not isinstance(value, list) or not minimum <= len(value) <= maximum:
        raise ReviewContractError(f"{path} cardinality must be between {minimum} and {maximum}")
    return value


def _positive_integer(value: Any, path: str, maximum: int | None = None) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 1 or maximum is not None and value > maximum:
        raise ReviewContractError(f"{path} must be a positive integer")
    return value


def _nonnegative_integer(value: Any, path: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ReviewContractError(f"{path} must be a non-negative integer")
    return value


def _artifact_filename(value: Any, path: str) -> str:
    text = _text(value, path, 100)
    if "/" in text or "\\" in text or text in {".", ".."} or not text.endswith(".html"):
        raise ReviewContractError(f"{path} must be a local .html basename")
    return text


def validate_review_packet(raw: Any) -> dict[str, Any]:
    """Validate and defensively copy a V1 chat-only XLSX review packet."""
    packet = _object(raw, "review")
    _closed(packet, _PACKET_FIELDS, "review")
    if packet["contract_version"] != "1.0":
        raise ReviewContractError("review.contract_version must be 1.0")
    if packet["kind"] != "xlsx_chat_review":
        raise ReviewContractError("review.kind must be xlsx_chat_review")
    if packet["interaction"] != "chat_only":
        raise ReviewContractError("review.interaction must be chat_only")
    _identifier(packet["workbook_id"], "review.workbook_id")
    _digest(packet["revision"], "review.revision")
    if packet["fidelity"] != "structural_preview_not_excel_render":
        raise ReviewContractError("review.fidelity is unsupported")

    limitations = _array(packet["limitations"], "review.limitations", 20, minimum=1)
    for index, limitation in enumerate(limitations):
        _text(limitation, f"review.limitations[{index}]", 500)

    diagnostics = _object(packet["diagnostics"], "review.diagnostics")
    _closed(diagnostics, _DIAGNOSTIC_FIELDS, "review.diagnostics")
    truncated = _array(diagnostics["truncated_sheets"], "review.diagnostics.truncated_sheets", 20)
    for index, raw_issue in enumerate(truncated):
        path = f"review.diagnostics.truncated_sheets[{index}]"
        issue = _object(raw_issue, path)
        _closed(issue, _TRUNCATION_FIELDS, path)
        _text(issue["sheet"], f"{path}.sheet", 31)
        _positive_integer(issue["rows_shown"], f"{path}.rows_shown", 100)
        _positive_integer(issue["columns_shown"], f"{path}.columns_shown", 40)
        _positive_integer(issue["omitted_cells"], f"{path}.omitted_cells")

    omitted = _array(diagnostics["omitted_sheets"], "review.diagnostics.omitted_sheets", 256)
    for index, raw_issue in enumerate(omitted):
        path = f"review.diagnostics.omitted_sheets[{index}]"
        issue = _object(raw_issue, path)
        _closed(issue, _OMISSION_FIELDS, path)
        _text(issue["sheet"], f"{path}.sheet", 31)
        if issue["reason"] not in _OMISSION_REASONS:
            raise ReviewContractError(f"{path}.reason is unsupported")

    sheets = _array(packet["sheets"], "review.sheets", 20, minimum=1)
    sheet_names: set[str] = set()
    filenames: set[str] = set()
    for index, raw_sheet in enumerate(sheets):
        path = f"review.sheets[{index}]"
        sheet = _object(raw_sheet, path)
        _closed(sheet, _SHEET_FIELDS, path)
        name = _text(sheet["sheet"], f"{path}.sheet", 31)
        if name in sheet_names:
            raise ReviewContractError(f"duplicate review sheet: {name}")
        sheet_names.add(name)
        number = _positive_integer(sheet["number"], f"{path}.number", 20)
        if number != index + 1:
            raise ReviewContractError(f"{path}.number must match review order")
        filename = _artifact_filename(sheet["html_file"], f"{path}.html_file")
        if filename in filenames:
            raise ReviewContractError(f"{path}.html_file must be unique")
        filenames.add(filename)
        _digest(sheet["html_sha256"], f"{path}.html_sha256")

    return copy.deepcopy(packet)
