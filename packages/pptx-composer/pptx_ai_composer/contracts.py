"""Closed contracts for artifacts crossing composer domain boundaries.

STICKY INVARIANT: this module accepts semantic intent, never renderer coordinates,
OOXML, python-pptx objects, or executable callbacks. Keep it boring. Adding a field
here creates a durable LLM-facing capability and therefore requires a failing
contract test plus an explicit product decision.
"""

from __future__ import annotations

import copy
import math
import re
from typing import Any


class ContractError(ValueError):
    """A deterministic refusal at an artifact boundary."""


_HEX = re.compile(r"^[0-9A-Fa-f]{6}$")
_ID = re.compile(r"^[A-Za-z][A-Za-z0-9_.-]{0,63}$")
_SUPPORTED_ARCHETYPES = {"cover", "comparison", "chart_with_takeaway", "process", "timeline", "decision_matrix", "kpi_grid"}
_DECK_FIELDS = {"contract_version", "deck_id", "title", "brand", "sources", "assets", "slides"}
_BRAND_FIELDS = {"name", "colors", "fonts"}
_COLOR_FIELDS = {"background", "surface", "text", "muted", "accent", "accent_secondary"}
_FONT_FIELDS = {"heading", "body"}
_SOURCE_FIELDS = {"source_id", "label"}
_ASSET_FIELDS = {"asset_id", "kind", "path", "sha256", "alt_text", "fallback_path", "fallback_sha256"}
_COMMON_SLIDE_FIELDS = {"slide_id", "archetype", "title", "source_ids"}
_ARCHETYPE_FIELDS = {
    "cover": _COMMON_SLIDE_FIELDS | {"subtitle", "asset_id"},
    "comparison": _COMMON_SLIDE_FIELDS | {"left", "right"},
    "chart_with_takeaway": _COMMON_SLIDE_FIELDS | {"chart", "takeaway"},
    "process": _COMMON_SLIDE_FIELDS | {"steps"},
    "timeline": _COMMON_SLIDE_FIELDS | {"milestones"},
    "decision_matrix": _COMMON_SLIDE_FIELDS | {"criteria", "options"},
    "kpi_grid": _COMMON_SLIDE_FIELDS | {"metrics"},
}
_SIDE_FIELDS = {"label", "items"}
_STEP_FIELDS = {"label", "description"}
_MILESTONE_FIELDS = {"period", "label"}
_OPTION_FIELDS = {"label", "ratings"}
_METRIC_FIELDS = {"label", "value", "note"}
_CHART_FIELDS = {"type", "data_source_id", "categories", "series"}
_SERIES_FIELDS = {"name", "values"}


def _object(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ContractError(f"{path} must be an object")
    return value


def _array(value: Any, path: str, *, minimum: int = 0, maximum: int = 100) -> list[Any]:
    if not isinstance(value, list):
        raise ContractError(f"{path} must be an array")
    if not minimum <= len(value) <= maximum:
        raise ContractError(f"{path} cardinality must be between {minimum} and {maximum}")
    return value


def _closed(value: dict[str, Any], allowed: set[str], path: str) -> None:
    unknown = sorted(set(value) - allowed)
    if unknown:
        raise ContractError(f"{path} unknown field: {unknown[0]}")


def _required(value: dict[str, Any], fields: set[str], path: str) -> None:
    missing = sorted(field for field in fields if field not in value)
    if missing:
        raise ContractError(f"{path} missing field: {missing[0]}")


def _text(value: Any, path: str, *, maximum: int = 500) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ContractError(f"{path} must be non-empty text")
    if len(value) > maximum:
        raise ContractError(f"{path} exceeds {maximum} characters")
    return value


def _stable_id(value: Any, path: str) -> str:
    text = _text(value, path, maximum=64)
    if not _ID.fullmatch(text):
        raise ContractError(f"{path} must be a stable identifier")
    return text


def _validate_brand(raw: Any) -> None:
    brand = _object(raw, "brand")
    _closed(brand, _BRAND_FIELDS, "brand")
    _required(brand, _BRAND_FIELDS, "brand")
    _text(brand["name"], "brand.name", maximum=100)
    colors = _object(brand["colors"], "brand.colors")
    _closed(colors, _COLOR_FIELDS, "brand.colors")
    _required(colors, _COLOR_FIELDS, "brand.colors")
    for name, color in colors.items():
        if not isinstance(color, str) or not _HEX.fullmatch(color):
            raise ContractError(f"brand.colors.{name} must be six hex digits without #")
    fonts = _object(brand["fonts"], "brand.fonts")
    _closed(fonts, _FONT_FIELDS, "brand.fonts")
    _required(fonts, _FONT_FIELDS, "brand.fonts")
    for name, font in fonts.items():
        _text(font, f"brand.fonts.{name}", maximum=100)


def _validate_side(raw: Any, path: str) -> None:
    side = _object(raw, path)
    _closed(side, _SIDE_FIELDS, path)
    _required(side, _SIDE_FIELDS, path)
    _text(side["label"], f"{path}.label", maximum=80)
    items = _array(side["items"], f"{path}.items", minimum=1, maximum=5)
    for index, item in enumerate(items):
        _text(item, f"{path}.items[{index}]", maximum=140)


def _validate_chart(raw: Any, path: str) -> None:
    chart = _object(raw, path)
    _closed(chart, _CHART_FIELDS, path)
    _required(chart, _CHART_FIELDS, path)
    if chart["type"] not in {"bar", "line"}:
        raise ContractError(f"{path}.type must be bar or line")
    categories = _array(chart["categories"], f"{path}.categories", minimum=2, maximum=12)
    for index, category in enumerate(categories):
        _text(category, f"{path}.categories[{index}]", maximum=60)
    series = _array(chart["series"], f"{path}.series", minimum=1, maximum=3)
    for index, raw_series in enumerate(series):
        item_path = f"{path}.series[{index}]"
        item = _object(raw_series, item_path)
        _closed(item, _SERIES_FIELDS, item_path)
        _required(item, _SERIES_FIELDS, item_path)
        _text(item["name"], f"{item_path}.name", maximum=80)
        values = _array(item["values"], f"{item_path}.values", minimum=1, maximum=12)
        if len(values) != len(categories):
            raise ContractError(f"{item_path} categories and values must have equal length")
        if any(isinstance(value, bool) or not isinstance(value, (int, float)) for value in values):
            raise ContractError(f"{item_path}.values must be numeric")
        if any(not math.isfinite(float(value)) for value in values):
            raise ContractError(f"{item_path}.values must be finite")


def _validate_records(raw: Any, path: str, fields: set[str], *, minimum: int, maximum: int, limits: dict[str, int]) -> list[dict[str, Any]]:
    records = _array(raw, path, minimum=minimum, maximum=maximum)
    for index, raw_record in enumerate(records):
        item_path = f"{path}[{index}]"
        record = _object(raw_record, item_path)
        _closed(record, fields, item_path)
        _required(record, fields, item_path)
        for field, limit in limits.items():
            _text(record[field], f"{item_path}.{field}", maximum=limit)
    return records


def validate_deck_spec(raw: Any) -> dict[str, Any]:
    """Validate and return a defensive copy of the V1 vertical-slice DeckSpec."""
    deck = _object(raw, "deck")
    _closed(deck, _DECK_FIELDS, "deck")
    _required(deck, _DECK_FIELDS, "deck")
    if deck["contract_version"] != "1.0":
        raise ContractError("contract_version must be 1.0")
    _stable_id(deck["deck_id"], "deck.deck_id")
    _text(deck["title"], "deck.title", maximum=160)
    _validate_brand(deck["brand"])

    sources = _array(deck["sources"], "deck.sources", maximum=50)
    source_ids: set[str] = set()
    source_labels: dict[str, str] = {}
    for index, raw_source in enumerate(sources):
        path = f"deck.sources[{index}]"
        source = _object(raw_source, path)
        _closed(source, _SOURCE_FIELDS, path)
        _required(source, _SOURCE_FIELDS, path)
        source_id = _stable_id(source["source_id"], f"{path}.source_id")
        if source_id in source_ids:
            raise ContractError(f"duplicate source_id: {source_id}")
        source_ids.add(source_id)
        _text(source["label"], f"{path}.label", maximum=160)
        source_labels[source_id] = source["label"]

    assets = _array(deck["assets"], "deck.assets", maximum=50)
    asset_ids: set[str] = set()
    for index, raw_asset in enumerate(assets):
        path = f"deck.assets[{index}]"
        asset = _object(raw_asset, path)
        _closed(asset, _ASSET_FIELDS, path)
        _required(asset, {"asset_id", "kind", "path", "sha256", "alt_text"}, path)
        asset_id = _stable_id(asset["asset_id"], f"{path}.asset_id")
        if asset_id in asset_ids:
            raise ContractError(f"duplicate asset_id: {asset_id}")
        asset_ids.add(asset_id)
        if asset["kind"] not in {"png", "jpeg", "svg"}:
            raise ContractError(f"{path}.kind must be png, jpeg or svg")
        _text(asset["path"], f"{path}.path", maximum=1000)
        if not isinstance(asset["sha256"], str) or not re.fullmatch(r"[0-9a-f]{64}", asset["sha256"]):
            raise ContractError(f"{path}.sha256 must be lowercase SHA-256")
        _text(asset["alt_text"], f"{path}.alt_text", maximum=300)
        if asset["kind"] == "svg":
            _required(asset, {"fallback_path", "fallback_sha256"}, path)
            _text(asset["fallback_path"], f"{path}.fallback_path", maximum=1000)
            if not isinstance(asset["fallback_sha256"], str) or not re.fullmatch(r"[0-9a-f]{64}", asset["fallback_sha256"]):
                raise ContractError(f"{path}.fallback_sha256 must be lowercase SHA-256")
        elif "fallback_path" in asset or "fallback_sha256" in asset:
            raise ContractError(f"{path} fallback fields are only valid for svg")

    slides = _array(deck["slides"], "deck.slides", minimum=1, maximum=20)
    slide_ids: set[str] = set()
    for index, raw_slide in enumerate(slides):
        path = f"deck.slides[{index}]"
        slide = _object(raw_slide, path)
        archetype = slide.get("archetype")
        if archetype not in _SUPPORTED_ARCHETYPES:
            raise ContractError(f"{path}.archetype is unsupported")
        _closed(slide, _ARCHETYPE_FIELDS[archetype], path)
        required = set(_COMMON_SLIDE_FIELDS)
        required |= {
            "cover": {"subtitle"},
            "comparison": {"left", "right"},
            "chart_with_takeaway": {"chart", "takeaway"},
            "process": {"steps"},
            "timeline": {"milestones"},
            "decision_matrix": {"criteria", "options"},
            "kpi_grid": {"metrics"},
        }[archetype]
        _required(slide, required, path)
        slide_id = _stable_id(slide["slide_id"], f"{path}.slide_id")
        if slide_id in slide_ids:
            raise ContractError(f"duplicate slide_id: {slide_id}")
        slide_ids.add(slide_id)
        _text(slide["title"], f"{path}.title", maximum=160)
        references = _array(slide["source_ids"], f"{path}.source_ids", maximum=20)
        for reference in references:
            if reference not in source_ids:
                raise ContractError(f"{path} references unknown source_id: {reference}")
        footer_length = len("Источники: ") + sum(len(source_labels[reference]) for reference in references) + max(0, len(references) - 1) * 2
        if footer_length > 300:
            raise ContractError(f"{path} source footer exceeds 300 characters")
        if archetype == "cover":
            _text(slide["subtitle"], f"{path}.subtitle", maximum=240)
            if "asset_id" in slide and slide["asset_id"] not in asset_ids:
                raise ContractError(f"{path} references unknown asset_id: {slide['asset_id']}")
        elif archetype == "comparison":
            _validate_side(slide["left"], f"{path}.left")
            _validate_side(slide["right"], f"{path}.right")
        elif archetype == "chart_with_takeaway":
            _validate_chart(slide["chart"], f"{path}.chart")
            data_source_id = slide["chart"]["data_source_id"]
            if data_source_id not in source_ids:
                raise ContractError(f"{path}.chart references unknown data_source_id: {data_source_id}")
            if data_source_id not in references:
                raise ContractError(f"{path}.chart data_source_id must also appear in source_ids")
            _text(slide["takeaway"], f"{path}.takeaway", maximum=260)
        elif archetype == "process":
            _validate_records(slide["steps"], f"{path}.steps", _STEP_FIELDS, minimum=2, maximum=6, limits={"label": 60, "description": 140})
        elif archetype == "timeline":
            _validate_records(slide["milestones"], f"{path}.milestones", _MILESTONE_FIELDS, minimum=2, maximum=6, limits={"period": 40, "label": 100})
        elif archetype == "decision_matrix":
            criteria = _array(slide["criteria"], f"{path}.criteria", minimum=2, maximum=4)
            for criterion_index, criterion in enumerate(criteria):
                _text(criterion, f"{path}.criteria[{criterion_index}]", maximum=70)
            options = _array(slide["options"], f"{path}.options", minimum=2, maximum=4)
            for option_index, raw_option in enumerate(options):
                option_path = f"{path}.options[{option_index}]"
                option = _object(raw_option, option_path)
                _closed(option, _OPTION_FIELDS, option_path)
                _required(option, _OPTION_FIELDS, option_path)
                _text(option["label"], f"{option_path}.label", maximum=60)
                ratings = _array(option["ratings"], f"{option_path}.ratings", minimum=1, maximum=4)
                if len(ratings) != len(criteria):
                    raise ContractError(f"{option_path} ratings and criteria must have equal length")
                if any(rating not in {"positive", "neutral", "negative"} for rating in ratings):
                    raise ContractError(f"{option_path}.ratings contains unsupported value")
        else:
            _validate_records(slide["metrics"], f"{path}.metrics", _METRIC_FIELDS, minimum=2, maximum=6, limits={"label": 60, "value": 50, "note": 100})

    return copy.deepcopy(deck)
