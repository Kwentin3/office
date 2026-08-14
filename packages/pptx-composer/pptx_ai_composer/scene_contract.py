"""Closed contract between composition recipes and rendering backends.

STICKY DOMAIN BOUNDARY: SceneSpec is runtime-facing, not LLM-facing. It is the
smallest shared rendering language required by the SVG/PNG preview and native
PPTX backends. It admits finite drawing primitives and absolute canvas boxes
only after semantic DeckSpec has been validated and compiled by trusted code.
No OOXML, callbacks, backend objects, or recursive arbitrary groups belong here.
"""

from __future__ import annotations

import copy
import math
import re
from typing import Any


class SceneContractError(ValueError):
    """A deterministic refusal at the compiler/backend boundary."""


_HEX = re.compile(r"^[0-9A-Fa-f]{6}$")
_ID = re.compile(r"^[A-Za-z][A-Za-z0-9_.-]{0,63}$")
_SCENE_FIELDS = {"scene_version", "deck_id", "canvas", "slides"}
_CANVAS_FIELDS = {"width", "height"}
_SLIDE_FIELDS = {"slide_id", "archetype", "variant", "nodes"}
_BOX_FIELDS = {"x", "y", "w", "h"}
_TEXT_FIELDS = {"node_id", "kind", "role", "box", "text", "style"}
_TEXT_STYLE_FIELDS = {"font", "size", "color", "bold", "align", "valign"}
_SHAPE_FIELDS = {"node_id", "kind", "role", "box", "shape", "fill", "stroke", "radius", "opacity"}
_IMAGE_FIELDS = {"node_id", "kind", "role", "box", "asset_id", "fit"}
_CHART_FIELDS = {"node_id", "kind", "role", "box", "chart"}
_NODE_FIELDS = {
    "text": _TEXT_FIELDS,
    "shape": _SHAPE_FIELDS,
    "image": _IMAGE_FIELDS,
    "chart": _CHART_FIELDS,
}


def _object(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise SceneContractError(f"{path} must be an object")
    return value


def _closed(value: dict[str, Any], allowed: set[str], path: str) -> None:
    unknown = sorted(set(value) - allowed)
    if unknown:
        raise SceneContractError(f"{path} unknown field: {unknown[0]}")


def _required(value: dict[str, Any], fields: set[str], path: str) -> None:
    missing = sorted(fields - set(value))
    if missing:
        raise SceneContractError(f"{path} missing field: {missing[0]}")


def _array(value: Any, path: str, *, minimum: int = 0, maximum: int = 100) -> list[Any]:
    if not isinstance(value, list):
        raise SceneContractError(f"{path} must be an array")
    if not minimum <= len(value) <= maximum:
        raise SceneContractError(f"{path} cardinality must be between {minimum} and {maximum}")
    return value


def _text(value: Any, path: str, maximum: int) -> str:
    if not isinstance(value, str) or not value.strip() or len(value) > maximum:
        raise SceneContractError(f"{path} must be non-empty text up to {maximum} characters")
    return value


def _identifier(value: Any, path: str) -> str:
    text = _text(value, path, 64)
    if not _ID.fullmatch(text):
        raise SceneContractError(f"{path} must be a stable identifier")
    return text


def _number(value: Any, path: str, *, minimum: float, maximum: float) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not minimum <= value <= maximum:
        raise SceneContractError(f"{path} must be between {minimum} and {maximum}")
    return float(value)


def _color(value: Any, path: str) -> str:
    if not isinstance(value, str) or not _HEX.fullmatch(value):
        raise SceneContractError(f"{path} must be six hex digits")
    return value.upper()


def _validate_box(raw: Any, path: str, width: float, height: float) -> None:
    box = _object(raw, path)
    _closed(box, _BOX_FIELDS, path)
    _required(box, _BOX_FIELDS, path)
    x = _number(box["x"], f"{path}.x", minimum=0, maximum=width)
    y = _number(box["y"], f"{path}.y", minimum=0, maximum=height)
    w = _number(box["w"], f"{path}.w", minimum=0.01, maximum=width)
    h = _number(box["h"], f"{path}.h", minimum=0.01, maximum=height)
    if x + w > width + 1e-6 or y + h > height + 1e-6:
        raise SceneContractError(f"{path} is outside canvas")


def _validate_node(raw: Any, path: str, width: float, height: float) -> str:
    node = _object(raw, path)
    kind = node.get("kind")
    if kind not in _NODE_FIELDS:
        raise SceneContractError(f"{path}.kind is unsupported")
    allowed = _NODE_FIELDS[kind]
    required = {
        "text": _TEXT_FIELDS,
        "shape": {"node_id", "kind", "role", "box", "shape", "fill"},
        "image": _IMAGE_FIELDS,
        "chart": _CHART_FIELDS,
    }[kind]
    _closed(node, allowed, path)
    _required(node, required, path)
    node_id = _identifier(node["node_id"], f"{path}.node_id")
    _text(node["role"], f"{path}.role", 80)
    _validate_box(node["box"], f"{path}.box", width, height)
    if kind == "text":
        _text(node["text"], f"{path}.text", 2000)
        style = _object(node["style"], f"{path}.style")
        _closed(style, _TEXT_STYLE_FIELDS, f"{path}.style")
        _required(style, _TEXT_STYLE_FIELDS, f"{path}.style")
        _text(style["font"], f"{path}.style.font", 100)
        _number(style["size"], f"{path}.style.size", minimum=6, maximum=96)
        _color(style["color"], f"{path}.style.color")
        if not isinstance(style["bold"], bool):
            raise SceneContractError(f"{path}.style.bold must be boolean")
        if style["align"] not in {"left", "center", "right"}:
            raise SceneContractError(f"{path}.style.align is unsupported")
        if style["valign"] not in {"top", "middle", "bottom"}:
            raise SceneContractError(f"{path}.style.valign is unsupported")
    elif kind == "shape":
        if node["shape"] not in {"rect", "round_rect", "ellipse", "line"}:
            raise SceneContractError(f"{path}.shape is unsupported")
        _color(node["fill"], f"{path}.fill")
        if "stroke" in node:
            _color(node["stroke"], f"{path}.stroke")
        if "radius" in node:
            _number(node["radius"], f"{path}.radius", minimum=0, maximum=200)
        if "opacity" in node:
            _number(node["opacity"], f"{path}.opacity", minimum=0, maximum=1)
    elif kind == "image":
        _identifier(node["asset_id"], f"{path}.asset_id")
        if node["fit"] not in {"contain", "cover"}:
            raise SceneContractError(f"{path}.fit is unsupported")
    else:
        chart = _object(node["chart"], f"{path}.chart")
        if set(chart) != {"type", "data_source_id", "categories", "series"}:
            raise SceneContractError(f"{path}.chart has invalid fields")
        _identifier(chart["data_source_id"], f"{path}.chart.data_source_id")
        if chart["type"] not in {"bar", "line"}:
            raise SceneContractError(f"{path}.chart.type is unsupported")
        categories = _array(chart["categories"], f"{path}.chart.categories", minimum=2, maximum=12)
        series = _array(chart["series"], f"{path}.chart.series", minimum=1, maximum=3)
        for index, category in enumerate(categories):
            _text(category, f"{path}.chart.categories[{index}]", 60)
        for index, item in enumerate(series):
            item_path = f"{path}.chart.series[{index}]"
            item = _object(item, item_path)
            if set(item) != {"name", "values"}:
                raise SceneContractError(f"{item_path} has invalid fields")
            _text(item["name"], f"{item_path}.name", 80)
            values = _array(item["values"], f"{item_path}.values", minimum=len(categories), maximum=len(categories))
            if any(isinstance(value, bool) or not isinstance(value, (int, float)) for value in values):
                raise SceneContractError(f"{item_path}.values must be numeric")
            if any(not math.isfinite(float(value)) for value in values):
                raise SceneContractError(f"{item_path}.values must be finite")
    return node_id


def validate_scene_spec(raw: Any) -> dict[str, Any]:
    scene = _object(raw, "scene")
    _closed(scene, _SCENE_FIELDS, "scene")
    _required(scene, _SCENE_FIELDS, "scene")
    if scene["scene_version"] != "1.0":
        raise SceneContractError("scene_version must be 1.0")
    _identifier(scene["deck_id"], "scene.deck_id")
    canvas = _object(scene["canvas"], "scene.canvas")
    _closed(canvas, _CANVAS_FIELDS, "scene.canvas")
    _required(canvas, _CANVAS_FIELDS, "scene.canvas")
    width = _number(canvas["width"], "scene.canvas.width", minimum=320, maximum=4096)
    height = _number(canvas["height"], "scene.canvas.height", minimum=180, maximum=4096)
    slides = _array(scene["slides"], "scene.slides", minimum=1, maximum=20)
    slide_ids: set[str] = set()
    for slide_index, raw_slide in enumerate(slides):
        path = f"scene.slides[{slide_index}]"
        slide = _object(raw_slide, path)
        _closed(slide, _SLIDE_FIELDS, path)
        _required(slide, _SLIDE_FIELDS, path)
        slide_id = _identifier(slide["slide_id"], f"{path}.slide_id")
        if slide_id in slide_ids:
            raise SceneContractError(f"duplicate slide_id: {slide_id}")
        slide_ids.add(slide_id)
        _identifier(slide["archetype"], f"{path}.archetype")
        _identifier(slide["variant"], f"{path}.variant")
        nodes = _array(slide["nodes"], f"{path}.nodes", minimum=1, maximum=80)
        node_ids: set[str] = set()
        for node_index, raw_node in enumerate(nodes):
            node_id = _validate_node(raw_node, f"{path}.nodes[{node_index}]", width, height)
            if node_id in node_ids:
                raise SceneContractError(f"duplicate node_id: {node_id}")
            node_ids.add(node_id)
    return copy.deepcopy(scene)
