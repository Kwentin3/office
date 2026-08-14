"""Trusted compiler from semantic DeckSpec to bounded SceneSpec.

STICKY DOMAIN BOUNDARY: recipes own all geometry. The planner selects semantic
archetypes and, optionally, named visual variants. It never supplies a SceneSpec.
Backends render validated nodes and do not reinterpret narrative intent.
"""

from __future__ import annotations

import copy
from typing import Any, Iterable

from .contracts import validate_deck_spec
from .library import get_catalog
from .scene_contract import validate_scene_spec


class CompileError(ValueError):
    """The semantic deck cannot be compiled by the managed recipe library."""


_WIDTH, _HEIGHT = 1280, 720
_RECIPE_COMPILERS = {
    "cover": "_compile_cover",
    "comparison": "_compile_comparison",
    "chart_with_takeaway": "_compile_chart",
    "process": "_compile_process",
    "timeline": "_compile_timeline",
    "decision_matrix": "_compile_decision_matrix",
    "kpi_grid": "_compile_kpi_grid",
}


def _supported_variants() -> dict[str, set[str]]:
    catalog = get_catalog()
    return {name: set(entry["variants"]) for name, entry in catalog["archetypes"].items()}


def _box(x: float, y: float, w: float, h: float) -> dict[str, float]:
    return {"x": x, "y": y, "w": w, "h": h}


def _text(node_id: str, role: str, box: dict[str, float], value: str, brand: dict[str, Any], *, size: int, bold: bool = False, color: str | None = None, align: str = "left", valign: str = "top") -> dict[str, Any]:
    return {
        "node_id": node_id,
        "kind": "text",
        "role": role,
        "box": box,
        "text": value,
        "style": {
            "font": brand["fonts"]["heading" if bold else "body"],
            "size": size,
            "color": color or brand["colors"]["text"],
            "bold": bold,
            "align": align,
            "valign": valign,
        },
    }


def _shape(node_id: str, role: str, box: dict[str, float], shape: str, fill: str, *, stroke: str | None = None, radius: float | None = None, opacity: float | None = None) -> dict[str, Any]:
    node: dict[str, Any] = {"node_id": node_id, "kind": "shape", "role": role, "box": box, "shape": shape, "fill": fill}
    if stroke is not None:
        node["stroke"] = stroke
    if radius is not None:
        node["radius"] = radius
    if opacity is not None:
        node["opacity"] = opacity
    return node


def _source_map(deck: dict[str, Any]) -> dict[str, str]:
    return {source["source_id"]: source["label"] for source in deck["sources"]}


def _footer(spec: dict[str, Any], sources: dict[str, str], brand: dict[str, Any]) -> list[dict[str, Any]]:
    labels = [sources[source_id] for source_id in spec["source_ids"]]
    if not labels:
        return []
    prefix = "Источник: " if len(labels) == 1 else "Источники: "
    return [_text("source-footer", "source.footer", _box(70, 674, 1140, 28), prefix + "; ".join(labels), brand, size=9, color=brand["colors"]["muted"])]


def _compile_cover(spec: dict[str, Any], deck: dict[str, Any], variant: str) -> list[dict[str, Any]]:
    brand = deck["brand"]; colors = brand["colors"]
    background = _shape("background", "background", _box(0, 0, _WIDTH, _HEIGHT), "rect", colors["background"])
    accent = _shape("accent-stripe", "cover.accent", _box(0, 0, 22, 720), "rect", colors["accent"])
    panel = _shape("cover-panel", "cover.panel", _box(790, 0, 490, 720), "rect", colors["accent_secondary"])
    motif = _shape("cover-motif", "cover.motif", _box(885, 110, 280, 280), "ellipse", colors["surface"], opacity=0.86)
    title = _text("title", "title", _box(72, 110, 650, 190), spec["title"], brand, size=42, bold=True, valign="middle")
    subtitle = _text("subtitle", "subtitle", _box(75, 330, 620, 130), spec["subtitle"], brand, size=22, color=colors["muted"])
    visual = [panel, motif]
    if "asset_id" in spec:
        visual = [{"node_id": "cover-image", "kind": "image", "role": "cover.image", "box": _box(790, 0, 490, 720), "asset_id": spec["asset_id"], "fit": "cover"}]
    if variant == "dark":
        background["fill"] = colors["text"]
        title["style"]["color"] = colors["surface"]
        subtitle["style"]["color"] = colors["background"]
    nodes = [background, accent, *visual, title, subtitle]
    nodes.extend(_footer(spec, _source_map(deck), brand))
    return nodes


def _compile_comparison(spec: dict[str, Any], deck: dict[str, Any], variant: str) -> list[dict[str, Any]]:
    brand = deck["brand"]; colors = brand["colors"]
    card_y, card_h = (145, 470) if variant == "balanced" else (155, 405)
    nodes = [
        _shape("background", "background", _box(0, 0, _WIDTH, _HEIGHT), "rect", colors["background"]),
        _text("title", "title", _box(70, 45, 1140, 80), spec["title"], brand, size=34, bold=True),
    ]
    for side_name, side, x, accent in (("left", spec["left"], 70, colors["accent"]), ("right", spec["right"], 660, colors["accent_secondary"])):
        nodes.extend([
            _shape(f"{side_name}-card", f"comparison.{side_name}.card", _box(x, card_y, 550, card_h), "round_rect", colors["surface"], stroke="D9DEE3", radius=22),
            _shape(f"{side_name}-badge", f"comparison.{side_name}.badge", _box(x + 25, card_y + 32, 46, 46), "ellipse", accent),
            _text(f"{side_name}-label", f"comparison.{side_name}.label", _box(x + 88, card_y + 35, 420, 65), side["label"], brand, size=23, bold=True),
            _text(f"{side_name}-items", f"comparison.{side_name}.items", _box(x + 42, card_y + 125, 465, card_h - 155), "\n".join(f"• {item}" for item in side["items"]), brand, size=18),
        ])
    nodes.extend(_footer(spec, _source_map(deck), brand))
    return nodes


def _compile_chart(spec: dict[str, Any], deck: dict[str, Any], variant: str) -> list[dict[str, Any]]:
    brand = deck["brand"]; colors = brand["colors"]
    chart_w, takeaway_x = (760, 885) if variant == "balanced" else (820, 945)
    nodes = [
        _shape("background", "background", _box(0, 0, _WIDTH, _HEIGHT), "rect", colors["background"]),
        _text("title", "title", _box(70, 45, 1140, 80), spec["title"], brand, size=34, bold=True),
        {"node_id": "chart", "kind": "chart", "role": "chart.primary", "box": _box(75, 165, chart_w, 440), "chart": copy.deepcopy(spec["chart"])},
        _shape("takeaway-card", "takeaway.card", _box(takeaway_x, 165, 1280 - takeaway_x - 70, 440), "round_rect", colors["text"], radius=18),
        _text("takeaway-label", "takeaway.label", _box(takeaway_x + 40, 205, 240, 30), "ВЫВОД", brand, size=11, bold=True, color=colors["accent"]),
        _text("takeaway", "takeaway.text", _box(takeaway_x + 40, 265, 1280 - takeaway_x - 150, 270), spec["takeaway"], brand, size=21, bold=True, color=colors["surface"], valign="middle"),
    ]
    nodes.extend(_footer(spec, _source_map(deck), brand))
    return nodes


def _common_nodes(spec: dict[str, Any], deck: dict[str, Any]) -> list[dict[str, Any]]:
    brand = deck["brand"]
    return [
        _shape("background", "background", _box(0, 0, _WIDTH, _HEIGHT), "rect", brand["colors"]["background"]),
        _text("title", "title", _box(70, 45, 1140, 80), spec["title"], brand, size=34, bold=True),
    ]


def _compile_process(spec: dict[str, Any], deck: dict[str, Any], variant: str) -> list[dict[str, Any]]:
    brand = deck["brand"]; colors = brand["colors"]; steps = spec["steps"]
    nodes = _common_nodes(spec, deck)
    left, right, y = 75, 1205, 235
    usable = right - left; gap = 28 if variant == "balanced" else 18
    card_w = (usable - gap * (len(steps) - 1)) / len(steps)
    for index, step in enumerate(steps):
        x = left + index * (card_w + gap)
        if index:
            nodes.append(_shape(f"connector-{index}", "process.connector", _box(x - gap + 4, y + 87, gap - 8, 1), "line", colors["muted"], stroke=colors["muted"]))
        nodes.extend([
            _shape(f"step-card-{index}", "process.step.card", _box(x, y, card_w, 255), "round_rect", colors["surface"], stroke="D9DEE3", radius=20),
            _shape(f"step-badge-{index}", "process.step.badge", _box(x + 24, y + 24, 54, 54), "ellipse", colors["accent"] if index == 0 else colors["accent_secondary"]),
            _text(f"step-number-{index}", "process.step.number", _box(x + 24, y + 32, 54, 36), str(index + 1), brand, size=20, bold=True, color=colors["surface"], align="center"),
            _text(f"step-label-{index}", "process.step.label", _box(x + 95, y + 25, card_w - 115, 60), step["label"], brand, size=21, bold=True),
            _text(f"step-description-{index}", "process.step.description", _box(x + 28, y + 115, card_w - 56, 105), step["description"], brand, size=17, color=colors["muted"]),
        ])
    nodes.extend(_footer(spec, _source_map(deck), brand))
    return nodes


def _compile_timeline(spec: dict[str, Any], deck: dict[str, Any], variant: str) -> list[dict[str, Any]]:
    brand = deck["brand"]; colors = brand["colors"]; milestones = spec["milestones"]
    nodes = _common_nodes(spec, deck)
    start, end, axis_y = 155, 1125, 300
    nodes.append(_shape("timeline-axis", "timeline.axis", _box(start, axis_y, end - start, 1), "line", colors["muted"], stroke=colors["muted"]))
    step = (end - start) / max(len(milestones) - 1, 1)
    card_w = min(210, step * 0.78 if len(milestones) > 1 else 210)
    for index, milestone in enumerate(milestones):
        x = start + index * step
        nodes.extend([
            _shape(f"milestone-{index}", "timeline.milestone", _box(x - 25, axis_y - 25, 50, 50), "ellipse", colors["accent"] if index in {0, len(milestones)-1} else colors["accent_secondary"]),
            _text(f"period-{index}", "timeline.period", _box(x - 75, axis_y - 105, 150, 45), milestone["period"], brand, size=17, bold=True, align="center"),
            _shape(f"timeline-card-{index}", "timeline.card", _box(x - card_w / 2, axis_y + 55, card_w, 135), "round_rect", colors["surface"], stroke="D9DEE3", radius=16),
            _text(f"milestone-label-{index}", "timeline.label", _box(x - card_w / 2 + 18, axis_y + 84, card_w - 36, 78), milestone["label"], brand, size=16, bold=True, align="center", valign="middle"),
        ])
    nodes.extend(_footer(spec, _source_map(deck), brand))
    return nodes


def _compile_decision_matrix(spec: dict[str, Any], deck: dict[str, Any], variant: str) -> list[dict[str, Any]]:
    brand = deck["brand"]; colors = brand["colors"]
    nodes = _common_nodes(spec, deck)
    x0, y0, label_w, cell_h = 90, 175, 270, 88
    cell_w = (1100 - label_w) / len(spec["criteria"])
    nodes.append(_shape("matrix-frame", "decision.frame", _box(x0, y0, 1100, cell_h * (len(spec["options"]) + 1)), "round_rect", colors["surface"], stroke="D9DEE3", radius=16))
    for criterion_index, criterion in enumerate(spec["criteria"]):
        nodes.append(_text(f"criterion-{criterion_index}", "decision.criterion", _box(x0 + label_w + criterion_index * cell_w, y0 + 15, cell_w, 55), criterion, brand, size=16, bold=True, align="center", valign="middle"))
    rating_colors = {"positive": colors["accent_secondary"], "neutral": colors["muted"], "negative": colors["accent"]}
    rating_labels = {"positive": "+", "neutral": "•", "negative": "−"}
    for option_index, option in enumerate(spec["options"]):
        y = y0 + cell_h * (option_index + 1)
        nodes.append(_text(f"option-{option_index}", "decision.option", _box(x0 + 25, y + 20, label_w - 50, 48), option["label"], brand, size=18, bold=True, valign="middle"))
        for criterion_index, rating in enumerate(option["ratings"]):
            cx = x0 + label_w + criterion_index * cell_w + cell_w / 2
            nodes.extend([
                _shape(f"rating-badge-{option_index}-{criterion_index}", "decision.rating", _box(cx - 22, y + 22, 44, 44), "ellipse", rating_colors[rating]),
                _text(f"rating-label-{option_index}-{criterion_index}", "decision.rating.label", _box(cx - 22, y + 27, 44, 30), rating_labels[rating], brand, size=20, bold=True, color=colors["surface"], align="center"),
            ])
    nodes.extend(_footer(spec, _source_map(deck), brand))
    return nodes


def _compile_kpi_grid(spec: dict[str, Any], deck: dict[str, Any], variant: str) -> list[dict[str, Any]]:
    brand = deck["brand"]; colors = brand["colors"]; metrics = spec["metrics"]
    nodes = _common_nodes(spec, deck)
    columns = 3 if len(metrics) > 2 else 2; rows = (len(metrics) + columns - 1) // columns
    left, top, total_w, total_h, gap = 75, 165, 1130, 445, 24
    card_w = (total_w - gap * (columns - 1)) / columns; card_h = (total_h - gap * (rows - 1)) / rows
    for index, metric in enumerate(metrics):
        row, column = divmod(index, columns); x = left + column * (card_w + gap); y = top + row * (card_h + gap)
        nodes.extend([
            _shape(f"metric-card-{index}", "kpi.card", _box(x, y, card_w, card_h), "round_rect", colors["surface"], stroke="D9DEE3", radius=18),
            _text(f"metric-label-{index}", "kpi.label", _box(x + 28, y + 25, card_w - 56, 40), metric["label"], brand, size=16, bold=True, color=colors["muted"]),
            _text(f"metric-value-{index}", "kpi.value", _box(x + 28, y + 75, card_w - 56, 65), metric["value"], brand, size=27, bold=True, color=colors["accent_secondary"]),
            _text(f"metric-note-{index}", "kpi.note", _box(x + 28, y + 145, card_w - 56, max(45, card_h - 170)), metric["note"], brand, size=15, color=colors["text"]),
        ])
    nodes.extend(_footer(spec, _source_map(deck), brand))
    return nodes


def compile_deck(deck_spec: dict[str, Any], *, variants: dict[str, str] | None = None, slide_ids: Iterable[str] | None = None) -> dict[str, Any]:
    """Compile approved semantic intent to a validated, defensive SceneSpec."""
    deck = validate_deck_spec(deck_spec)
    if variants is not None and not isinstance(variants, dict):
        raise CompileError("variants must be an object")
    variants = dict(variants or {})
    if slide_ids is not None and not isinstance(slide_ids, (list, tuple)):
        raise CompileError("slide_ids must be an array")
    selected_ids = list(slide_ids) if slide_ids is not None else [slide["slide_id"] for slide in deck["slides"]]
    if any(not isinstance(slide_id, str) for slide_id in selected_ids):
        raise CompileError("slide_ids must contain strings")
    known_ids = {slide["slide_id"] for slide in deck["slides"]}
    known_archetypes = {slide["archetype"] for slide in deck["slides"]}
    unknown_targets = sorted(set(variants) - known_ids - known_archetypes)
    if unknown_targets:
        raise CompileError(f"unknown variant target: {unknown_targets[0]}")
    if any(not isinstance(value, str) for value in variants.values()):
        raise CompileError("variant names must be strings")
    unknown = [slide_id for slide_id in selected_ids if slide_id not in known_ids]
    if unknown:
        raise CompileError(f"unknown slide_id: {unknown[0]}")
    selected = [slide for slide in deck["slides"] if slide["slide_id"] in set(selected_ids)]
    compiled_slides: list[dict[str, Any]] = []
    supported_variants = _supported_variants()
    if set(supported_variants) != set(_RECIPE_COMPILERS):
        raise CompileError("managed catalog and recipe compiler registry are out of sync")
    for spec in selected:
        archetype = spec["archetype"]
        variant = variants.get(spec["slide_id"], variants.get(archetype, "balanced"))
        if variant not in supported_variants[archetype]:
            raise CompileError(f"unsupported variant for {archetype}: {variant}")
        if archetype == "cover":
            nodes = _compile_cover(spec, deck, variant)
        elif archetype == "comparison":
            nodes = _compile_comparison(spec, deck, variant)
        elif archetype == "chart_with_takeaway":
            nodes = _compile_chart(spec, deck, variant)
        elif archetype == "process":
            nodes = _compile_process(spec, deck, variant)
        elif archetype == "timeline":
            nodes = _compile_timeline(spec, deck, variant)
        elif archetype == "decision_matrix":
            nodes = _compile_decision_matrix(spec, deck, variant)
        elif archetype == "kpi_grid":
            nodes = _compile_kpi_grid(spec, deck, variant)
        else:
            raise CompileError(f"managed archetype has no recipe compiler: {archetype}")
        compiled_slides.append({"slide_id": spec["slide_id"], "archetype": archetype, "variant": variant, "nodes": nodes})
    scene = {"scene_version": "1.0", "deck_id": deck["deck_id"], "canvas": {"width": _WIDTH, "height": _HEIGHT}, "slides": compiled_slides}
    return validate_scene_spec(scene)
