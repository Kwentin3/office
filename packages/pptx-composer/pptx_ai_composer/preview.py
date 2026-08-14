"""Fast structural preview backend for conversational iteration.

STICKY CLAIM BOUNDARY: this backend consumes only validated SceneSpec nodes. It
is an immediate design proxy, not a PowerPoint render. Semantic archetypes and
layout recipes belong to the compiler; this module knows only bounded drawing
primitives and atomic publication.
"""

from __future__ import annotations

import hashlib
import html
import json
import os
import shutil
import tempfile
from functools import lru_cache
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree

from PIL import Image, ImageDraw, ImageFont

from .compiler import compile_deck
from .contracts import validate_deck_spec
from .scene_contract import validate_scene_spec


class PreviewError(ValueError):
    """A deterministic preview refusal."""


_PREVIEW_FONT = Path(__file__).with_name("assets") / "NotoSans-Regular.ttf"
_PREVIEW_FONT_SHA256 = "bfb7bb691513f12e734dc346c03a03f784912432d7e3fa8e56efcf906fe86b3d"


@lru_cache(maxsize=32)
def _preview_font(size: int) -> ImageFont.FreeTypeFont:
    if not _PREVIEW_FONT.is_file() or _PREVIEW_FONT.is_symlink():
        raise PreviewError("bundled preview font is missing or unsafe")
    if hashlib.sha256(_PREVIEW_FONT.read_bytes()).hexdigest() != _PREVIEW_FONT_SHA256:
        raise PreviewError("bundled preview font hash mismatch")
    return ImageFont.truetype(str(_PREVIEW_FONT), size=size)


def _esc(value: Any) -> str:
    return html.escape(str(value), quote=True)


def _wrap_paragraph(text: str, limit: int) -> list[str]:
    words = text.split()
    lines: list[str] = []
    current: list[str] = []
    for word in words:
        candidate = " ".join([*current, word])
        if current and len(candidate) > limit:
            lines.append(" ".join(current))
            current = [word]
        else:
            current.append(word)
    if current:
        lines.append(" ".join(current))
    return lines or [""]


def _wrapped_lines(text: str, box_width: float, size: float) -> list[str]:
    limit = max(8, int(box_width / max(size * 0.55, 1)))
    lines: list[str] = []
    for paragraph in text.splitlines() or [text]:
        lines.extend(_wrap_paragraph(paragraph, limit))
    return lines


def _svg_text(node: dict[str, Any]) -> str:
    box = node["box"]; style = node["style"]
    lines = _wrapped_lines(node["text"], box["w"], style["size"])
    line_height = round(style["size"] * 1.28)
    max_lines = max(1, int(box["h"] // line_height))
    lines = lines[:max_lines]
    x = box["x"]
    anchor = {"left": "start", "center": "middle", "right": "end"}[style["align"]]
    if style["align"] == "center":
        x += box["w"] / 2
    elif style["align"] == "right":
        x += box["w"]
    text_height = len(lines) * line_height
    if style["valign"] == "middle":
        y = box["y"] + max(style["size"], (box["h"] - text_height) / 2 + style["size"])
    elif style["valign"] == "bottom":
        y = box["y"] + box["h"] - text_height + style["size"]
    else:
        y = box["y"] + style["size"]
    spans = "".join(f'<tspan x="{x:.2f}" dy="{0 if index == 0 else line_height}">{_esc(line)}</tspan>' for index, line in enumerate(lines))
    return (
        f'<text data-node-id="{_esc(node["node_id"])}" data-role="{_esc(node["role"])}" '
        f'x="{x:.2f}" y="{y:.2f}" text-anchor="{anchor}" font-family="Noto Sans, sans-serif" '
        f'font-size="{style["size"]}" font-weight="{700 if style["bold"] else 400}" '
        f'fill="#{style["color"]}">{spans}</text>'
    )


def _svg_shape(node: dict[str, Any]) -> str:
    box = node["box"]
    common = f'data-node-id="{_esc(node["node_id"])}" data-role="{_esc(node["role"])}" fill="#{node["fill"]}"'
    if "stroke" in node:
        common += f' stroke="#{node["stroke"]}"'
    if "opacity" in node:
        common += f' opacity="{node["opacity"]}"'
    if node["shape"] in {"rect", "round_rect"}:
        radius = node.get("radius", 0) if node["shape"] == "round_rect" else 0
        return f'<rect {common} x="{box["x"]}" y="{box["y"]}" width="{box["w"]}" height="{box["h"]}" rx="{radius}"/>'
    if node["shape"] == "ellipse":
        return f'<ellipse {common} cx="{box["x"] + box["w"] / 2}" cy="{box["y"] + box["h"] / 2}" rx="{box["w"] / 2}" ry="{box["h"] / 2}"/>'
    return f'<line data-node-id="{_esc(node["node_id"])}" data-role="{_esc(node["role"])}" x1="{box["x"]}" y1="{box["y"]}" x2="{box["x"] + box["w"]}" y2="{box["y"] + box["h"]}" stroke="#{node.get("stroke", node["fill"])}" stroke-width="3"/>'


def _svg_chart(node: dict[str, Any]) -> str:
    box = node["box"]; chart = node["chart"]
    values = [value for series in chart["series"] for value in series["values"]]
    maximum = max([abs(float(value)) for value in values] or [1.0]) or 1.0
    primary = chart["series"][0]["values"]
    categories = chart["categories"]
    gap = box["w"] / max(len(categories), 1)
    pieces = [f'<g data-node-id="{_esc(node["node_id"])}" data-role="{_esc(node["role"])}">', f'<rect x="{box["x"]}" y="{box["y"]}" width="{box["w"]}" height="{box["h"]}" rx="18" fill="#FFFFFF" stroke="#D9DEE3"/>']
    if chart["type"] == "bar":
        for index, value in enumerate(primary):
            height = abs(float(value)) / maximum * box["h"] * 0.62
            x = box["x"] + gap * index + gap * 0.22
            width = gap * 0.56
            y = box["y"] + box["h"] * 0.72 - height
            pieces.append(f'<rect x="{x:.2f}" y="{y:.2f}" width="{width:.2f}" height="{height:.2f}" fill="#D95D39"/>')
    else:
        points = []
        for index, value in enumerate(primary):
            x = box["x"] + 45 + index * ((box["w"] - 90) / max(len(primary) - 1, 1))
            y = box["y"] + box["h"] * 0.72 - (float(value) / maximum * box["h"] * 0.55)
            points.append(f"{x:.2f},{y:.2f}")
        pieces.append(f'<polyline points="{" ".join(points)}" fill="none" stroke="#D95D39" stroke-width="4"/>')
    for index, category in enumerate(categories):
        x = box["x"] + gap * index + gap / 2
        y = box["y"] + box["h"] * 0.84
        pieces.append(f'<text x="{x:.2f}" y="{y:.2f}" text-anchor="middle" font-family="Noto Sans, sans-serif" font-size="12" fill="#62707D">{_esc(category)}</text>')
    pieces.append("</g>")
    return "".join(pieces)


def _slide_svg(scene: dict[str, Any], slide: dict[str, Any]) -> str:
    width = scene["canvas"]["width"]; height = scene["canvas"]["height"]
    out = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">']
    for node in slide["nodes"]:
        if node["kind"] == "shape":
            out.append(_svg_shape(node))
        elif node["kind"] == "text":
            out.append(_svg_text(node))
        elif node["kind"] == "chart":
            out.append(_svg_chart(node))
        else:
            box = node["box"]
            out.append(f'<rect data-node-id="{_esc(node["node_id"])}" data-role="{_esc(node["role"])}" x="{box["x"]}" y="{box["y"]}" width="{box["w"]}" height="{box["h"]}" fill="#D9DEE3"/>')
            out.append(f'<text x="{box["x"] + box["w"] / 2}" y="{box["y"] + box["h"] / 2}" text-anchor="middle" font-family="Noto Sans, sans-serif" font-size="18" fill="#62707D">ASSET: {_esc(node["asset_id"])}</text>')
    out.append("</svg>")
    return "".join(out)


def _rasterize_bounded_svg(svg: str, output: Path, width: int, height: int) -> None:
    """Rasterize only this backend's finite SVG primitive subset."""
    root = ElementTree.fromstring(svg)
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    for element in root.iter():
        tag = element.tag.rsplit("}", 1)[-1]
        if tag == "rect":
            x = float(element.get("x", "0")); y = float(element.get("y", "0")); w = float(element.get("width", str(width))); h = float(element.get("height", str(height)))
            box = (round(x), round(y), round(x + w), round(y + h)); radius = round(float(element.get("rx", "0")))
            if radius:
                draw.rounded_rectangle(box, radius=radius, fill=element.get("fill", "#FFFFFF"), outline=element.get("stroke"))
            else:
                draw.rectangle(box, fill=element.get("fill", "#FFFFFF"), outline=element.get("stroke"))
        elif tag in {"circle", "ellipse"}:
            cx = float(element.get("cx", "0")); cy = float(element.get("cy", "0")); rx = float(element.get("rx", element.get("r", "0"))); ry = float(element.get("ry", element.get("r", "0")))
            draw.ellipse((round(cx-rx), round(cy-ry), round(cx+rx), round(cy+ry)), fill=element.get("fill", "#000000"), outline=element.get("stroke"))
        elif tag == "line":
            draw.line((float(element.get("x1", "0")), float(element.get("y1", "0")), float(element.get("x2", "0")), float(element.get("y2", "0"))), fill=element.get("stroke", "#000000"), width=round(float(element.get("stroke-width", "3"))))
        elif tag == "polyline":
            points = [tuple(map(float, point.split(","))) for point in element.get("points", "").split()]
            if len(points) > 1:
                draw.line(points, fill=element.get("stroke", "#000000"), width=round(float(element.get("stroke-width", "3"))))
        elif tag == "text":
            size = max(8, int(float(element.get("font-size", "16")))); font = _preview_font(size)
            x = float(element.get("x", "0")); y = float(element.get("y", "0")) - size
            anchor = element.get("text-anchor", "start")
            children = list(element) or [element]
            for child_index, child in enumerate(children):
                if child_index:
                    y += float(child.get("dy", str(round(size * 1.28))))
                value = "".join(child.itertext()); bbox = draw.textbbox((0, 0), value, font=font); text_width = bbox[2] - bbox[0]
                draw_x = x - text_width / 2 if anchor == "middle" else x - text_width if anchor == "end" else x
                draw.text((draw_x, y), value, fill=element.get("fill", "#000000"), font=font)
    image.save(output, "PNG")


def _text_overflow_diagnostics(scene: dict[str, Any]) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    for slide in scene["slides"]:
        for node in slide["nodes"]:
            if node["kind"] != "text":
                continue
            box = node["box"]; style = node["style"]
            required_lines = len(_wrapped_lines(node["text"], box["w"], style["size"]))
            line_height = round(style["size"] * 1.28)
            available_lines = max(1, int(box["h"] // line_height))
            if required_lines > available_lines:
                issues.append({
                    "slide_id": slide["slide_id"],
                    "node_id": node["node_id"],
                    "role": node["role"],
                    "required_lines": required_lines,
                    "available_lines": available_lines,
                })
    return issues


def _contains_protected(container: Path, protected: set[Path]) -> bool:
    return any(container == path or container in path.parents for path in protected)


def _assert_safe_publication_paths(output: Path, backup: Path, protected: set[Path]) -> None:
    # STICKY SAFETY BOUNDARY: output and its deterministic backup are both
    # destructive publication targets, so both must pass the same admission.
    if output.is_symlink() or backup.is_symlink():
        raise PreviewError("preview output and backup must not be symlinks")
    if output in protected:
        raise PreviewError("output must not collide with a protected path")
    if _contains_protected(output, protected):
        raise PreviewError("output directory must not contain a protected path")
    if backup in protected or _contains_protected(backup, protected):
        raise PreviewError("preview backup must not collide with or contain a protected path")
    if output.exists() and not output.is_dir():
        raise PreviewError("existing preview output must be a directory")
    if backup.exists() and not backup.is_dir():
        raise PreviewError("preview backup path must be a directory")


def render_scene_preview(scene_spec: dict[str, Any], output: str | Path, *, protected_paths: Iterable[str | Path] = ()) -> dict[str, Any]:
    """Render validated SceneSpec to atomically published SVG/PNG previews."""
    scene = validate_scene_spec(scene_spec)
    raw_output = Path(output).expanduser().absolute()
    raw_backup = raw_output.with_name(f".{raw_output.name}.old")
    if raw_output.is_symlink() or raw_backup.is_symlink():
        raise PreviewError("preview output and backup must not be symlinks")
    output_path = raw_output.resolve(strict=False)
    protected = {Path(path).resolve(strict=False) for path in protected_paths}
    backup = raw_backup.resolve(strict=False)
    _assert_safe_publication_paths(output_path, backup, protected)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = Path(tempfile.mkdtemp(prefix=f".{output_path.name}.", dir=output_path.parent))
    try:
        slides: list[dict[str, Any]] = []
        width = round(scene["canvas"]["width"]); height = round(scene["canvas"]["height"])
        for index, slide in enumerate(scene["slides"], 1):
            svg = _slide_svg(scene, slide)
            svg_name = f"slide-{index:02d}.svg"; png_name = f"slide-{index:02d}.png"
            (temporary / svg_name).write_text(svg, encoding="utf-8")
            _rasterize_bounded_svg(svg, temporary / png_name, width, height)
            slides.append({"slide_id": slide["slide_id"], "archetype": slide["archetype"], "variant": slide["variant"], "file": svg_name, "png_file": png_name, "preview_fidelity": "structural"})
        manifest = {
            "manifest_version": "1.0",
            "deck_id": scene["deck_id"],
            "scene_version": scene["scene_version"],
            "fidelity": "structural_preview_not_powerpoint_render",
            "limitations": ["font metrics and substitution are not PowerPoint-accurate", "PowerPoint application repair behavior is not validated", "native charts are represented approximately"],
            "diagnostics": {"text_overflow": _text_overflow_diagnostics(scene)},
            "slides": slides,
        }
        (temporary / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        backup = output_path.with_name(f".{output_path.name}.old")
        if backup.exists():
            shutil.rmtree(backup)
        if output_path.exists():
            os.replace(output_path, backup)
        os.replace(temporary, output_path)
        if backup.exists():
            shutil.rmtree(backup)
    except Exception:
        shutil.rmtree(temporary, ignore_errors=True)
        raise
    return {"status": "previewed", "deck_id": scene["deck_id"], "slide_count": len(scene["slides"]), "output": str(output_path)}


def render_preview(deck_spec: dict[str, Any], output: str | Path, *, protected_paths: Iterable[str | Path] = (), variants: dict[str, str] | None = None, slide_ids: Iterable[str] | None = None) -> dict[str, Any]:
    """Compile semantic DeckSpec, then invoke the isolated SceneSpec backend."""
    deck = validate_deck_spec(deck_spec)
    asset_paths = [Path(asset["path"]).resolve() for asset in deck["assets"]]
    asset_paths.extend(Path(asset["fallback_path"]).resolve() for asset in deck["assets"] if asset["kind"] == "svg")
    scene = compile_deck(deck, variants=variants, slide_ids=slide_ids)
    return render_scene_preview(scene, output, protected_paths=[*protected_paths, *asset_paths])
