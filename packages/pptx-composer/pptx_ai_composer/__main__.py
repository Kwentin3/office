"""Small JSON-lines CLI for agent/runtime integration."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

from .contracts import ContractError
from .renderer import RenderError, render_deck
from .preview import PreviewError, render_preview
from .library import get_catalog
from .validator import validate_presentation


def _load_object(path: Any, field: str) -> tuple[Path, dict[str, Any]]:
    if not isinstance(path, str) or not path:
        raise ValueError(f"{field} must be a path string")
    resolved = Path(path).resolve()
    with resolved.open("r", encoding="utf-8") as stream:
        value = json.load(stream)
    if not isinstance(value, dict):
        raise ValueError(f"{field} must contain a JSON object")
    return resolved, value


def _dispatch(request: Any) -> dict[str, Any]:
    if not isinstance(request, dict):
        raise ValueError("request must be an object")
    action = request.get("action")
    allowed_by_action = {
        "render": {"action", "spec", "output", "variants", "slide_ids"},
        "preview": {"action", "spec", "output", "variants", "slide_ids"},
        "validate": {"action", "spec", "source"},
        "catalog": {"action"},
    }
    if action not in allowed_by_action:
        raise ValueError("unsupported action")
    unknown = set(request) - allowed_by_action[action]
    if unknown:
        raise ValueError(f"request contains unknown field: {sorted(unknown)[0]}")
    if action == "render":
        spec_path, spec = _load_object(request.get("spec"), "spec")
        output = request.get("output")
        if not isinstance(output, str) or not output:
            raise ValueError("output must be a path string")
        return render_deck(spec, output, protected_paths=[spec_path], variants=request.get("variants"), slide_ids=request.get("slide_ids"))
    if action == "validate":
        _, spec = _load_object(request.get("spec"), "spec")
        source = request.get("source")
        if not isinstance(source, str) or not source:
            raise ValueError("source must be a path string")
        return validate_presentation(source, spec)
    if action == "preview":
        spec_path, spec = _load_object(request.get("spec"), "spec")
        output = request.get("output")
        if not isinstance(output, str) or not output:
            raise ValueError("output must be a path string")
        return render_preview(spec, output, protected_paths=[spec_path], variants=request.get("variants"), slide_ids=request.get("slide_ids"))
    if action == "catalog":
        return {"status": "ok", **get_catalog()}
    raise ValueError("unsupported action")


def main() -> int:
    try:
        request = json.loads(sys.stdin.readline())
        response = _dispatch(request)
        status = 0
    except (json.JSONDecodeError, OSError, ValueError, ContractError, RenderError, PreviewError) as exc:
        response = {"status": "error", "code": "invalid_request", "message": str(exc)}
        status = 2
    print(json.dumps(response, ensure_ascii=False, separators=(",", ":")))
    return status


if __name__ == "__main__":
    raise SystemExit(main())
