"""Managed catalog of bounded composition capabilities.

STICKY BOUNDARY: an archetype is a named composition of reusable internal
components. Components may eventually compile to PPTX, SVG preview, or another
backend. The catalog describes capability and lifecycle; it is not executable
Python and never contains free-form renderer coordinates from the LLM.
"""

from __future__ import annotations

import copy
from typing import Any


class LibraryError(ValueError):
    """The managed composition catalog is internally inconsistent."""


_CATALOG: dict[str, Any] = {
    "catalog_version": "1.0",
    "components": {
        "background": {"kind": "surface", "status": "stable", "bounded": True},
        "text": {"kind": "content", "status": "stable", "bounded": True},
        "shape": {"kind": "visual", "status": "stable", "bounded": True},
        "image": {"kind": "visual", "status": "stable", "bounded": True},
        "native_chart": {"kind": "data", "status": "stable", "bounded": True},
        "title": {"kind": "composition", "status": "stable", "bounded": True},
        "cover_panel": {"kind": "composition", "status": "stable", "bounded": True},
        "two_column_cards": {"kind": "composition", "status": "stable", "bounded": True},
        "chart_panel": {"kind": "composition", "status": "stable", "bounded": True},
        "takeaway_card": {"kind": "composition", "status": "stable", "bounded": True},
        "source_footer": {"kind": "composition", "status": "stable", "bounded": True},
        "process_steps": {"kind": "composition", "status": "stable", "bounded": True},
        "timeline_track": {"kind": "composition", "status": "stable", "bounded": True},
        "decision_grid": {"kind": "composition", "status": "stable", "bounded": True},
        "metric_cards": {"kind": "composition", "status": "stable", "bounded": True},
    },
    "archetypes": {
        "cover": {
            "status": "stable",
            "variants": ["balanced", "dark"],
            "components": ["background", "title", "cover_panel", "source_footer"],
            "preview_fidelity": "structural",
        },
        "comparison": {
            "status": "stable",
            "variants": ["balanced", "compact"],
            "components": ["background", "title", "two_column_cards", "source_footer"],
            "preview_fidelity": "structural",
        },
        "chart_with_takeaway": {
            "status": "stable",
            "variants": ["balanced", "emphasis"],
            "components": ["background", "title", "chart_panel", "takeaway_card", "source_footer"],
            "preview_fidelity": "structural",
        },
        "process": {
            "status": "stable",
            "variants": ["balanced", "compact"],
            "components": ["background", "title", "process_steps", "source_footer"],
            "preview_fidelity": "structural",
        },
        "timeline": {
            "status": "stable",
            "variants": ["balanced"],
            "components": ["background", "title", "timeline_track", "source_footer"],
            "preview_fidelity": "structural",
        },
        "decision_matrix": {
            "status": "stable",
            "variants": ["balanced"],
            "components": ["background", "title", "decision_grid", "source_footer"],
            "preview_fidelity": "structural",
        },
        "kpi_grid": {
            "status": "stable",
            "variants": ["balanced"],
            "components": ["background", "title", "metric_cards", "source_footer"],
            "preview_fidelity": "structural",
        },
    },
}


def validate_catalog(raw: Any) -> dict[str, Any]:
    if not isinstance(raw, dict) or set(raw) != {"catalog_version", "components", "archetypes"}:
        raise LibraryError("catalog must have exactly version, components and archetypes")
    if raw["catalog_version"] != "1.0":
        raise LibraryError("unsupported catalog version")
    components = raw["components"]
    archetypes = raw["archetypes"]
    if not isinstance(components, dict) or not isinstance(archetypes, dict):
        raise LibraryError("components and archetypes must be objects")
    for name, component in components.items():
        if not isinstance(name, str) or not name or not isinstance(component, dict):
            raise LibraryError("invalid component entry")
        if set(component) != {"kind", "status", "bounded"}:
            raise LibraryError(f"component {name} has invalid fields")
        if component["bounded"] is not True:
            raise LibraryError(f"component {name} must be bounded")
        if component["status"] not in {"experimental", "stable", "deprecated"}:
            raise LibraryError(f"component {name} has invalid status")
    for name, archetype in archetypes.items():
        if not isinstance(name, str) or not name or not isinstance(archetype, dict):
            raise LibraryError("invalid archetype entry")
        if set(archetype) != {"status", "variants", "components", "preview_fidelity"}:
            raise LibraryError(f"archetype {name} has invalid fields")
        if archetype["status"] not in {"experimental", "stable", "deprecated"}:
            raise LibraryError(f"archetype {name} has invalid status")
        if archetype["preview_fidelity"] not in {"structural", "renderer", "application"}:
            raise LibraryError(f"archetype {name} has invalid preview fidelity")
        if not isinstance(archetype["variants"], list) or not archetype["variants"] or len(archetype["variants"]) != len(set(archetype["variants"])):
            raise LibraryError(f"archetype {name} has invalid variants")
        if any(not isinstance(variant, str) or not variant for variant in archetype["variants"]):
            raise LibraryError(f"archetype {name} has invalid variant")
        if not isinstance(archetype["components"], list) or not archetype["components"]:
            raise LibraryError(f"archetype {name} needs components")
        for component_name in archetype["components"]:
            if component_name not in components:
                raise LibraryError(f"archetype {name} references unknown component: {component_name}")
    return copy.deepcopy(raw)


def get_catalog() -> dict[str, Any]:
    """Return a validated defensive copy suitable for agents and tooling."""
    return validate_catalog(copy.deepcopy(_CATALOG))
