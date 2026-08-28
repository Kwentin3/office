from __future__ import annotations

import json
from pathlib import Path

import pytest
from office_artifact_tool.docx.preview import render_docx_preview
from office_artifact_tool.docx.review_contract import ReviewContractError, validate_review_packet


def test_review_packet_is_recursively_closed_and_chat_only(tmp_path: Path) -> None:
    output = tmp_path / "review"
    render_docx_preview({"blocks": [{"type": "paragraph", "text": "Review me"}]}, output)
    packet = json.loads((output / "review.json").read_text(encoding="utf-8"))

    assert validate_review_packet(packet) == packet
    assert packet["interaction"] == "chat_only"
    assert packet["fidelity"] == "styled_layout_proxy_not_word_render"
    assert packet["presentation_id"] == "professional-a4/v2"

    forbidden = {"prompt", "callback", "office", "coordinates", "mutation", "instructions"}

    def keys(value: object) -> set[str]:
        if isinstance(value, dict):
            return set(value) | set().union(*(keys(item) for item in value.values()))
        if isinstance(value, list):
            return set().union(*(keys(item) for item in value))
        return set()

    assert keys(packet).isdisjoint(forbidden)

    unknown_cases = [
        (packet, "unexpected"),
        (packet["artifact"], "unexpected"),
        (packet["diagnostics"], "unexpected"),
        (packet["diagnostics"]["limits"], "unexpected"),
    ]
    for container, key in unknown_cases:
        broken = json.loads(json.dumps(packet))
        target = broken
        if container is packet["artifact"]:
            target = broken["artifact"]
        elif container is packet["diagnostics"]:
            target = broken["diagnostics"]
        elif container is packet["diagnostics"]["limits"]:
            target = broken["diagnostics"]["limits"]
        target[key] = True
        with pytest.raises(ReviewContractError, match="unknown field"):
            validate_review_packet(broken)
