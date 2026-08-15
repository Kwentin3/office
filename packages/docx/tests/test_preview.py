from __future__ import annotations

import hashlib
import json
import re
from copy import deepcopy
from pathlib import Path

import pytest
from office_artifact_tool import api as api_module
from office_artifact_tool import DocxArtifactTool
from office_artifact_tool.docx import preview as preview_module
from office_artifact_tool.docx.preview import (
    MAX_LIST_ITEMS_PER_BLOCK,
    MAX_PREVIEW_BLOCKS,
    MAX_TABLE_COLUMNS_PER_ROW,
    MAX_TABLE_ROWS_PER_BLOCK,
    MAX_TEXT_CHARACTERS_PER_VALUE,
    render_docx_preview,
)
from office_artifact_tool.core.contracts import validate_create_model


def _revision(model: dict) -> str:
    payload = json.dumps(model, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def test_revision_is_lowercase_sha256_of_full_exact_model_and_changes_with_model(tmp_path: Path) -> None:
    first = {"document_id": "quarterly-note", "blocks": [{"type": "paragraph", "text": "Alpha"}]}
    second = {"document_id": "quarterly-note", "blocks": [{"type": "paragraph", "text": "Beta"}]}

    first_result = render_docx_preview(first, tmp_path / "first")
    second_result = render_docx_preview(second, tmp_path / "second")

    assert first_result["revision"] == _revision(first)
    assert second_result["revision"] == _revision(second)
    assert first_result["revision"] != second_result["revision"]


def test_review_digest_binds_exact_published_html_bytes(tmp_path: Path) -> None:
    output = tmp_path / "review"
    result = render_docx_preview({"blocks": [{"type": "heading", "text": "Digest me"}]}, output)

    html_path = output / "document.html"
    review = json.loads((output / "review.json").read_text(encoding="utf-8"))
    assert set(output.iterdir()) == {html_path, output / "review.json"}
    assert review["artifact"] == {
        "file": "document.html",
        "sha256": hashlib.sha256(html_path.read_bytes()).hexdigest(),
    }
    assert result["display_artifacts"] == [str(html_path.resolve())]
    assert result["review_contract"] == str((output / "review.json").resolve())


def test_heading_level_cannot_inject_markup_or_preview_model_refused_by_create(tmp_path: Path) -> None:
    model = {
        "blocks": [
            {
                "type": "heading",
                "level": '<script id="pwn">alert(1)</script>',
                "text": "safe",
            }
        ]
    }
    output = tmp_path / "refused"

    preview = render_docx_preview(model, output)
    created = DocxArtifactTool(tmp_path / "work").create(model, tmp_path / "refused.docx")

    expected = {"status": "refused", "reason": "validation_failure", "details": "invalid heading"}
    assert preview == expected
    assert created == expected
    assert not output.exists()
    assert not (tmp_path / "refused.docx").exists()


def test_validated_model_is_a_deep_defensive_copy() -> None:
    model = {"blocks": [{"type": "heading", "level": 1, "text": "safe"}]}

    validated = validate_create_model(model)
    model["blocks"][0]["level"] = 2

    assert validated == {"blocks": [{"type": "heading", "level": 1, "text": "safe"}]}
    assert validated is not model
    assert validated["blocks"] is not model["blocks"]


def test_preview_uses_one_validated_snapshot_when_caller_mutates_after_revision(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    model = {"blocks": [{"type": "heading", "level": 1, "text": "safe"}]}
    original = deepcopy(model)
    real_revision = preview_module._revision

    def mutate_caller_after_revision(validated: dict) -> str:
        revision = real_revision(validated)
        model["blocks"][0]["level"] = '<script id="pwn">alert(1)</script>'
        return revision

    monkeypatch.setattr(preview_module, "_revision", mutate_caller_after_revision)
    output = tmp_path / "review"
    result = render_docx_preview(model, output)

    rendered = (output / "document.html").read_text(encoding="utf-8")
    assert result["revision"] == _revision(original)
    assert "<h1>safe</h1>" in rendered
    assert "<script" not in rendered.lower()


def test_create_uses_the_validated_snapshot_when_caller_mutates_after_admission(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    model = {"blocks": [{"type": "heading", "level": 1, "text": "safe"}]}
    real_validate = api_module.validate_create_model

    def mutate_caller_after_admission(candidate: dict) -> dict:
        validated = real_validate(candidate)
        model["blocks"][0]["level"] = '<script id="pwn">alert(1)</script>'
        return validated

    monkeypatch.setattr(api_module, "validate_create_model", mutate_caller_after_admission)
    output = tmp_path / "safe.docx"
    result = DocxArtifactTool(tmp_path / "work").create(model, output)

    assert result["status"] == "ok"
    assert output.is_file()


def test_unknown_styles_and_non_xml_text_are_refused_by_preview_and_create(tmp_path: Path) -> None:
    invalid_models = [
        {"blocks": [{"type": "paragraph", "text": "safe", "style": "Unknown Style"}]},
        {"blocks": [{"type": "table", "rows": [["safe"]], "style": "Unknown Table"}]},
        {"metadata": {"title": "bad\x00title"}, "blocks": []},
        {"blocks": [{"type": "paragraph", "text": "bad\x01text"}]},
        {"blocks": [{"type": "bulleted_list", "items": ["bad\x0bitem"]}]},
        {"blocks": [{"type": "table", "rows": [["bad\ud800cell"]]}]},
    ]
    tool = DocxArtifactTool(tmp_path / "work")

    for index, model in enumerate(invalid_models):
        preview_output = tmp_path / f"preview-{index}"
        docx_output = tmp_path / f"invalid-{index}.docx"
        preview = render_docx_preview(model, preview_output)
        created = tool.create(model, docx_output)

        assert preview["status"] == "refused"
        assert preview["reason"] == "validation_failure"
        assert created["status"] == "refused"
        assert created["reason"] == "validation_failure"
        assert not preview_output.exists()
        assert not docx_output.exists()


def test_html_escapes_all_user_content_and_has_no_active_or_external_content(tmp_path: Path) -> None:
    hostile = '<script src="https://evil.invalid/x.js">alert("x")</script>&\"'
    model = {
        "metadata": {"title": hostile},
        "blocks": [
            {"type": "heading", "text": hostile},
            {"type": "paragraph", "text": hostile},
            {"type": "numbered_list", "items": [hostile]},
            {"type": "bulleted_list", "items": [hostile]},
            {"type": "table", "rows": [[hostile, 7]]},
        ],
    }

    output = tmp_path / "escaped"
    render_docx_preview(model, output)
    document = (output / "document.html").read_text(encoding="utf-8")
    lowered = document.lower()

    assert hostile not in document
    assert "&lt;script src=&quot;https://evil.invalid/x.js&quot;&gt;" in document
    assert "<script" not in lowered
    assert re.search(r"<[^>]+\s(?:src|href)\s*=", lowered) is None
    assert "<link" not in lowered
    assert "@import" not in lowered


def test_explicit_and_fallback_block_ids_are_stable(tmp_path: Path) -> None:
    model = {
        "document_id": "stable-document",
        "blocks": [
            {"type": "heading", "block_id": "intro", "text": "Introduction"},
            {"type": "paragraph", "text": "Body"},
        ],
    }

    first = tmp_path / "first-ids"
    second = tmp_path / "second-ids"
    render_docx_preview(model, first)
    render_docx_preview(model, second)

    pattern = r'data-block-id="([^"]+)"'
    first_ids = re.findall(pattern, (first / "document.html").read_text(encoding="utf-8"))
    second_ids = re.findall(pattern, (second / "document.html").read_text(encoding="utf-8"))
    assert first_ids == ["intro", "block-0002"]
    assert second_ids == first_ids


def test_preview_is_bounded_and_reports_every_truncation(tmp_path: Path) -> None:
    long_text = "X" * (MAX_TEXT_CHARACTERS_PER_VALUE + 1)
    rows = [
        [f"r{row}c{column}" for column in range(MAX_TABLE_COLUMNS_PER_ROW + 1)]
        for row in range(MAX_TABLE_ROWS_PER_BLOCK + 1)
    ]
    blocks = [
        {"type": "paragraph", "block_id": "long-text", "text": long_text},
        {
            "type": "bulleted_list",
            "block_id": "long-list",
            "items": [f"item-{index}" for index in range(MAX_LIST_ITEMS_PER_BLOCK + 1)],
        },
        {"type": "table", "block_id": "large-table", "rows": rows},
        *[
            {"type": "paragraph", "text": "DROP ME" if index == MAX_PREVIEW_BLOCKS - 3 else f"body-{index}"}
            for index in range(MAX_PREVIEW_BLOCKS - 2)
        ],
    ]
    output = tmp_path / "bounded"

    render_docx_preview({"blocks": blocks}, output)

    document = (output / "document.html").read_text(encoding="utf-8")
    review = json.loads((output / "review.json").read_text(encoding="utf-8"))
    diagnostics = review["diagnostics"]
    assert document.count("<section data-block-id=") == MAX_PREVIEW_BLOCKS
    assert document.count("<li>") == MAX_LIST_ITEMS_PER_BLOCK
    assert document.count("<tr>") == MAX_TABLE_ROWS_PER_BLOCK
    assert document.count("<td>") == MAX_TABLE_ROWS_PER_BLOCK * MAX_TABLE_COLUMNS_PER_ROW
    assert long_text not in document
    assert f"<p>{'X' * MAX_TEXT_CHARACTERS_PER_VALUE}</p>" in document
    assert "DROP ME" not in document
    assert diagnostics["truncated"] is True
    assert diagnostics["total_blocks"] == MAX_PREVIEW_BLOCKS + 1
    assert diagnostics["rendered_blocks"] == MAX_PREVIEW_BLOCKS
    assert diagnostics["omitted_blocks"] == 1
    assert diagnostics["limits"] == {
        "max_blocks": MAX_PREVIEW_BLOCKS,
        "max_list_items_per_block": MAX_LIST_ITEMS_PER_BLOCK,
        "max_table_rows_per_block": MAX_TABLE_ROWS_PER_BLOCK,
        "max_table_columns_per_row": MAX_TABLE_COLUMNS_PER_ROW,
        "max_text_characters_per_value": MAX_TEXT_CHARACTERS_PER_VALUE,
    }
    assert diagnostics["truncations"] == [
        {"block_id": "long-text", "content": "text_characters", "omitted": 1},
        {"block_id": "long-list", "content": "list_items", "omitted": 1},
        {"block_id": "large-table", "content": "table_rows", "omitted": 1},
        {
            "block_id": "large-table",
            "content": "table_columns",
            "omitted": MAX_TABLE_ROWS_PER_BLOCK,
        },
    ]


def test_refuses_symlink_or_non_directory_output_without_publication(tmp_path: Path) -> None:
    model = {"blocks": [{"type": "paragraph", "text": "safe"}]}
    target = tmp_path / "target"
    target.mkdir()
    symlink = tmp_path / "linked-review"
    symlink.symlink_to(target, target_is_directory=True)

    with pytest.raises(ValueError, match="symlink"):
        render_docx_preview(model, symlink)
    assert not (target / "review.json").exists()

    regular_file = tmp_path / "not-a-directory"
    regular_file.write_text("keep", encoding="utf-8")
    with pytest.raises(ValueError, match="directory"):
        render_docx_preview(model, regular_file)
    assert regular_file.read_text(encoding="utf-8") == "keep"


def test_replacing_review_never_makes_published_directory_disappear(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    output = tmp_path / "review"
    render_docx_preview({"blocks": [{"type": "paragraph", "text": "first"}]}, output)
    original_replace = preview_module.os.replace
    visibility: list[bool] = []

    def observe_replace(source: str | Path, target: str | Path) -> None:
        visibility.append(output.exists() and (output / "document.html").exists() and (output / "review.json").exists())
        original_replace(source, target)
        visibility.append(output.exists() and (output / "document.html").exists() and (output / "review.json").exists())

    monkeypatch.setattr(preview_module.os, "replace", observe_replace)
    result = render_docx_preview({"blocks": [{"type": "paragraph", "text": "second"}]}, output)

    assert result["status"] == "previewed"
    assert visibility
    assert all(visibility)


def test_failed_atomic_swap_restores_previous_complete_review(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    output = tmp_path / "review"
    render_docx_preview({"blocks": [{"type": "paragraph", "text": "first"}]}, output)
    original_html = (output / "document.html").read_bytes()
    original_review = (output / "review.json").read_bytes()

    def fail_publish(source: Path, target: Path) -> None:
        raise OSError("simulated publish failure")

    monkeypatch.setattr(preview_module, "_exchange_directories", fail_publish)
    with pytest.raises(OSError, match="simulated publish failure"):
        render_docx_preview({"blocks": [{"type": "paragraph", "text": "second"}]}, output)

    assert (output / "document.html").read_bytes() == original_html
    assert (output / "review.json").read_bytes() == original_review


def test_cleanup_failure_after_commit_returns_successful_new_review(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    output = tmp_path / "review"
    first = render_docx_preview({"blocks": [{"type": "paragraph", "text": "first"}]}, output)
    original_rmtree = preview_module.shutil.rmtree
    backup = output.with_name(f".{output.name}.old")

    def fail_backup_cleanup(path: str | Path, *args: object, **kwargs: object) -> None:
        if Path(path) == backup:
            raise OSError("simulated backup cleanup failure")
        original_rmtree(path, *args, **kwargs)

    monkeypatch.setattr(preview_module.shutil, "rmtree", fail_backup_cleanup)
    second = render_docx_preview({"blocks": [{"type": "paragraph", "text": "second"}]}, output)

    assert second["status"] == "previewed"
    assert second["revision"] != first["revision"]
    assert "second" in (output / "document.html").read_text(encoding="utf-8")
    assert json.loads((output / "review.json").read_text(encoding="utf-8"))["revision"] == second["revision"]
    assert backup.exists()


def test_non_oserror_cleanup_failure_after_commit_returns_successful_new_review(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    output = tmp_path / "review"
    first = render_docx_preview({"blocks": [{"type": "paragraph", "text": "first"}]}, output)
    original_rmtree = preview_module.shutil.rmtree
    backup = output.with_name(f".{output.name}.old")

    def fail_backup_cleanup(path: str | Path, *args: object, **kwargs: object) -> None:
        if Path(path) == backup:
            raise RecursionError("simulated deeply nested backup")
        original_rmtree(path, *args, **kwargs)

    monkeypatch.setattr(preview_module.shutil, "rmtree", fail_backup_cleanup)
    second = render_docx_preview({"blocks": [{"type": "paragraph", "text": "second"}]}, output)

    assert second["status"] == "previewed"
    assert second["revision"] != first["revision"]
    assert "second" in (output / "document.html").read_text(encoding="utf-8")
    assert json.loads((output / "review.json").read_text(encoding="utf-8"))["revision"] == second["revision"]
    assert backup.exists()


def test_failed_old_generation_rename_is_cleaned_without_unbounded_temporary_directory(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    output = tmp_path / "review"
    render_docx_preview({"blocks": [{"type": "paragraph", "text": "first"}]}, output)
    original_replace = preview_module.os.replace
    backup = output.with_name(f".{output.name}.old")

    def fail_old_generation_rename(source: str | Path, target: str | Path) -> None:
        if Path(target) == backup:
            raise RecursionError("simulated old-generation rename failure")
        original_replace(source, target)

    monkeypatch.setattr(preview_module.os, "replace", fail_old_generation_rename)
    result = render_docx_preview({"blocks": [{"type": "paragraph", "text": "second"}]}, output)

    assert result["status"] == "previewed"
    assert "second" in (output / "document.html").read_text(encoding="utf-8")
    assert not backup.exists()
    assert [path for path in tmp_path.iterdir() if path != output] == []
