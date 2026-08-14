from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Literal, TypedDict

from lxml import etree

from ..core.hashes import file_sha256
from .inspect import _id, _style, _text

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W}
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)
BLOCKER_FEATURES = (
    "tracked_changes",
    "content_controls",
    "fields",
    "alt_chunks",
    "ole_objects",
    "activex",
    "document_protection",
    "macros",
    "signatures",
)
WARNING_FEATURES = (
    "comments",
    "images",
    "text_boxes",
    "merged_cells",
    "hyperlinks",
    "external_relationships",
    "charts",
    "smartart",
    "embedded_packages",
)
FINDINGS_LIMIT = 1000
GLOBAL_BLOCKER_FEATURES = ("ole_objects", "activex", "document_protection", "macros", "signatures")


class MutationPolicy(TypedDict):
    decision: Literal["safe", "safe_with_warnings", "refuse_mutation"]
    blockers: list[str]
    warnings: list[str]


class Inventory(TypedDict):
    schema_version: int
    status: Literal["ok"]
    artifact_type: Literal["docx"]
    view: Literal["inventory"]
    features: dict[str, int]
    findings: list[dict[str, str]]
    findings_truncated: bool
    mutation_policy: MutationPolicy


def _xml(archive: zipfile.ZipFile, name: str) -> etree._Element:
    return etree.fromstring(archive.read(name), parser=_PARSER)


def _target_id(source: str, part: str, root: etree._Element, node: etree._Element) -> str | None:
    cells = node.xpath("ancestor-or-self::w:tc[1]", namespaces=NS)
    if cells:
        cell = cells[0]
        tables = root.xpath(".//w:tbl", namespaces=NS)
        table_ancestors = cell.xpath("ancestor::w:tbl[1]", namespaces=NS)
        if not table_ancestors:
            return None
        table = table_ancestors[0]
        table_index = next((index for index, item in enumerate(tables) if item is table), None)
        rows = table.xpath("./w:tr", namespaces=NS)
        row_ancestors = cell.xpath("ancestor::w:tr[1]", namespaces=NS)
        if table_index is None or not row_ancestors:
            return None
        row = row_ancestors[0]
        row_index = next((index for index, item in enumerate(rows) if item is row), None)
        row_cells = row.xpath("./w:tc", namespaces=NS)
        cell_index = next((index for index, item in enumerate(row_cells) if item is cell), None)
        if row_index is None or cell_index is None:
            return None
        return _id(source, part, "cell", f"{table_index}/{row_index}/{cell_index}", _text(cell))
    paragraphs = root.xpath(".//w:p[not(ancestor::w:tc)]", namespaces=NS)
    ancestors = node.xpath("ancestor-or-self::w:p[not(ancestor::w:tc)][1]", namespaces=NS)
    if not ancestors:
        return None
    paragraph = ancestors[0]
    text = _text(paragraph)
    if not text:
        return None
    index = next((index for index, item in enumerate(paragraphs) if item is paragraph), None)
    if index is None:
        return None
    kind = "heading" if _style(paragraph).lower().startswith("heading") else "paragraph"
    return _id(source, part, kind, str(index), text)


def mutation_blockers_for_story_parts(inventory: Inventory, story_parts: set[str]) -> list[str]:
    """Return blockers that affect exact mutation parts, failing closed when findings were truncated."""
    blocked = []
    if inventory["findings_truncated"]:
        return [feature for feature in BLOCKER_FEATURES if inventory["features"][feature]]
    for feature in BLOCKER_FEATURES:
        count = inventory["features"][feature]
        if not count:
            continue
        if feature in GLOBAL_BLOCKER_FEATURES:
            blocked.append(feature)
            continue
        findings = [item for item in inventory["findings"] if item["feature"] == feature]
        if len(findings) < count or any(item["scope"] != "story" or item["part"] in story_parts for item in findings):
            blocked.append(feature)
    return blocked


def inspect_inventory(path: Path) -> Inventory:
    source = file_sha256(path)
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        story_names = sorted(
            name
            for name in names
            if name == "word/document.xml"
            or name.startswith("word/header")
            and name.endswith(".xml")
            or name.startswith("word/footer")
            and name.endswith(".xml")
            or name in {"word/footnotes.xml", "word/endnotes.xml"}
        )
        roots_by_part = {name: _xml(archive, name) for name in story_names}
        roots = list(roots_by_part.values())
        relationship_roots = [
            _xml(archive, name)
            for name in names
            if name.startswith("word/") and "/_rels/" in name and name.endswith(".rels")
        ]
        settings = _xml(archive, "word/settings.xml") if "word/settings.xml" in names else None
        comment_count = 0
        if "word/comments.xml" in names:
            comments = _xml(archive, "word/comments.xml")
            comment_count = len(comments.xpath("./w:comment", namespaces=NS))
    count = lambda expression: sum(len(root.xpath(expression, namespaces=NS)) for root in roots)
    relationship_count = lambda expression: sum(
        len(root.xpath(expression, namespaces={"pr": PR})) for root in relationship_roots
    )
    ole_relationships = relationship_count("./pr:Relationship[contains(@Type, '/oleObject')]")
    features = {
        "tracked_changes": count(".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo"),
        "content_controls": count(".//w:sdt"),
        "comments": comment_count,
        "images": sum(name.startswith("word/media/") and not name.endswith("/") for name in names),
        "fields": count(".//w:fldSimple | .//w:instrText"),
        "alt_chunks": count(".//w:altChunk"),
        "text_boxes": count(".//w:txbxContent"),
        "merged_cells": count(".//w:gridSpan | .//w:vMerge"),
        "hyperlinks": count(".//w:hyperlink"),
        "external_relationships": relationship_count("./pr:Relationship[@TargetMode='External']"),
        "headers": sum(name.startswith("word/header") and name.endswith(".xml") for name in names),
        "footers": sum(name.startswith("word/footer") and name.endswith(".xml") for name in names),
        "footnotes": count("self::w:footnotes/w:footnote"),
        "endnotes": count("self::w:endnotes/w:endnote"),
        "charts": sum(name.startswith("word/charts/") and name.endswith(".xml") for name in names),
        "smartart": sum(name.startswith("word/diagrams/") and name.endswith(".xml") for name in names),
        "embedded_packages": sum(name.startswith("word/embeddings/") and not name.endswith("/") for name in names),
        "ole_objects": max(count(".//*[local-name()='OLEObject']"), ole_relationships),
        "activex": relationship_count("./pr:Relationship[contains(@Type, '/activeX') or contains(@Type, '/control') ]"),
        "document_protection": 0 if settings is None else len(settings.xpath(".//w:documentProtection", namespaces=NS)),
        "macros": int("word/vbaProject.bin" in names),
        "signatures": sum(name.startswith("_xmlsignatures/") and not name.endswith("/") for name in names),
    }
    blockers = [name for name in BLOCKER_FEATURES if features[name]]
    warnings = [name for name in WARNING_FEATURES if features[name]]
    decision: Literal["safe", "safe_with_warnings", "refuse_mutation"] = (
        "refuse_mutation" if blockers else "safe_with_warnings" if warnings else "safe"
    )
    findings = []
    for part, root in roots_by_part.items():
        for feature, expression in (
            ("tracked_changes", ".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo"),
            ("hyperlinks", ".//w:hyperlink"),
            ("merged_cells", ".//w:gridSpan | .//w:vMerge"),
        ):
            for node in root.xpath(expression, namespaces=NS):
                finding = {"feature": feature, "scope": "story", "part": part}
                target_id = _target_id(source, part, root, node)
                if target_id is not None:
                    finding["target_id"] = target_id
                findings.append(finding)
    findings = sorted(
        findings,
        key=lambda item: tuple(item.get(key, "") for key in ("feature", "scope", "part", "target_id")),
    )
    findings_truncated = len(findings) > FINDINGS_LIMIT
    return {
        "schema_version": 2,
        "status": "ok",
        "artifact_type": "docx",
        "view": "inventory",
        "features": features,
        "findings": findings[:FINDINGS_LIMIT],
        "findings_truncated": findings_truncated,
        "mutation_policy": {"decision": decision, "blockers": blockers, "warnings": warnings},
    }
