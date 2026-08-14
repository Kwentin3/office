from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Literal, TypedDict

from lxml import etree

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
    mutation_policy: MutationPolicy


def _xml(archive: zipfile.ZipFile, name: str) -> etree._Element:
    return etree.fromstring(archive.read(name), parser=_PARSER)


def inspect_inventory(path: Path) -> Inventory:
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
        roots = [_xml(archive, name) for name in story_names]
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
    return {
        "schema_version": 1,
        "status": "ok",
        "artifact_type": "docx",
        "view": "inventory",
        "features": features,
        "mutation_policy": {"decision": decision, "blockers": blockers, "warnings": warnings},
    }
