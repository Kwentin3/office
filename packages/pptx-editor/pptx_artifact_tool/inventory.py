from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Literal, TypedDict

from lxml import etree

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"p": P, "a": A}
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)
BLOCKER_FEATURES = ("ole_objects", "activex", "modification_protection", "macros", "signatures")
WARNING_FEATURES = (
    "duplicate_slot_keys",
    "group_shapes",
    "charts",
    "smartart",
    "media",
    "speaker_notes",
    "animations",
    "transitions",
    "hyperlinks",
    "external_relationships",
    "comments",
    "embedded_packages",
)


class MutationPolicy(TypedDict):
    decision: Literal["safe", "safe_with_warnings", "refuse_mutation"]
    blockers: list[str]
    warnings: list[str]


class Inventory(TypedDict):
    schema_version: int
    status: Literal["ok"]
    artifact_type: Literal["pptx"]
    view: Literal["inventory"]
    features: dict[str, int]
    mutation_policy: MutationPolicy


def _xml(archive: zipfile.ZipFile, name: str) -> etree._Element:
    return etree.fromstring(archive.read(name), parser=_PARSER)


def inspect_inventory(path: Path) -> Inventory:
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        slide_names = sorted(name for name in names if name.startswith("ppt/slides/slide") and name.endswith(".xml"))
        slide_roots = [_xml(archive, name) for name in slide_names]
        presentation = _xml(archive, "ppt/presentation.xml")
        relationship_roots = [
            _xml(archive, name)
            for name in names
            if name.startswith("ppt/") and "/_rels/" in name and name.endswith(".rels")
        ]
        comment_count = 0
        for name in names:
            if name.startswith("ppt/comments/") and name.endswith(".xml"):
                root = _xml(archive, name)
                comment_count += len(root.xpath(".//*[local-name()='cm' or local-name()='comment']"))
    count = lambda expression: sum(len(root.xpath(expression, namespaces=NS)) for root in slide_roots)
    relationship_count = lambda expression: sum(
        len(root.xpath(expression, namespaces={"pr": PR})) for root in relationship_roots
    )
    slot_names = [
        value
        for root in slide_roots
        for value in root.xpath(".//p:cNvPr[starts-with(@name, 'slot:')]/@name", namespaces=NS)
    ]
    features = {
        "slides": len(slide_names),
        "managed_slots": len(slot_names),
        "duplicate_slot_keys": len(slot_names) - len(set(slot_names)),
        "group_shapes": count(".//p:grpSp"),
        "table_shapes": count(".//a:tbl"),
        "charts": sum(name.startswith("ppt/charts/") and name.endswith(".xml") for name in names),
        "smartart": sum(name.startswith("ppt/diagrams/") and name.endswith(".xml") for name in names),
        "media": sum(name.startswith("ppt/media/") and not name.endswith("/") for name in names),
        "speaker_notes": sum(name.startswith("ppt/notesSlides/") and name.endswith(".xml") for name in names),
        "animations": count(".//p:timing"),
        "transitions": count(".//p:transition"),
        "hyperlinks": count(".//a:hlinkClick | .//a:hlinkHover"),
        "external_relationships": relationship_count("./pr:Relationship[@TargetMode='External']"),
        "comments": comment_count,
        "embedded_packages": sum(name.startswith("ppt/embeddings/") and not name.endswith("/") for name in names),
        "ole_objects": max(
            count(".//p:oleObj"), relationship_count("./pr:Relationship[contains(@Type, '/oleObject')]")
        ),
        "activex": relationship_count("./pr:Relationship[contains(@Type, '/activeX') or contains(@Type, '/control') ]"),
        "modification_protection": len(presentation.xpath(".//p:modifyVerifier", namespaces=NS)),
        "macros": int("ppt/vbaProject.bin" in names),
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
        "artifact_type": "pptx",
        "view": "inventory",
        "features": features,
        "mutation_policy": {"decision": decision, "blockers": blockers, "warnings": warnings},
    }
