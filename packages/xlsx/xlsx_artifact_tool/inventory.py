from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Literal, TypedDict

from lxml import etree

S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"s": S}
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)
BLOCKER_FEATURES = (
    "external_links",
    "connections",
    "ole_objects",
    "activex",
    "workbook_protection",
    "sheet_protection",
    "macros",
    "signatures",
)
WARNING_FEATURES = (
    "external_relationships",
    "charts",
    "pivots",
    "tables",
    "comments",
    "data_validations",
    "conditional_formatting",
    "defined_names",
    "merged_ranges",
    "shared_formulas",
    "array_formulas",
    "drawings",
    "embedded_packages",
)


class MutationPolicy(TypedDict):
    decision: Literal["safe", "safe_with_warnings", "refuse_mutation"]
    blockers: list[str]
    warnings: list[str]


class Inventory(TypedDict):
    schema_version: int
    status: Literal["ok"]
    artifact_type: Literal["xlsx"]
    view: Literal["inventory"]
    features: dict[str, int]
    mutation_policy: MutationPolicy


def _xml(archive: zipfile.ZipFile, name: str) -> etree._Element:
    return etree.fromstring(archive.read(name), parser=_PARSER)


def inspect_inventory(path: Path) -> Inventory:
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        worksheet_names = sorted(name for name in names if name.startswith("xl/worksheets/") and name.endswith(".xml"))
        worksheet_roots = [_xml(archive, name) for name in worksheet_names]
        workbook = _xml(archive, "xl/workbook.xml")
        relationship_roots = [
            _xml(archive, name)
            for name in names
            if name.startswith("xl/") and "/_rels/" in name and name.endswith(".rels")
        ]
        comment_count = 0
        for name in names:
            if name.startswith(("xl/comments", "xl/threadedComments/")) and name.endswith(".xml"):
                root = _xml(archive, name)
                comment_count += len(root.xpath(".//*[local-name()='comment' or local-name()='threadedComment']"))
    count = lambda expression: sum(len(root.xpath(expression, namespaces=NS)) for root in worksheet_roots)
    relationship_count = lambda expression: sum(
        len(root.xpath(expression, namespaces={"pr": PR})) for root in relationship_roots
    )
    protection_nodes = workbook.xpath(".//s:workbookProtection | .//s:fileSharing", namespaces=NS)
    active_workbook_protection = sum(
        any(str(value).casefold() not in {"", "0", "false"} for value in node.attrib.values())
        for node in protection_nodes
    )
    features = {
        "external_links": relationship_count("./pr:Relationship[contains(@Type, '/externalLink')]"),
        "external_relationships": relationship_count("./pr:Relationship[@TargetMode='External']"),
        "charts": sum(name.startswith("xl/charts/") and name.endswith(".xml") for name in names),
        "pivots": sum(name.startswith("xl/pivot") and name.endswith(".xml") for name in names),
        "tables": sum(name.startswith("xl/tables/") and name.endswith(".xml") for name in names),
        "comments": comment_count,
        "data_validations": count(".//s:dataValidation"),
        "conditional_formatting": count(".//s:conditionalFormatting"),
        "defined_names": len(workbook.xpath("./s:definedNames/s:definedName", namespaces=NS)),
        "merged_ranges": count(".//s:mergeCell"),
        "formula_cells": count(".//s:c[s:f]"),
        "shared_formulas": count(".//s:f[@t='shared']"),
        "array_formulas": count(".//s:f[@t='array']"),
        "drawings": sum(name.startswith("xl/drawings/") and name.endswith(".xml") for name in names),
        "connections": int("xl/connections.xml" in names),
        "embedded_packages": sum(name.startswith("xl/embeddings/") and not name.endswith("/") for name in names),
        "ole_objects": max(
            count(".//*[local-name()='oleObject']"),
            relationship_count("./pr:Relationship[contains(@Type, '/oleObject')]"),
        ),
        "activex": relationship_count("./pr:Relationship[contains(@Type, '/activeX') or contains(@Type, '/control') ]"),
        "workbook_protection": active_workbook_protection,
        "sheet_protection": count(".//s:sheetProtection"),
        "macros": int("xl/vbaProject.bin" in names),
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
        "artifact_type": "xlsx",
        "view": "inventory",
        "features": features,
        "mutation_policy": {"decision": decision, "blockers": blockers, "warnings": warnings},
    }
