from __future__ import annotations

import posixpath
import re
import zipfile
from pathlib import Path, PurePosixPath
from typing import Literal, TypedDict

from lxml import etree

S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"s": S, "r": R, "pr": PR}
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)
_FINDING_LIMIT = 1000
BLOCKER_FEATURES = (
    "external_links",
    "external_relationships",
    "connections",
    "ole_objects",
    "activex",
    "workbook_protection",
    "sheet_protection",
    "macros",
    "signatures",
)
WARNING_FEATURES = (
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


class Finding(TypedDict, total=False):
    feature: str
    scope: Literal["global", "workbook", "worksheet", "member"]
    part: str
    sheet: str
    range: str


class Inventory(TypedDict):
    schema_version: int
    status: Literal["ok"]
    artifact_type: Literal["xlsx"]
    view: Literal["inventory"]
    features: dict[str, int]
    findings: list[Finding]
    findings_truncated: bool
    mutation_policy: MutationPolicy


def _xml(archive: zipfile.ZipFile, name: str) -> etree._Element:
    return etree.fromstring(archive.read(name), parser=_PARSER)


def _part(source_part: str, target: str) -> str:
    if target.startswith("/"):
        normalized = target.lstrip("/")
    else:
        normalized = posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))
    path = PurePosixPath(normalized)
    if normalized.startswith("/") or ".." in path.parts:
        raise ValueError("unsafe_package")
    return str(path)


def _rels_owner(name: str) -> str | None:
    marker = "/_rels/"
    if marker not in name or not name.endswith(".rels"):
        return None
    prefix, leaf = name.split(marker, 1)
    return f"{prefix}/{leaf[:-5]}"


def _relationship_index(archive: zipfile.ZipFile, names: set[str]):
    by_source: dict[str, list[dict[str, str]]] = {}
    rel_roots: list[tuple[str, etree._Element]] = []
    for rels_name in sorted(name for name in names if name.endswith(".rels")):
        root = _xml(archive, rels_name)
        rel_roots.append((rels_name, root))
        owner = _rels_owner(rels_name)
        if owner is None:
            continue
        relationships = []
        for node in root.xpath("./pr:Relationship", namespaces=NS):
            target = node.get("Target", "")
            relationships.append(
                {
                    "id": node.get("Id", ""),
                    "type": node.get("Type", ""),
                    "target": target if node.get("TargetMode") == "External" else _part(owner, target),
                    "mode": node.get("TargetMode", "Internal"),
                    "rels_part": rels_name,
                }
            )
        by_source[owner] = relationships
    return by_source, rel_roots


def _simple_defined_name(value: str, sheet_names: set[str]) -> tuple[str, str] | None:
    match = re.fullmatch(r"(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)", value)
    if not match:
        return None
    sheet = (match.group(1) or match.group(2)).replace("''", "'")
    if sheet not in sheet_names:
        return None
    return sheet, match.group(3).replace("$", "")


def _drawing_range(root: etree._Element) -> str | None:
    anchors = root.xpath("./*[local-name()='oneCellAnchor' or local-name()='twoCellAnchor']")
    ranges: list[str] = []
    for anchor in anchors:
        start = anchor.xpath("./*[local-name()='from'][1]")
        if not start:
            continue
        col = start[0].xpath("string(./*[local-name()='col'])")
        row = start[0].xpath("string(./*[local-name()='row'])")
        if not col.isdigit() or not row.isdigit():
            continue
        # Avoid importing openpyxl merely for inventory address conversion.
        number = int(col) + 1
        letters = ""
        while number:
            number, remainder = divmod(number - 1, 26)
            letters = chr(65 + remainder) + letters
        first = f"{letters}{int(row) + 1}"
        end = anchor.xpath("./*[local-name()='to'][1]")
        if end:
            end_col = end[0].xpath("string(./*[local-name()='col'])")
            end_row = end[0].xpath("string(./*[local-name()='row'])")
            if end_col.isdigit() and end_row.isdigit():
                number = int(end_col) + 1
                letters = ""
                while number:
                    number, remainder = divmod(number - 1, 26)
                    letters = chr(65 + remainder) + letters
                last = f"{letters}{int(end_row) + 1}"
                first = first if first == last else f"{first}:{last}"
        ranges.append(first)
    return " ".join(ranges) if ranges else None


def inspect_inventory(path: Path) -> Inventory:
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        workbook = _xml(archive, "xl/workbook.xml")
        relationships, relationship_roots = _relationship_index(archive, names)
        workbook_rels = {item["id"]: item for item in relationships.get("xl/workbook.xml", [])}
        sheets: dict[str, str] = {}
        for node in workbook.xpath("./s:sheets/s:sheet", namespaces=NS):
            relation = workbook_rels.get(node.get(f"{{{R}}}id", ""))
            if relation and relation["target"] in names:
                sheets[relation["target"]] = node.get("name", "")
        worksheet_names = sorted(sheets)
        worksheet_roots = {name: _xml(archive, name) for name in worksheet_names}
        findings: list[Finding] = []

        def add(feature: str, scope: str, part: str, sheet: str | None = None, range_ref: str | None = None):
            finding: Finding = {"feature": feature, "scope": scope, "part": part}  # type: ignore[typeddict-item]
            if sheet is not None:
                finding["sheet"] = sheet
            if range_ref:
                finding["range"] = range_ref
            findings.append(finding)

        def worksheet_nodes(feature: str, expression: str, range_attribute: str | None = None) -> int:
            total = 0
            for part in worksheet_names:
                nodes = worksheet_roots[part].xpath(expression, namespaces=NS)
                total += len(nodes)
                for node in nodes:
                    add(feature, "worksheet", part, sheets[part], node.get(range_attribute) if range_attribute else None)
            return total

        comment_count = 0
        for sheet_part in worksheet_names:
            sheet_name = sheets[sheet_part]
            for relation in relationships.get(sheet_part, []):
                relation_type = relation["type"]
                related_part = relation["target"]
                if relation_type.endswith("/comments") and related_part in names:
                    root = _xml(archive, related_part)
                    nodes = root.xpath(".//*[local-name()='comment' or local-name()='threadedComment']")
                    comment_count += len(nodes)
                    for node in nodes:
                        add("comments", "worksheet", related_part, sheet_name, node.get("ref"))
                elif relation_type.endswith("/table") and related_part in names:
                    root = _xml(archive, related_part)
                    add("tables", "worksheet", related_part, sheet_name, root.get("ref"))
                elif relation_type.endswith("/pivotTable") and related_part in names:
                    root = _xml(archive, related_part)
                    location = root.xpath("./s:location/@ref", namespaces=NS)
                    add("pivots", "worksheet", related_part, sheet_name, location[0] if location else None)
                elif relation_type.endswith("/drawing") and related_part in names:
                    drawing = _xml(archive, related_part)
                    drawing_range = _drawing_range(drawing)
                    add("drawings", "worksheet", related_part, sheet_name, drawing_range)
                    for drawing_relation in relationships.get(related_part, []):
                        if drawing_relation["type"].endswith("/chart"):
                            add("charts", "worksheet", drawing_relation["target"], sheet_name, drawing_range)

        data_validations = worksheet_nodes("data_validations", ".//s:dataValidation", "sqref")
        conditional_formatting = worksheet_nodes("conditional_formatting", ".//s:conditionalFormatting", "sqref")
        merged_ranges = worksheet_nodes("merged_ranges", ".//s:mergeCell", "ref")
        formula_cells = worksheet_nodes("formula_cells", ".//s:c[s:f]", "r")
        shared_formulas = worksheet_nodes("shared_formulas", ".//s:f[@t='shared']", "ref")
        array_formulas = worksheet_nodes("array_formulas", ".//s:f[@t='array']", "ref")
        sheet_protection = worksheet_nodes("sheet_protection", ".//s:sheetProtection")
        # Protection is intentionally global policy even when its physical node is in one worksheet.
        for finding in findings:
            if finding["feature"] == "sheet_protection":
                finding["scope"] = "global"
                finding.pop("sheet", None)
                finding.pop("range", None)

        defined_name_nodes = workbook.xpath("./s:definedNames/s:definedName", namespaces=NS)
        sheet_names = set(sheets.values())
        for node in defined_name_nodes:
            location = _simple_defined_name(node.text or "", sheet_names)
            if location:
                add("defined_names", "worksheet", "xl/workbook.xml", location[0], location[1])
            else:
                add("defined_names", "workbook", "xl/workbook.xml")

        protection_nodes = workbook.xpath(".//s:workbookProtection | .//s:fileSharing", namespaces=NS)
        active_workbook_protection = sum(
            any(str(value).casefold() not in {"", "0", "false"} for value in node.attrib.values())
            for node in protection_nodes
        )
        for _ in range(active_workbook_protection):
            add("workbook_protection", "global", "xl/workbook.xml")

        external_links = 0
        external_relationships = 0
        activex = 0
        ole_relationships = 0
        for rels_part, root in relationship_roots:
            for node in root.xpath("./pr:Relationship", namespaces=NS):
                relation_type = node.get("Type", "")
                if relation_type.endswith("/externalLink"):
                    external_links += 1
                    add("external_links", "global", rels_part)
                if node.get("TargetMode") == "External":
                    external_relationships += 1
                    add("external_relationships", "global", rels_part)
                if relation_type.endswith(("/activeX", "/control")):
                    activex += 1
                    add("activex", "global", rels_part)
                if relation_type.endswith("/oleObject"):
                    ole_relationships += 1
                    add("ole_objects", "global", rels_part)

        ole_nodes = sum(len(root.xpath(".//*[local-name()='oleObject']")) for root in worksheet_roots.values())
        if ole_nodes > ole_relationships:
            for part in worksheet_names:
                for _ in worksheet_roots[part].xpath(".//*[local-name()='oleObject']"):
                    add("ole_objects", "global", part)
        ole_objects = max(ole_nodes, ole_relationships)

        charts = sum(name.startswith("xl/charts/") and name.endswith(".xml") for name in names)
        pivots = sum(name.startswith("xl/pivot") and name.endswith(".xml") for name in names)
        tables = sum(name.startswith("xl/tables/") and name.endswith(".xml") for name in names)
        drawings = sum(name.startswith("xl/drawings/") and name.endswith(".xml") for name in names)
        # Preserve total counts even for malformed/unrelated graph fragments; add member-scoped fallbacks.
        for feature, prefix, count_value in (
            ("charts", "xl/charts/", charts),
            ("pivots", "xl/pivot", pivots),
            ("tables", "xl/tables/", tables),
            ("drawings", "xl/drawings/", drawings),
        ):
            located = {finding["part"] for finding in findings if finding["feature"] == feature}
            candidates = sorted(name for name in names if name.startswith(prefix) and name.endswith(".xml"))
            for part in candidates:
                if part not in located:
                    add(feature, "member", part)

        connections = int("xl/connections.xml" in names)
        if connections:
            add("connections", "global", "xl/connections.xml")
        embedded_members = sorted(name for name in names if name.startswith("xl/embeddings/") and not name.endswith("/"))
        for part in embedded_members:
            add("embedded_packages", "global", part)
        macros = int("xl/vbaProject.bin" in names)
        if macros:
            add("macros", "global", "xl/vbaProject.bin")
        signature_members = sorted(name for name in names if name.startswith("_xmlsignatures/") and not name.endswith("/"))
        for part in signature_members:
            add("signatures", "global", part)

        features = {
            "external_links": external_links,
            "external_relationships": external_relationships,
            "charts": charts,
            "pivots": pivots,
            "tables": tables,
            "comments": comment_count,
            "data_validations": data_validations,
            "conditional_formatting": conditional_formatting,
            "defined_names": len(defined_name_nodes),
            "merged_ranges": merged_ranges,
            "formula_cells": formula_cells,
            "shared_formulas": shared_formulas,
            "array_formulas": array_formulas,
            "drawings": drawings,
            "connections": connections,
            "embedded_packages": len(embedded_members),
            "ole_objects": ole_objects,
            "activex": activex,
            "workbook_protection": active_workbook_protection,
            "sheet_protection": sheet_protection,
            "macros": macros,
            "signatures": len(signature_members),
        }

    blockers = [name for name in BLOCKER_FEATURES if features[name]]
    warnings = [name for name in WARNING_FEATURES if features[name]]
    decision: Literal["safe", "safe_with_warnings", "refuse_mutation"] = (
        "refuse_mutation" if blockers else "safe_with_warnings" if warnings else "safe"
    )
    findings.sort(key=lambda item: (item["feature"], item["scope"], item["part"], item.get("sheet", ""), item.get("range", "")))
    truncated = len(findings) > _FINDING_LIMIT
    return {
        "schema_version": 2,
        "status": "ok",
        "artifact_type": "xlsx",
        "view": "inventory",
        "features": features,
        "findings": findings[:_FINDING_LIMIT],
        "findings_truncated": truncated,
        "mutation_policy": {"decision": decision, "blockers": blockers, "warnings": warnings},
    }
