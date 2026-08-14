from __future__ import annotations

import posixpath
import zipfile
from pathlib import Path, PurePosixPath
from typing import Literal, TypedDict

from lxml import etree

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"p": P, "a": A, "r": R}
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)
MAX_FINDINGS = 1000
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
    findings: list[dict[str, object]]
    findings_truncated: bool
    mutation_policy: MutationPolicy


def _xml(archive: zipfile.ZipFile, name: str) -> etree._Element:
    return etree.fromstring(archive.read(name), parser=_PARSER)


def _rels_name(part: str) -> str:
    path = PurePosixPath(part)
    return str(path.parent / "_rels" / f"{path.name}.rels")


def _source_part_from_rels(rels_part: str) -> str:
    path = PurePosixPath(rels_part)
    if path.parent.name != "_rels" or not path.name.endswith(".rels"):
        raise ValueError("validation_failure")
    return str(path.parent.parent / path.name.removesuffix(".rels"))


def _target_part(source_part: str, target: str) -> str:
    if not target or "\\" in target:
        raise ValueError("validation_failure")
    candidate = target.lstrip("/") if target.startswith("/") else posixpath.normpath(
        posixpath.join(posixpath.dirname(source_part), target)
    )
    if candidate in {"", "."} or candidate.startswith("/") or ".." in PurePosixPath(candidate).parts:
        raise ValueError("validation_failure")
    return candidate


def _ordered_slides(archive: zipfile.ZipFile) -> list[str]:
    presentation = _xml(archive, "ppt/presentation.xml")
    rels = _xml(archive, "ppt/_rels/presentation.xml.rels")
    targets = {
        relationship.get("Id"): _target_part("ppt/presentation.xml", relationship.get("Target", ""))
        for relationship in rels.findall(f"{{{PR}}}Relationship")
    }
    return [targets[rid] for rid in presentation.xpath("./p:sldIdLst/p:sldId/@r:id", namespaces=NS)]


def _shape_target(node: etree._Element) -> dict[str, object] | None:
    current: etree._Element | None = node
    shape_tags = {f"{{{P}}}sp", f"{{{P}}}graphicFrame", f"{{{P}}}grpSp", f"{{{P}}}pic"}
    while current is not None and current.tag not in shape_tags:
        current = current.getparent()
    if current is None:
        return None
    values = current.xpath(
        "./p:nvSpPr/p:cNvPr | ./p:nvGraphicFramePr/p:cNvPr | ./p:nvGrpSpPr/p:cNvPr | ./p:nvPicPr/p:cNvPr",
        namespaces=NS,
    )
    if len(values) != 1 or not values[0].get("id"):
        return None
    return {"id": int(values[0].get("id")), "name": values[0].get("name", "")}


def _relationship_shape(root: etree._Element, relationship_id: str) -> dict[str, object] | None:
    references = root.xpath(".//*[@r:id=$rid or @r:embed=$rid or @r:link=$rid]", namespaces=NS, rid=relationship_id)
    targets = [target for node in references if (target := _shape_target(node)) is not None]
    unique = {(target["id"], target["name"]): target for target in targets}
    return next(iter(unique.values())) if len(unique) == 1 else None


def _relationship_feature(relation_type: str) -> str | None:
    suffix = relation_type.rsplit("/", 1)[-1]
    if suffix == "chart":
        return "charts"
    if suffix.startswith("diagram"):
        return "smartart"
    if suffix in {"image", "audio", "video", "media"}:
        return "media"
    if suffix == "notesSlide":
        return "speaker_notes"
    if suffix in {"comment", "comments"}:
        return "comments"
    if suffix == "package":
        return "embedded_packages"
    return None


def _finding(
    feature: str,
    scope: Literal["global", "presentation", "slide", "shape", "relationship"],
    part: str,
    *,
    slide_index: int | None = None,
    slide_part: str | None = None,
    shape: dict[str, object] | None = None,
) -> dict[str, object]:
    result: dict[str, object] = {"feature": feature, "scope": scope, "part": part}
    if slide_index is not None and slide_part is not None:
        result["slide"] = {"index": slide_index, "part": slide_part}
    if shape is not None:
        result["shape"] = shape
    return result


def _finding_key(finding: dict[str, object]) -> tuple[object, ...]:
    slide = finding.get("slide") if isinstance(finding.get("slide"), dict) else {}
    shape = finding.get("shape") if isinstance(finding.get("shape"), dict) else {}
    return (
        finding["feature"],
        finding["scope"],
        finding["part"],
        slide.get("index", -1),
        slide.get("part", ""),
        shape.get("id", -1),
        shape.get("name", ""),
    )


def inspect_inventory(path: Path) -> Inventory:
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        slide_names = _ordered_slides(archive)
        slide_roots = [_xml(archive, name) for name in slide_names]
        presentation = _xml(archive, "ppt/presentation.xml")
        relationship_names = sorted(
            name for name in names if name.startswith("ppt/") and "/_rels/" in name and name.endswith(".rels")
        )
        relationship_roots = [(name, _xml(archive, name)) for name in relationship_names]
        comment_count = 0
        for name in sorted(names):
            if name.startswith("ppt/comments/") and name.endswith(".xml"):
                root = _xml(archive, name)
                comment_count += len(root.xpath(".//*[local-name()='cm' or local-name()='comment']"))

        count = lambda expression: sum(len(root.xpath(expression, namespaces=NS)) for root in slide_roots)
        relationship_count = lambda expression: sum(
            len(root.xpath(expression, namespaces={"pr": PR})) for _, root in relationship_roots
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
            "activex": relationship_count(
                "./pr:Relationship[contains(@Type, '/activeX') or contains(@Type, '/control') ]"
            ),
            "modification_protection": len(presentation.xpath(".//p:modifyVerifier", namespaces=NS)),
            "macros": int("ppt/vbaProject.bin" in names),
            "signatures": sum(name.startswith("_xmlsignatures/") and not name.endswith("/") for name in names),
        }

        findings: list[dict[str, object]] = []
        slide_by_rels: dict[str, tuple[int, str, etree._Element]] = {}
        duplicate_names = {name for name in slot_names if slot_names.count(name) > 1}
        seen_duplicate_names: set[str] = set()
        for slide_index, (slide_part, root) in enumerate(zip(slide_names, slide_roots)):
            slide = {"slide_index": slide_index, "slide_part": slide_part}
            for node in root.xpath(".//p:cNvPr[starts-with(@name, 'slot:')]", namespaces=NS):
                shape = _shape_target(node)
                if shape is not None:
                    findings.append(_finding("managed_slots", "shape", slide_part, shape=shape, **slide))
                    name = node.get("name", "")
                    if name in duplicate_names and name in seen_duplicate_names:
                        findings.append(_finding("duplicate_slot_keys", "shape", slide_part, shape=shape, **slide))
                    seen_duplicate_names.add(name)
            for node in root.xpath(".//p:grpSp", namespaces=NS):
                findings.append(_finding("group_shapes", "shape", slide_part, shape=_shape_target(node), **slide))
            for node in root.xpath(".//p:graphicFrame[.//a:tbl]", namespaces=NS):
                findings.append(_finding("table_shapes", "shape", slide_part, shape=_shape_target(node), **slide))
            for _node in root.xpath(".//p:timing", namespaces=NS):
                findings.append(_finding("animations", "slide", slide_part, **slide))
            for _node in root.xpath(".//p:transition", namespaces=NS):
                findings.append(_finding("transitions", "slide", slide_part, **slide))
            for node in root.xpath(".//a:hlinkClick | .//a:hlinkHover", namespaces=NS):
                findings.append(_finding("hyperlinks", "shape", slide_part, shape=_shape_target(node), **slide))

            rels_name = _rels_name(slide_part)
            if rels_name in names:
                slide_by_rels[rels_name] = (slide_index, slide_part, root)

            for node in root.xpath(".//p:oleObj", namespaces=NS):
                findings.append(
                    _finding("ole_objects", "global", slide_part, shape=_shape_target(node), **slide)
                )

        for _node in presentation.xpath(".//p:modifyVerifier", namespaces=NS):
            findings.append(_finding("modification_protection", "global", "ppt/presentation.xml"))
        if "ppt/vbaProject.bin" in names:
            findings.append(_finding("macros", "global", "ppt/vbaProject.bin"))
        for name in sorted(names):
            if name.startswith("_xmlsignatures/") and not name.endswith("/"):
                findings.append(_finding("signatures", "global", name))
        for rels_part, root in relationship_roots:
            source_part = _source_part_from_rels(rels_part)
            slide_location = slide_by_rels.get(rels_part)
            for relationship in root.findall(f"{{{PR}}}Relationship"):
                relation_type = relationship.get("Type", "")
                target_mode = relationship.get("TargetMode")
                if target_mode == "External":
                    if slide_location is None:
                        findings.append(_finding("external_relationships", "relationship", rels_part))
                    else:
                        slide_index, slide_part, slide_root = slide_location
                        findings.append(
                            _finding(
                                "external_relationships",
                                "relationship",
                                rels_part,
                                slide_index=slide_index,
                                slide_part=slide_part,
                                shape=_relationship_shape(slide_root, relationship.get("Id", "")),
                            )
                        )
                local_feature = _relationship_feature(relation_type)
                if local_feature is not None and target_mode != "External":
                    target = _target_part(source_part, relationship.get("Target", ""))
                    if slide_location is None:
                        findings.append(_finding(local_feature, "presentation", target))
                    else:
                        slide_index, slide_part, slide_root = slide_location
                        findings.append(
                            _finding(
                                local_feature,
                                "relationship",
                                target,
                                slide_index=slide_index,
                                slide_part=slide_part,
                                shape=_relationship_shape(slide_root, relationship.get("Id", "")),
                            )
                        )
                feature = None
                if relation_type.endswith("/oleObject"):
                    feature = "ole_objects"
                elif relation_type.endswith(("/activeX", "/activeXControl", "/control")):
                    feature = "activex"
                if feature is not None:
                    target = _target_part(source_part, relationship.get("Target", ""))
                    findings.append(_finding(feature, "global", target))

        findings.sort(key=_finding_key)
        truncated = len(findings) > MAX_FINDINGS
        findings = findings[:MAX_FINDINGS]

    blockers = [name for name in BLOCKER_FEATURES if features[name]]
    warnings = [name for name in WARNING_FEATURES if features[name]]
    decision: Literal["safe", "safe_with_warnings", "refuse_mutation"] = (
        "refuse_mutation" if blockers else "safe_with_warnings" if warnings else "safe"
    )
    return {
        "schema_version": 2,
        "status": "ok",
        "artifact_type": "pptx",
        "view": "inventory",
        "features": features,
        "findings": findings,
        "findings_truncated": truncated,
        "mutation_policy": {"decision": decision, "blockers": blockers, "warnings": warnings},
    }
