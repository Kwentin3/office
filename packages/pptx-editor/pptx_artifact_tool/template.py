from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree

TOKEN = re.compile(r"\{\{([A-Za-z][A-Za-z0-9_.-]{0,79})\}\}")
MAX_RENDERED_LENGTH = 32767


def _parse_xml(payload: bytes):
    if b"<!DOCTYPE" in payload or b"<!ENTITY" in payload:
        raise ValueError("validation_failure")
    parser = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)
    root = etree.fromstring(payload, parser=parser)
    document_info = root.getroottree().docinfo
    if document_info.doctype or document_info.internalDTD is not None or document_info.externalDTD is not None:
        raise ValueError("validation_failure")
    return root


def validate_values(values: Any) -> dict[str, str]:
    if not isinstance(values, dict) or not values:
        raise ValueError("validation_failure")
    result: dict[str, str] = {}
    for key, value in values.items():
        if not isinstance(key, str) or TOKEN.fullmatch("{{" + key + "}}") is None:
            raise ValueError("validation_failure")
        if not isinstance(value, str) or len(value) > MAX_RENDERED_LENGTH or "{{" in value or "}}" in value:
            raise ValueError("validation_failure")
        result[key] = value
    return result


def token_names(text: str) -> set[str]:
    return set(TOKEN.findall(text))


def has_marker(text: str) -> bool:
    return "{{" in text or "}}" in text


def package_has_marker(source: Path) -> bool:
    with zipfile.ZipFile(source) as archive:
        for name in archive.namelist():
            if not name.endswith((".xml", ".rels")):
                continue
            root = _parse_xml(archive.read(name))
            pieces = [value for node in root.iter() for value in node.attrib.values()]
            pieces.extend(piece for node in root.iter() for piece in (node.text, node.tail) if piece)
            if has_marker("".join(pieces)):
                return True
    return False


def well_formed(text: str) -> bool:
    covered = {index for match in TOKEN.finditer(text) for index in range(*match.span())}
    for marker in ("{{", "}}"):
        start = text.find(marker)
        while start >= 0:
            if start not in covered or start + 1 not in covered:
                return False
            start = text.find(marker, start + 1)
    return True


def render(text: str, values: dict[str, str]) -> str:
    if not well_formed(text):
        raise ValueError("validation_failure")
    result = TOKEN.sub(lambda match: values[match.group(1)], text)
    if len(result) > MAX_RENDERED_LENGTH:
        raise ValueError("validation_failure")
    return result


def unsupported_package_scope_has_marker(source: Path, slide_members: set[str]) -> bool:
    with zipfile.ZipFile(source) as archive:
        for name in archive.namelist():
            if not name.endswith((".xml", ".rels")) or name in slide_members:
                continue
            root = _parse_xml(archive.read(name))
            pieces: list[str] = []
            for node in root.iter():
                pieces.extend(node.attrib.values())
                if node.text:
                    pieces.append(node.text)
                if node.tail:
                    pieces.append(node.tail)
            if has_marker("".join(pieces)):
                return True
    return False


def unmanaged_slide_scope_has_marker(source: Path, slide_members: set[str]) -> bool:
    namespaces = {
        "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    }
    shape_tags = {
        "{http://schemas.openxmlformats.org/presentationml/2006/main}sp",
        "{http://schemas.openxmlformats.org/presentationml/2006/main}graphicFrame",
    }
    with zipfile.ZipFile(source) as archive:
        for name in slide_members:
            root = _parse_xml(archive.read(name))
            if any(has_marker(value) for node in root.iter() for value in node.attrib.values()):
                return True
            shapes = root.xpath(".//p:sp | .//p:graphicFrame", namespaces=namespaces)
            for shape in shapes:
                text_nodes = shape.xpath(".//a:t", namespaces=namespaces)
                if not has_marker("".join(node.text or "" for node in text_nodes)):
                    continue
                if shape.tag not in shape_tags or shape.getparent().tag != f"{{{namespaces['p']}}}spTree":
                    return True
                names = shape.xpath(
                    "./p:nvSpPr/p:cNvPr/@name | ./p:nvGraphicFramePr/p:cNvPr/@name",
                    namespaces=namespaces,
                )
                if len(names) != 1 or not names[0].startswith("slot:"):
                    return True
            outside = "".join(
                node.text or ""
                for node in root.xpath(
                    ".//a:t[not(ancestor::p:sp) and not(ancestor::p:graphicFrame)]",
                    namespaces=namespaces,
                )
            )
            if has_marker(outside):
                return True
    return False
