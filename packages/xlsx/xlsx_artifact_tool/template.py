from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree

S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
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


def unsupported_scope_has_marker(source: Path) -> bool:
    with zipfile.ZipFile(source) as archive:
        for name in archive.namelist():
            if not name.endswith((".xml", ".rels")):
                continue
            if name == "xl/sharedStrings.xml":
                continue
            root = _parse_xml(archive.read(name))
            pieces: list[str] = []
            if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                for node in root.iter():
                    if any(ancestor.tag == f"{{{S}}}c" for ancestor in node.iterancestors()):
                        continue
                    if node.tag == f"{{{S}}}c":
                        continue
                    pieces.extend(node.attrib.values())
                    if node.text:
                        pieces.append(node.text)
                    if node.tail:
                        pieces.append(node.tail)
            else:
                for node in root.iter():
                    pieces.extend(node.attrib.values())
                    if node.text:
                        pieces.append(node.text)
                    if node.tail:
                        pieces.append(node.tail)
            if has_marker("".join(pieces)):
                return True
    return False
