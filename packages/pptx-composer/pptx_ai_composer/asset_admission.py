"""Shared immutable asset-admission boundary for PPTX outputs."""

from __future__ import annotations

import hashlib
import io
import os
import stat
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from lxml import etree
from PIL import Image, UnidentifiedImageError

_MAX_ASSET_BYTES = 20 * 1024 * 1024
_MAX_RASTER_WIDTH = 8192
_MAX_RASTER_HEIGHT = 8192
_MAX_RASTER_PIXELS = 40_000_000
_EXPECTED_FORMAT = {"png": "PNG", "jpeg": "JPEG"}


class AssetAdmissionError(ValueError):
    """An asset record does not authenticate to safe, bounded local bytes."""


@dataclass(frozen=True, slots=True)
class AdmittedAsset:
    """Private immutable byte snapshot consumed by rendering backends."""

    asset_id: str
    kind: str
    source_bytes: bytes
    raster_bytes: bytes


def _snapshot_bytes(path_value: str, digest: str, *, kind: str, asset_id: str) -> bytes:
    path = Path(path_value)
    flags = (
        os.O_RDONLY
        | getattr(os, "O_NOFOLLOW", 0)
        | getattr(os, "O_CLOEXEC", 0)
        | getattr(os, "O_NONBLOCK", 0)
    )
    descriptor = -1
    try:
        descriptor = os.open(path, flags)
        before = os.fstat(descriptor)
        if not stat.S_ISREG(before.st_mode):
            raise AssetAdmissionError(f"{kind} is missing or unsafe: {asset_id}")
        if before.st_size > _MAX_ASSET_BYTES:
            raise AssetAdmissionError(f"{kind} exceeds byte limit: {asset_id}")
        with os.fdopen(descriptor, "rb") as stream:
            descriptor = -1
            payload = stream.read(_MAX_ASSET_BYTES + 1)
            after = os.fstat(stream.fileno())
    except AssetAdmissionError:
        raise
    except OSError as exc:
        raise AssetAdmissionError(f"{kind} is missing or unsafe: {asset_id}") from exc
    finally:
        if descriptor >= 0:
            os.close(descriptor)
    if len(payload) > _MAX_ASSET_BYTES:
        raise AssetAdmissionError(f"{kind} exceeds byte limit: {asset_id}")
    before_identity = (before.st_dev, before.st_ino, before.st_size, before.st_mtime_ns)
    after_identity = (after.st_dev, after.st_ino, after.st_size, after.st_mtime_ns)
    if before_identity != after_identity or len(payload) != before.st_size:
        raise AssetAdmissionError(f"{kind} changed during admission: {asset_id}")
    if hashlib.sha256(payload).hexdigest() != digest:
        raise AssetAdmissionError(f"{kind} hash mismatch: {asset_id}")
    return payload


def _validate_raster(payload: bytes, expected_kind: str, *, kind: str, asset_id: str) -> None:
    expected_format = _EXPECTED_FORMAT[expected_kind]
    try:
        with Image.open(io.BytesIO(payload)) as image:
            if image.format != expected_format:
                raise AssetAdmissionError(f"raster format mismatch: {asset_id}")
            width, height = image.size
            if width < 1 or height < 1:
                raise AssetAdmissionError(f"invalid raster dimensions: {asset_id}")
            if width > _MAX_RASTER_WIDTH or height > _MAX_RASTER_HEIGHT:
                raise AssetAdmissionError(f"raster dimensions exceed limit: {asset_id}")
            if width * height > _MAX_RASTER_PIXELS:
                raise AssetAdmissionError(f"raster pixel budget exceeded: {asset_id}")
            image.load()
    except AssetAdmissionError:
        raise
    except (Image.DecompressionBombError, UnidentifiedImageError, OSError, SyntaxError, ValueError) as exc:
        raise AssetAdmissionError(f"invalid raster {kind}: {asset_id}") from exc


def _validate_svg(payload: bytes, asset_id: str) -> None:
    try:
        root = etree.fromstring(
            payload,
            parser=etree.XMLParser(resolve_entities=False, no_network=True, recover=False),
        )
    except Exception as exc:
        raise AssetAdmissionError(f"unsafe svg: {asset_id}: {exc}") from exc
    for element in root.iter():
        local_name = etree.QName(element).localname
        if local_name in {"script", "foreignObject", "iframe", "object", "embed"}:
            raise AssetAdmissionError(f"unsafe svg: forbidden {local_name}")
        for name, value in element.attrib.items():
            attr_name = etree.QName(name).localname.lower()
            lowered = value.strip().lower()
            if attr_name.startswith("on") or lowered.startswith(("javascript:", "data:text/html")):
                raise AssetAdmissionError("unsafe svg: active attribute")
            if attr_name in {"href", "src"} and not lowered.startswith(("#", "data:image/")):
                raise AssetAdmissionError("unsafe svg: external reference")


def admit_assets(deck: dict[str, Any]) -> dict[str, AdmittedAsset]:
    """Authenticate declared assets into private immutable byte snapshots."""
    assets: dict[str, AdmittedAsset] = {}
    for asset in deck["assets"]:
        asset_id = asset["asset_id"]
        source_bytes = _snapshot_bytes(
            asset["path"],
            asset["sha256"],
            kind="asset",
            asset_id=asset_id,
        )
        if asset["kind"] == "svg":
            _validate_svg(source_bytes, asset_id)
            raster_bytes = _snapshot_bytes(
                asset["fallback_path"],
                asset["fallback_sha256"],
                kind="asset fallback",
                asset_id=asset_id,
            )
            _validate_raster(raster_bytes, "png", kind="asset fallback", asset_id=asset_id)
        else:
            raster_bytes = source_bytes
            _validate_raster(raster_bytes, asset["kind"], kind="asset", asset_id=asset_id)
        assets[asset_id] = AdmittedAsset(
            asset_id=asset_id,
            kind=asset["kind"],
            source_bytes=source_bytes,
            raster_bytes=raster_bytes,
        )
    return assets
