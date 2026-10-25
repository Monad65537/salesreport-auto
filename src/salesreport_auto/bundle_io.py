"""Read-only bounded bundle snapshots and strict JSON, with no path discovery."""

import hashlib
import json
import os
import stat
from decimal import Decimal
from pathlib import Path

from .bundle_models import BUNDLE_MEMBERS, SheetMergeBundle
from .errors import BundleError

MEMBER_LIMIT = 20 * 1024 * 1024
AUDIT_LIMIT = 200 * 1024 * 1024


def reject_links(path: Path) -> None:
    """Refuse links/junctions anywhere in the supplied path, including dangling links."""
    for component in (path, *path.parents):
        if component.is_symlink() or component.is_junction():
            raise BundleError("INVALID_BUNDLE", "Bundle paths must not contain links or junctions")


def snapshot_member(path: Path, *, capture: bool) -> tuple[str, bytes | None]:
    """Stream a fixed known member with size and identity checks; optionally retain bytes."""
    reject_links(path)
    limit = AUDIT_LIMIT if path.name == "audit.jsonl" else MEMBER_LIMIT
    if not path.is_file():
        raise BundleError("INVALID_BUNDLE", f"Missing or nonregular bundle member: {path.name}")
    digest = hashlib.sha256()
    chunks = []
    size = 0
    with path.open("rb") as stream:
        before = os.fstat(stream.fileno())
        if not stat.S_ISREG(before.st_mode) or before.st_size > limit:
            raise BundleError("INVALID_BUNDLE", f"Member size/type limit failed: {path.name}")
        while chunk := stream.read(min(1024 * 1024, limit - size + 1)):
            size += len(chunk)
            if size > limit:
                raise BundleError("INVALID_BUNDLE", f"Member size limit exceeded: {path.name}")
            digest.update(chunk)
            if capture:
                chunks.append(chunk)
        after = os.fstat(stream.fileno())
    reject_links(path)
    current = path.stat()
    signatures = {(s.st_dev, s.st_ino, s.st_size, s.st_mtime_ns) for s in (before, after, current)}
    if len(signatures) != 1 or size != before.st_size:
        raise BundleError("INVALID_BUNDLE", f"Member changed during verification: {path.name}")
    return digest.hexdigest(), b"".join(chunks) if capture else None


def strict_json(text: str | bytes) -> object:
    """Reject duplicate object keys and nonstandard nonfinite JSON constants."""

    def pairs(items):
        result = {}
        for key, value in items:
            if key in result:
                raise BundleError("INVALID_BUNDLE", "JSON contains a duplicate object key")
            result[key] = value
        return result

    def constant(value):
        raise BundleError("INVALID_BUNDLE", "JSON contains a nonfinite constant")

    try:
        decoded = text.decode("utf-8") if isinstance(text, bytes) else text
        return json.loads(
            decoded, object_pairs_hook=pairs, parse_constant=constant, parse_float=Decimal
        )
    except (ValueError, UnicodeError, RecursionError) as exc:
        if isinstance(exc, BundleError):
            raise
        raise BundleError("INVALID_BUNDLE", "Malformed UTF-8 JSON metadata") from exc


def verify_bundle_unchanged(bundle: SheetMergeBundle) -> None:
    """Recheck all seven known members before/after report creation; never follow metadata paths."""
    try:
        for name in BUNDLE_MEMBERS:
            expected = (
                bundle.manifest_sha256 if name == "manifest.json" else bundle.output_sha256[name]
            )
            digest, _ = snapshot_member(bundle.directory / name, capture=False)
            if digest != expected:
                raise BundleError("INVALID_BUNDLE", f"Bundle member changed: {name}")
    except OSError as exc:
        raise BundleError("INVALID_BUNDLE", "Bundle cannot be reverified") from exc
