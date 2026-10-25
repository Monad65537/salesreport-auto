"""Optional SheetMerge adapter: verify artifacts and reconcile the selected master."""

from collections import Counter
from pathlib import Path
from types import MappingProxyType

from .bundle_io import reject_links, snapshot_member, strict_json, verify_bundle_unchanged
from .bundle_models import BUNDLE_MEMBERS, SheetMergeBundle, SourceRef, freeze_json
from .bundle_schema import parse_exclusions, parse_manifest, parse_reconciliation
from .config import CurrencyConfig
from .errors import BundleError
from .models import ALL_COLUMNS, ClassifiedBatch
from .quality import QualityReport


def load_bundle(directory: str | Path, *, config: CurrencyConfig | None = None) -> SheetMergeBundle:
    """Hash-check all seven members and parse version-1 declarations read-only.

    The selected master's own local validation/agreement is a separate mandatory
    step, performed by analyze_transactions(bundle=...) before attaching upstream.
    """
    root = Path(directory).absolute()
    try:
        reject_links(root)
        if not root.is_dir():
            raise BundleError("INVALID_BUNDLE", "Bundle must be an existing directory")
        hashes = {}
        captured = {}
        for name in BUNDLE_MEMBERS:
            digest, data = snapshot_member(
                root / name,
                capture=name
                in (
                    "manifest.json",
                    "reconciliation.json",
                    "exclusions.csv",
                ),
            )
            hashes[name] = digest
            if data is not None:
                captured[name] = data
        manifest = parse_manifest(captured["manifest.json"])
        for name, declared in manifest.output_sha256.items():
            if hashes[name] != declared:
                raise BundleError("INVALID_BUNDLE", f"SHA-256 mismatch: {name}")
        if config is not None:
            for code, origin in config.origins.items():
                if origin == "user" and (
                    code not in manifest.config.precision
                    or config.precision[code] != manifest.config.precision[code]
                ):
                    raise BundleError(
                        "INVALID_BUNDLE",
                        f"Explicit currency precision disagrees with bundle: {code}",
                    )
        reconciliation_data = strict_json(captured["reconciliation.json"])
        reconciliation = parse_reconciliation(reconciliation_data, manifest.config)
        parse_reconciliation(manifest.reconciliation, manifest.config)
        if reconciliation_data != manifest.reconciliation:
            raise BundleError(
                "INVALID_BUNDLE", "Manifest reconciliation disagrees with reconciliation.json"
            )
        result = SheetMergeBundle(
            root,
            manifest.tool_version,
            1,
            freeze_json(manifest.configuration),
            manifest.configuration_sha256,
            manifest.config,
            freeze_json(manifest.policies),
            manifest.inputs,
            reconciliation,
            parse_exclusions(captured["exclusions.csv"]),
            MappingProxyType(dict(manifest.output_sha256)),
            hashes["manifest.json"],
        )
        verify_bundle_unchanged(result)
        return result
    except (OSError, ValueError, RecursionError) as exc:
        if isinstance(exc, BundleError):
            raise
        raise BundleError("INVALID_BUNDLE", "Cannot read supported SheetMerge bundle") from exc


def verify_master_agreement(
    bundle: SheetMergeBundle, classified: ClassifiedBatch, quality: QualityReport
) -> None:
    """Compare exactly one master against accepted ledgers, inventory, and exclusion refs."""
    source = classified.validated.source
    name = source.path.name

    def mismatch(message: str) -> None:
        raise BundleError("BUNDLE_MISMATCH", message)

    if source.mode != "transactions" or name not in ("master.csv", "master.xlsx"):
        mismatch("Select exactly one input named master.csv or master.xlsx in transaction mode")
    if source.sha256 != bundle.output_sha256[name]:
        mismatch("Selected master bytes do not match the supplied bundle")
    if len(source.headers) != len(ALL_COLUMNS) or set(source.headers) != set(ALL_COLUMNS):
        mismatch("Selected master must contain exactly the nine SheetMerge columns")
    if dict(classified.validated.config.precision) != dict(bundle.config.precision):
        mismatch("Selected master did not use bundle currency precision")
    report = bundle.reconciliation
    if (
        quality.counts.rejected
        or quality.counts.redundant
        or quality.unquantified_count
        or quality.counts.accepted != report.counts.accepted
        or any(
            item.row.provenance.coverage != "complete" or item.row.warnings
            for item in classified.rows
        )
    ):
        mismatch(
            "Selected master contains invalid, redundant, missing, or untraceable accepted records"
        )
    local = {ledger.currency: ledger for ledger in quality.currencies}
    for ledger in report.currencies:
        actual = local.get(ledger.currency)
        if actual is None or (
            actual.counts.accepted != ledger.counts.accepted
            or actual.known_amount_minor.accepted != ledger.known_amount_minor.accepted
        ):
            mismatch("Selected master accepted currency counts or exact totals disagree")
    inventory = {item.source_file for item in bundle.inputs}
    accepted = set()
    for item in classified.rows:
        ref = item.row.provenance
        key = SourceRef(ref.source_file, ref.source_sheet or "", ref.source_row)
        if key.source_file not in inventory or key in accepted:
            mismatch("Accepted provenance is duplicated or absent from declared source inventory")
        accepted.add(key)
    seen = set(accepted)
    excluded = Counter()
    currency_excluded = Counter()
    for row in bundle.exclusions:
        if row.source.source_file not in inventory or row.source in seen:
            raise BundleError(
                "INVALID_BUNDLE",
                "Excluded provenance is duplicated or absent from source inventory",
            )
        seen.add(row.source)
        excluded[row.status] += 1
        currency_excluded[row.values["currency"], row.status] += 1
        if row.duplicate_of is not None and row.duplicate_of not in accepted:
            raise BundleError(
                "INVALID_BUNDLE", "Redundant exclusion does not reference an accepted master source"
            )
    if (
        excluded["rejected"] != report.counts.rejected
        or excluded["redundant"] != report.counts.redundant
        or len(seen) != report.counts.input
    ):
        raise BundleError(
            "INVALID_BUNDLE", "Exclusion/provenance counts disagree with upstream reconciliation"
        )
    for ledger in report.currencies:
        for status in ("rejected", "redundant"):
            if currency_excluded[ledger.currency, status] != getattr(ledger.counts, status):
                raise BundleError(
                    "INVALID_BUNDLE",
                    "Exclusion currency counts disagree with upstream reconciliation",
                )
