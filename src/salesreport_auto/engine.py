"""Read-only sales analysis with optional verified upstream bundle declarations."""

from dataclasses import dataclass
from pathlib import Path

from .aggregation import TransactionAggregates, aggregate_transactions
from .bundle import load_bundle, verify_master_agreement
from .bundle_io import verify_bundle_unchanged
from .bundle_models import SheetMergeBundle
from .classification import classify_transactions
from .config import CurrencyConfig, InputLimits
from .errors import BundleError, InputError
from .models import ClassifiedBatch, ReportMode
from .product_classification import classify_products
from .product_validation import validate_products
from .quality import QualityReport, reconcile_transactions
from .readers import read_products, read_transactions
from .validation import validate_transactions


@dataclass(frozen=True)
class TransactionAnalysis:
    """Typed local results with optional internally verified upstream declarations."""

    classified: ClassifiedBatch
    quality: QualityReport
    aggregates: TransactionAggregates
    upstream: SheetMergeBundle | None = None

    @property
    def mode(self) -> ReportMode:
        """Return the explicit input mode; counts refer to that mode's record grain."""
        return self.classified.validated.source.mode


@dataclass(frozen=True)
class ProductAnalysis(TransactionAnalysis):
    """Product lines using the shared quality and transaction-metric result shape."""


def analyze_transactions(
    path: str | Path,
    *,
    config: CurrencyConfig | None = None,
    limits: InputLimits | None = None,
    bundle: str | Path | None = None,
) -> TransactionAnalysis:
    """Analyze one transaction file, optionally requiring SheetMerge bundle agreement.

    Row failures are retained as local rejected occurrences. Structural or
    invariant failures raise InputError/IntegrityError. No files are written.
    """
    upstream = load_bundle(bundle, config=config) if bundle is not None else None
    try:
        source = read_transactions(path, limits=limits)
    except InputError as exc:
        if upstream is not None:
            raise BundleError(
                "BUNDLE_MISMATCH", "Selected master cannot be read under the transaction schema"
            ) from exc
        raise
    resolved = upstream.config if upstream is not None else config
    classified = classify_transactions(validate_transactions(source, resolved))
    quality = reconcile_transactions(classified)
    if upstream is not None:
        verify_master_agreement(upstream, classified, quality)
    aggregates = aggregate_transactions(classified)
    if upstream is not None:
        verify_bundle_unchanged(upstream)
    return TransactionAnalysis(classified, quality, aggregates, upstream)


def analyze_products(
    path: str | Path, *, config: CurrencyConfig | None = None, limits: InputLimits | None = None
) -> ProductAnalysis:
    """Read product lines, quarantine whole invalid transactions, and aggregate exactly."""
    source = read_products(path, limits=limits)
    classified = classify_products(validate_products(source, config))
    quality = reconcile_transactions(classified)
    return ProductAnalysis(classified, quality, aggregate_transactions(classified))
