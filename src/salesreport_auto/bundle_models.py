"""Immutable, typed SheetMerge declarations kept separate from local import quality."""

from collections import Counter
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from types import MappingProxyType
from typing import Literal

from .config import CurrencyConfig
from .models import Issue
from .quality import CurrencyLedger, OutcomeCounts

BUNDLE_OUTPUTS = (
    "master.csv",
    "master.xlsx",
    "exclusions.csv",
    "audit.jsonl",
    "reconciliation.csv",
    "reconciliation.json",
)
BUNDLE_MEMBERS = (*BUNDLE_OUTPUTS, "manifest.json")
VERIFICATION_SCOPE = (
    "Supplied manifest, artifact hashes, metadata conservation, selected-master agreement, "
    "and exclusion references checked. Hash consistency is not authenticity or proof of "
    "business completeness or source truth. Original sources and configuration fingerprints "
    "were not independently verified; audit.jsonl and reconciliation.csv were hash-checked only."
)


def freeze_json(value: object) -> object:
    """Retain additive JSON metadata as deeply immutable mappings/tuples."""
    if isinstance(value, Mapping):
        return MappingProxyType({key: freeze_json(item) for key, item in value.items()})
    if isinstance(value, list):
        return tuple(freeze_json(item) for item in value)
    return value


@dataclass(frozen=True, order=True)
class SourceRef:
    """Upstream physical identity, distinct from a local master import location."""

    source_file: str
    source_sheet: str
    source_row: int


@dataclass(frozen=True)
class SourceInput:
    """Declared original source inventory; paths are never opened."""

    source_file: str
    sha256: str
    profile: str


@dataclass(frozen=True)
class UpstreamExclusion:
    """One rejected/redundant original occurrence, preserving canonical and raw text."""

    source: SourceRef
    status: Literal["rejected", "redundant"]
    values: Mapping[str, str]
    reason_codes: tuple[str, ...]
    reasons: tuple[Issue, ...]
    duplicate_of: SourceRef | None
    raw_values: Mapping[str, object]


@dataclass(frozen=True)
class UpstreamReconciliation:
    """Validated declarations; excluded known-money ledgers are not source-audited."""

    counts: OutcomeCounts
    currencies: tuple[CurrencyLedger, ...]
    unknown_currency_count: int
    unquantified_count: int

    @property
    def monetary_complete(self) -> bool:
        """Known coverage of original occurrences, independent of acceptance."""
        return self.unquantified_count == 0


@dataclass(frozen=True)
class SheetMergeBundle:
    """Hash-checked artifact snapshots; the engine also verifies master agreement."""

    directory: Path
    tool_version: str
    manifest_version: int
    configuration: Mapping[str, object]
    configuration_sha256: str
    config: CurrencyConfig
    policies: Mapping[str, object]
    inputs: tuple[SourceInput, ...]
    reconciliation: UpstreamReconciliation
    exclusions: tuple[UpstreamExclusion, ...]
    output_sha256: Mapping[str, str]
    manifest_sha256: str

    @property
    def acceptance_complete(self) -> bool:
        """Derived rejection-free scope, never a claim of real-business completeness."""
        return self.reconciliation.counts.rejected == 0

    @property
    def review_reasons(self) -> tuple[str, ...]:
        """Propagate upstream rejection and unquantified-amount review independently."""
        return tuple(
            code
            for code, applies in (
                ("UPSTREAM_REJECTIONS", self.reconciliation.counts.rejected > 0),
                ("UPSTREAM_UNQUANTIFIED_AMOUNTS", not self.reconciliation.monetary_complete),
            )
            if applies
        )

    @property
    def reason_counts(self) -> Mapping[str, int]:
        """Count reason occurrences, which may overlap within excluded records."""
        return MappingProxyType(
            dict(
                sorted(
                    Counter(code for row in self.exclusions for code in row.reason_codes).items()
                )
            )
        )
