"""Immutable typed raw records, normalized transactions, and local outcomes."""

from collections.abc import Mapping
from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path
from types import MappingProxyType
from typing import Literal

from .config import CurrencyConfig

BUSINESS_COLUMNS = (
    "branch_id",
    "transaction_id",
    "transaction_date",
    "customer_id",
    "amount",
    "currency",
)
PROVENANCE_COLUMNS = ("source_file", "source_sheet", "source_row")
ALL_COLUMNS = (*BUSINESS_COLUMNS, *PROVENANCE_COLUMNS)
PRODUCT_COLUMNS = (*BUSINESS_COLUMNS, "line_id", "product_id", "quantity", "unit_price")
PRODUCT_ALL_COLUMNS = (*PRODUCT_COLUMNS, "product_name", *PROVENANCE_COLUMNS)
ReportMode = Literal["transactions", "products"]
Status = Literal["accepted", "rejected", "redundant"]
CellKind = Literal["text", "numeric", "date", "boolean", "formula", "error", "other"]
RawValue = str | int | float | bool | date | datetime | time | None
ProvenanceCoverage = Literal["absent", "invalid", "complete", "partial"]


@dataclass(frozen=True, order=True)
class ImportRef:
    """Physical location in the imported file, separate from source provenance."""

    import_file: str
    import_sheet: str
    import_row: int


@dataclass(frozen=True)
class RawCell:
    """Decoded value and storage kind; text is preserved without normalization."""

    value: RawValue
    kind: CellKind = "text"


@dataclass(frozen=True)
class RawRecord:
    """One physical input occurrence, retaining header order and raw values."""

    location: ImportRef
    cells: Mapping[str, RawCell]

    def __post_init__(self) -> None:
        object.__setattr__(self, "cells", MappingProxyType(dict(self.cells)))

    @property
    def raw_values(self) -> Mapping[str, RawValue]:
        """Expose original decoded values without NA/type inference."""
        return MappingProxyType({name: cell.value for name, cell in self.cells.items()})


@dataclass(frozen=True)
class TransactionInput:
    """A single read-only file snapshot and its explicit record schema."""

    path: Path
    sha256: str
    headers: tuple[str, ...]
    selected_sheet: str
    ignored_sheets: tuple[str, ...]
    records: tuple[RawRecord, ...]
    mode: ReportMode = "transactions"


@dataclass(frozen=True)
class Issue:
    """Machine-readable rejection or warning with a non-sensitive explanation."""

    code: str
    field: str
    message: str


@dataclass(frozen=True)
class Provenance:
    """Optional unverified upstream reference; absent values are never invented."""

    source_file: str | None
    source_sheet: str | None
    source_row: int | None
    coverage: ProvenanceCoverage


@dataclass(frozen=True)
class ValidatedRow:
    """Independent field validation, before resolving transaction candidates."""

    raw: RawRecord
    branch_id: str | None
    transaction_id: str | None
    transaction_date: date | None
    customer_id: str | None
    amount: Decimal | None
    currency: str | None
    amount_minor: int | None
    provenance: Provenance
    issues: tuple[Issue, ...]
    warnings: tuple[Issue, ...]
    line_id: str | None = None
    product_id: str | None = None
    quantity: int | None = None
    unit_price: Decimal | None = None
    unit_price_minor: int | None = None
    product_name: str | None = None

    @property
    def business_key(self) -> tuple[str, str] | None:
        """Return the readable branch-scoped key, if both text IDs are present."""
        if self.branch_id is not None and self.transaction_id is not None:
            return self.branch_id, self.transaction_id
        return None

    @property
    def payload(self) -> tuple[object, ...]:
        """Normalized six-field equality; provenance does not affect duplicates."""
        return (
            self.branch_id,
            self.transaction_id,
            self.transaction_date,
            self.customer_id,
            self.amount,
            self.currency,
        )

    @property
    def line_key(self) -> tuple[str, str, str] | None:
        """Return the complete branch/transaction/line identifier when readable."""
        if self.business_key is not None and self.line_id is not None:
            return (*self.business_key, self.line_id)
        return None

    @property
    def product_payload(self) -> tuple[object, ...]:
        """Normalized line equality, including optional label but excluding provenance."""
        return (
            *self.payload,
            self.line_id,
            self.product_id,
            self.quantity,
            self.unit_price,
            self.product_name,
        )


@dataclass(frozen=True)
class ValidatedBatch:
    """The snapshot, resolved configuration, and all independently checked rows."""

    source: TransactionInput
    config: CurrencyConfig
    rows: tuple[ValidatedRow, ...]


@dataclass(frozen=True)
class ClassifiedRow:
    """One terminal outcome; issues may include transaction-wide quarantine."""

    row: ValidatedRow
    status: Status
    issues: tuple[Issue, ...] = ()
    duplicate_of: ImportRef | None = None


@dataclass(frozen=True)
class ClassifiedBatch:
    """All local occurrences, in physical input order, after key resolution."""

    validated: ValidatedBatch
    rows: tuple[ClassifiedRow, ...]

    @property
    def accepted(self) -> tuple[ClassifiedRow, ...]:
        """Return accepted representatives, not merely field-valid candidates."""
        return tuple(item for item in self.rows if item.status == "accepted")

    @property
    def rejected(self) -> tuple[ClassifiedRow, ...]:
        """Return all quarantined or invalid local occurrences."""
        return tuple(item for item in self.rows if item.status == "rejected")

    @property
    def redundant(self) -> tuple[ClassifiedRow, ...]:
        """Return equal copies whose representative is accepted."""
        return tuple(item for item in self.rows if item.status == "redundant")
