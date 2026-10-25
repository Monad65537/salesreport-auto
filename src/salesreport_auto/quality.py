"""Independent occurrence conservation and exact local currency ledgers."""

from collections import Counter, defaultdict
from dataclasses import dataclass

from .errors import IntegrityError, MoneyError
from .models import ClassifiedBatch, ClassifiedRow
from .money import to_minor_units
from .product_classification import verify_product_policy

_STATUSES = ("accepted", "rejected", "redundant")


@dataclass(frozen=True)
class OutcomeCounts:
    """Input occurrence counts partitioned into terminal outcomes."""

    input: int
    accepted: int
    rejected: int
    redundant: int

    @property
    def balanced(self) -> bool:
        """Whether the input equals the three outcome counts."""
        return self.input == self.accepted + self.rejected + self.redundant


@dataclass(frozen=True)
class OutcomeAmounts:
    """Signed Python integer minor-unit totals, not presumed sales revenue."""

    input: int
    accepted: int
    rejected: int
    redundant: int

    @property
    def balanced(self) -> bool:
        """Whether all known input money is conserved across outcomes."""
        return self.input == self.accepted + self.rejected + self.redundant


@dataclass(frozen=True)
class CurrencyLedger:
    """One configured currency's occurrence, known-money, and coverage ledgers."""

    currency: str
    precision: int
    counts: OutcomeCounts
    known_amount_minor: OutcomeAmounts
    unquantified_counts: OutcomeCounts

    @property
    def monetary_complete(self) -> bool:
        """Coverage of imported occurrences, independently of acceptance."""
        return self.unquantified_counts.input == 0


@dataclass(frozen=True)
class ProvenanceCounts:
    """Disjoint local coverage categories; complete references remain unverified."""

    absent: int
    invalid: int
    complete: int
    partial: int


@dataclass(frozen=True)
class QualityReport:
    """Local quality only; this result makes no upstream completeness claim."""

    counts: OutcomeCounts
    currencies: tuple[CurrencyLedger, ...]
    unknown_currency_count: int
    unquantified_count: int
    provenance: ProvenanceCounts
    review_reasons: tuple[str, ...]

    @property
    def local_monetary_complete(self) -> bool:
        """Whether every imported occurrence has quantifiable money."""
        return self.unquantified_count == 0

    @property
    def count_balanced(self) -> bool:
        """Whether global and per-currency count partitions balance."""
        return self.counts.balanced and all(ledger.counts.balanced for ledger in self.currencies)

    @property
    def known_amounts_balanced(self) -> bool:
        """Whether exact known-money ledgers balance independently of coverage."""
        return all(ledger.known_amount_minor.balanced for ledger in self.currencies)

    @property
    def review_required(self) -> bool:
        """Whether this core import has local reasons for review."""
        return bool(self.review_reasons)

    @property
    def status(self) -> str:
        """Core analysis status, not a claim that a report file was published."""
        return "review_required" if self.review_required else "completed"


def _verify_partition(result: ClassifiedBatch) -> None:
    source = result.validated.source
    original_refs = [raw.location for raw in source.records]
    validated = {row.raw.location: row for row in result.validated.rows}
    outcomes = {item.row.raw.location: item for item in result.rows}
    if (
        len(set(original_refs)) != len(original_refs)
        or len(validated) != len(result.validated.rows)
        or len(outcomes) != len(result.rows)
        or set(original_refs) != set(validated)
        or set(original_refs) != set(outcomes)
    ):
        raise IntegrityError("Input records were lost, added, or classified more than once")
    for raw in source.records:
        row = validated[raw.location]
        item = outcomes[raw.location]
        if row.raw != raw or item.row != row:
            raise IntegrityError("Classified or validated occurrence changed")


def _verify_policy(result: ClassifiedBatch) -> None:
    config = result.validated.config
    resolved = {item.row.raw.location: item for item in result.rows}
    groups: dict[tuple[str, str], list[ClassifiedRow]] = defaultdict(list)
    for item in result.rows:
        row = item.row
        if item.status not in _STATUSES:
            raise IntegrityError("Unknown terminal outcome")
        expected_minor = None
        if row.amount is not None and row.currency in config.precision:
            try:
                expected_minor = to_minor_units(row.amount, config.precision[row.currency])
            except MoneyError:
                pass
        if row.amount_minor != expected_minor or (
            row.amount_minor is not None and type(row.amount_minor) is not int
        ):
            raise IntegrityError("Incorrect minor-unit amount on validated record")
        if row.business_key is not None:
            groups[row.business_key].append(item)
        if item.status == "accepted":
            if row.issues or item.issues or row.business_key is None or row.amount_minor is None:
                raise IntegrityError("An invalid occurrence was accepted")
            if row.transaction_date is None or row.customer_id is None:
                raise IntegrityError("Accepted record lacks a required business field")
            if item.duplicate_of is not None:
                raise IntegrityError("Accepted occurrence cannot have a duplicate reference")
        else:
            if not item.issues:
                raise IntegrityError("Excluded occurrence lacks reasons")
            if item.status == "rejected" and item.duplicate_of is not None:
                raise IntegrityError("Rejected occurrence cannot be redundant")
        if item.status == "redundant":
            winner = resolved.get(item.duplicate_of)
            payload = (
                row.product_payload if result.validated.source.mode == "products" else row.payload
            )
            if (
                winner is None
                or winner.status != "accepted"
                or row.issues
                or payload
                != (
                    winner.row.product_payload
                    if result.validated.source.mode == "products"
                    else winner.row.payload
                )
            ):
                raise IntegrityError(
                    "Redundant occurrence has no identical accepted representative"
                )
    if result.validated.source.mode == "products":
        verify_product_policy(result)
        return
    for candidates in groups.values():
        winners = [item for item in candidates if item.status == "accepted"]
        all_valid_equal = (
            not any(item.row.issues for item in candidates)
            and len({item.row.payload for item in candidates}) == 1
        )
        if all_valid_equal:
            if len(winners) != 1 or any(item.status == "rejected" for item in candidates):
                raise IntegrityError("A valid equal group does not have exactly one representative")
            earliest = min(item.row.raw.location for item in candidates)
            if winners[0].row.raw.location != earliest:
                raise IntegrityError(
                    "Duplicate representative is not the first physical occurrence"
                )
        elif any(item.status != "rejected" for item in candidates):
            raise IntegrityError("A conflicting or invalid candidate group was not quarantined")


def reconcile_transactions(result: ClassifiedBatch) -> QualityReport:
    """Verify identity/policy and independently conserve counts and known money.

    Raises IntegrityError rather than returning an apparently valid partial ledger.
    Configured currencies with no occurrences have explicit zero local ledgers.
    """
    _verify_partition(result)
    _verify_policy(result)
    config = result.validated.config
    status_counts = Counter(item.status for item in result.rows)
    counts = OutcomeCounts(
        len(result.validated.source.records), *(status_counts[s] for s in _STATUSES)
    )
    input_counts = Counter()
    input_amounts = Counter()
    input_unquantified = Counter()
    outcome_counts = Counter()
    outcome_amounts = Counter()
    outcome_unquantified = Counter()
    unknown = 0
    unquantified = 0
    provenance = Counter()
    for row in result.validated.rows:
        if row.provenance.coverage not in ("absent", "invalid", "complete", "partial"):
            raise IntegrityError("Unknown provenance coverage category")
        provenance[row.provenance.coverage] += 1
        currency = row.currency
        if row.amount_minor is None:
            unquantified += 1
        if currency not in config.precision:
            unknown += 1
        else:
            input_counts[currency] += 1
            if row.amount_minor is None:
                input_unquantified[currency] += 1
            else:
                input_amounts[currency] += row.amount_minor
    for item in result.rows:
        row = item.row
        if row.currency in config.precision:
            key = row.currency, item.status
            outcome_counts[key] += 1
            if row.amount_minor is None:
                outcome_unquantified[key] += 1
            else:
                outcome_amounts[key] += row.amount_minor
    ledgers = []
    for currency, precision in sorted(config.precision.items()):
        ledger = CurrencyLedger(
            currency,
            precision,
            OutcomeCounts(
                input_counts[currency], *(outcome_counts[currency, s] for s in _STATUSES)
            ),
            OutcomeAmounts(
                input_amounts[currency], *(outcome_amounts[currency, s] for s in _STATUSES)
            ),
            OutcomeCounts(
                input_unquantified[currency],
                *(outcome_unquantified[currency, s] for s in _STATUSES),
            ),
        )
        if not all(
            (
                ledger.counts.balanced,
                ledger.known_amount_minor.balanced,
                ledger.unquantified_counts.balanced,
            )
        ):
            raise IntegrityError("Currency count, known-money, or unquantified conservation failed")
        ledgers.append(ledger)
    if not counts.balanced or sum(input_counts.values()) + unknown != counts.input:
        raise IntegrityError("Global count conservation failed")
    if sum(input_unquantified.values()) + unknown != unquantified:
        raise IntegrityError("Global unquantified conservation failed")
    warning_codes = {warning.code for row in result.validated.rows for warning in row.warnings}
    reasons = []
    for code, applies in (
        ("NO_ACCEPTED_RECORDS", counts.accepted == 0),
        ("LOCAL_REJECTIONS", counts.rejected > 0),
        ("LOCAL_UNQUANTIFIED_AMOUNTS", unquantified > 0),
        ("INVALID_PROVENANCE", "INVALID_PROVENANCE" in warning_codes),
        ("INCOMPLETE_PROVENANCE", "INCOMPLETE_PROVENANCE" in warning_codes),
        ("PRODUCT_LABEL_CONFLICT", "PRODUCT_LABEL_CONFLICT" in warning_codes),
    ):
        if applies:
            reasons.append(code)
    return QualityReport(
        counts,
        tuple(ledgers),
        unknown,
        unquantified,
        ProvenanceCounts(*(provenance[s] for s in ("absent", "invalid", "complete", "partial"))),
        tuple(reasons),
    )
