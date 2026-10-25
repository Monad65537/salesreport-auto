"""Line-key duplicate resolution and complete identified-transaction quarantine."""

from collections import defaultdict
from dataclasses import replace

from .errors import ConfigurationError, IntegrityError, MoneyError
from .models import ClassifiedBatch, ClassifiedRow, Issue, ValidatedBatch, ValidatedRow
from .money import to_minor_units
from .validation import sort_issues


def _inconsistent(rows: list[ValidatedRow]) -> bool:
    return any(
        len({value for row in rows if (value := getattr(row, field)) is not None}) > 1
        for field in ("transaction_date", "customer_id", "currency")
    )


def _label_conflicts(items: tuple[ClassifiedRow, ...]) -> set[tuple[str, str]]:
    labels: dict[tuple[str, str], set[str]] = defaultdict(set)
    for item in items:
        row = item.row
        if item.status == "accepted" and row.product_name is not None:
            labels[row.currency, row.product_id].add(row.product_name)
    return {key for key, values in labels.items() if len(values) > 1}


def classify_products(batch: ValidatedBatch) -> ClassifiedBatch:
    """Accept a valid complete transaction's unique lines, or reject all occurrences.

    Physical representatives are selected only after every readable member passes.
    Cross-line accepted label disagreements warn; they do not alter financial outcomes.
    """
    if batch.source.mode != "products":
        raise ConfigurationError("classify_products requires product-mode input")
    if any(
        warning.code == "PRODUCT_LABEL_CONFLICT" for row in batch.rows for warning in row.warnings
    ):
        # Recompute derived warnings when an existing validated result is reused.
        batch = replace(
            batch,
            rows=tuple(
                replace(
                    row,
                    warnings=tuple(
                        warning
                        for warning in row.warnings
                        if warning.code != "PRODUCT_LABEL_CONFLICT"
                    ),
                )
                for row in batch.rows
            ),
        )
    groups = defaultdict(list)
    seen = set()
    outcomes = {}
    for row in batch.rows:
        ref = row.raw.location
        if ref in seen:
            raise IntegrityError("Validated input contains duplicate physical locations")
        seen.add(ref)
        if row.business_key is None:
            outcomes[ref] = ClassifiedRow(row, "rejected", row.issues)
        else:
            groups[row.business_key].append(row)
    for rows in groups.values():
        lines = defaultdict(list)
        for row in rows:
            if row.line_key is not None:
                lines[row.line_key].append(row)
        group_issues = []
        if any(row.issues for row in rows):
            group_issues.append(
                Issue(
                    "TRANSACTION_HAS_INVALID_LINE",
                    "transaction",
                    "Identified transaction contains an invalid line",
                )
            )
        if any(
            len({row.product_payload for row in candidates if not row.issues}) > 1
            for candidates in lines.values()
        ):
            group_issues.append(
                Issue(
                    "CONFLICTING_LINE",
                    "transaction",
                    "Same line key has different normalized product values",
                )
            )
        if _inconsistent(rows):
            group_issues.append(
                Issue(
                    "INCONSISTENT_TRANSACTION",
                    "transaction",
                    "Transaction date, customer, or currency differs across lines",
                )
            )
        if group_issues:
            for row in rows:
                outcomes[row.raw.location] = ClassifiedRow(
                    row, "rejected", sort_issues([*row.issues, *group_issues], mode="products")
                )
            continue
        for candidates in lines.values():
            candidates.sort(key=lambda row: row.raw.location)
            winner = candidates[0]
            outcomes[winner.raw.location] = ClassifiedRow(winner, "accepted")
            for row in candidates[1:]:
                outcomes[row.raw.location] = ClassifiedRow(
                    row,
                    "redundant",
                    (Issue("EXACT_DUPLICATE", "transaction", "Identical normalized product line"),),
                    winner.raw.location,
                )
    ordered = tuple(outcomes[ref] for ref in sorted(outcomes))
    conflicts = _label_conflicts(ordered)
    warned_refs = {
        item.row.raw.location
        for item in ordered
        if item.status == "accepted" and (item.row.currency, item.row.product_id) in conflicts
    }
    if warned_refs:
        updated = {
            row.raw.location: replace(
                row,
                warnings=sort_issues(
                    [
                        *row.warnings,
                        Issue(
                            "PRODUCT_LABEL_CONFLICT",
                            "product_name",
                            "Accepted product lines have differing nonempty labels",
                        ),
                    ],
                    mode="products",
                ),
            )
            if row.raw.location in warned_refs
            else row
            for row in batch.rows
        }
        batch = replace(batch, rows=tuple(updated[row.raw.location] for row in batch.rows))
        ordered = tuple(replace(item, row=updated[item.row.raw.location]) for item in ordered)
    return ClassifiedBatch(batch, ordered)


def verify_product_policy(result: ClassifiedBatch) -> None:
    """Independently verify accepted line arithmetic, keys, and transaction quarantine."""
    groups = defaultdict(list)
    config = result.validated.config
    for item in result.rows:
        row = item.row
        if row.business_key is not None:
            groups[row.business_key].append(item)
        if item.status in ("accepted", "redundant"):
            try:
                expected_price = (
                    to_minor_units(row.unit_price, config.precision[row.currency])
                    if row.unit_price is not None
                    else None
                )
            except MoneyError as exc:
                raise IntegrityError(
                    "Accepted product price cannot be represented exactly"
                ) from exc
            if (
                row.line_key is None
                or row.product_id is None
                or type(row.quantity) is not int
                or abs(row.quantity) > 1_000_000_000
                or row.unit_price is None
                or row.unit_price < 0
                or type(row.unit_price_minor) is not int
                or expected_price != row.unit_price_minor
                or row.amount_minor != row.quantity * row.unit_price_minor
            ):
                raise IntegrityError(
                    "Accepted/redundant product line has invalid fields or arithmetic"
                )
    for candidates in groups.values():
        lines = defaultdict(list)
        for item in candidates:
            lines[item.row.line_key].append(item)
        valid = (
            not any(item.row.issues for item in candidates)
            and not _inconsistent([item.row for item in candidates])
            and all(
                len({item.row.product_payload for item in items}) == 1 for items in lines.values()
            )
        )
        if not valid:
            if any(item.status != "rejected" for item in candidates):
                raise IntegrityError("Product transaction was only partially quarantined")
        else:
            for items in lines.values():
                winners = [item for item in items if item.status == "accepted"]
                if len(winners) != 1 or any(item.status == "rejected" for item in items):
                    raise IntegrityError(
                        "Valid line group lacks exactly one accepted representative"
                    )
                if winners[0].row.raw.location != min(item.row.raw.location for item in items):
                    raise IntegrityError("Product line representative is not the first occurrence")
    conflicts = _label_conflicts(result.rows)
    for item in result.rows:
        row = item.row
        expected = item.status == "accepted" and (row.currency, row.product_id) in conflicts
        actual = any(warning.code == "PRODUCT_LABEL_CONFLICT" for warning in row.warnings)
        if expected != actual:
            raise IntegrityError("Product label conflict warnings disagree with accepted lines")
