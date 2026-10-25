"""Deterministic physical representatives and full-key conflict quarantine."""

from collections import defaultdict

from .errors import ConfigurationError, IntegrityError
from .models import ClassifiedBatch, ClassifiedRow, Issue, ValidatedBatch, ValidatedRow
from .validation import sort_issues


def classify_transactions(batch: ValidatedBatch) -> ClassifiedBatch:
    """Resolve readable branch/transaction keys without losing any occurrence."""
    if batch.source.mode != "transactions":
        raise ConfigurationError("classify_transactions requires transaction-mode input")
    groups: dict[tuple[str, str], list[ValidatedRow]] = defaultdict(list)
    outcomes = {}
    seen = set()
    for row in batch.rows:
        ref = row.raw.location
        if ref in seen:
            raise IntegrityError("Validated input contains duplicate physical locations")
        seen.add(ref)
        if row.business_key is None:
            outcomes[ref] = ClassifiedRow(row, "rejected", row.issues)
        else:
            groups[row.business_key].append(row)
    for candidates in groups.values():
        candidates.sort(key=lambda row: row.raw.location)
        invalid = any(row.issues for row in candidates)
        conflicting = len({row.payload for row in candidates if not row.issues}) > 1
        group_issues = []
        if conflicting:
            group_issues.append(
                Issue(
                    "CONFLICTING_TRANSACTION",
                    "transaction",
                    "Same-key normalized transactions disagree",
                )
            )
        if invalid:
            group_issues.append(
                Issue(
                    "KEY_HAS_INVALID_CANDIDATE",
                    "transaction",
                    "Readable key contains an invalid candidate",
                )
            )
        if invalid or conflicting:
            for row in candidates:
                outcomes[row.raw.location] = ClassifiedRow(
                    row, "rejected", sort_issues([*row.issues, *group_issues])
                )
        else:
            winner = candidates[0]
            outcomes[winner.raw.location] = ClassifiedRow(winner, "accepted")
            for row in candidates[1:]:
                outcomes[row.raw.location] = ClassifiedRow(
                    row,
                    "redundant",
                    (Issue("EXACT_DUPLICATE", "transaction", "Identical normalized transaction"),),
                    winner.raw.location,
                )
    return ClassifiedBatch(batch, tuple(outcomes[ref] for ref in sorted(outcomes)))
