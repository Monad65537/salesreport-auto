"""Pure accepted-only transaction metrics, partitioned by currency."""

from collections import defaultdict
from collections.abc import Callable
from dataclasses import dataclass

from .models import ClassifiedBatch, ValidatedRow
from .money import MonetaryAverage, format_minor
from .product_aggregation import ProductSales, product_sales
from .quality import reconcile_transactions


@dataclass(frozen=True)
class SalesMetrics:
    """Exact counts and integer revenue, with a retained average ratio."""

    precision: int
    accepted_record_count: int
    transaction_count: int
    unique_customer_count: int
    net_revenue_minor: int
    average: MonetaryAverage

    @property
    def net_revenue(self) -> str:
        """Return exact fixed-precision signed revenue text."""
        return format_minor(self.net_revenue_minor, self.precision)

    @property
    def average_transaction_value(self) -> str:
        """Return the final HALF_UP average display value."""
        return self.average.display


@dataclass(frozen=True)
class CurrencySummary:
    """One observed accepted currency's transaction metrics."""

    currency: str
    metrics: SalesMetrics


@dataclass(frozen=True)
class MonthlySales:
    """One observed calendar month in a single currency."""

    currency: str
    month: str
    metrics: SalesMetrics


@dataclass(frozen=True)
class BranchSales:
    """One branch in a single currency, with branch-scoped customer identities."""

    currency: str
    branch_id: str
    metrics: SalesMetrics


@dataclass(frozen=True)
class TransactionAggregates:
    """Ordered common rollups, with product metrics only for explicit product inputs."""

    summary: tuple[CurrencySummary, ...]
    monthly: tuple[MonthlySales, ...]
    branches: tuple[BranchSales, ...]
    products: tuple[ProductSales, ...] = ()


def _group(
    result: ClassifiedBatch, key: Callable[[ValidatedRow], tuple[str, ...]]
) -> dict[tuple[str, ...], SalesMetrics]:
    groups = defaultdict(list)
    for item in result.accepted:
        groups[key(item.row)].append(item.row)
    metrics = {}
    for group_key, rows in sorted(groups.items()):
        precision = result.validated.config.precision[group_key[0]]
        transaction_totals = defaultdict(int)
        for row in rows:
            transaction_totals[row.business_key] += row.amount_minor
        revenue = sum(transaction_totals.values())
        transactions = len(transaction_totals)
        customers = len({(row.branch_id, row.customer_id) for row in rows})
        metrics[group_key] = SalesMetrics(
            precision,
            len(rows),
            transactions,
            customers,
            revenue,
            MonetaryAverage(revenue, transactions, precision),
        )
    return metrics


def _summary(result: ClassifiedBatch) -> tuple[CurrencySummary, ...]:
    return tuple(
        CurrencySummary(key[0], value)
        for key, value in _group(result, lambda r: (r.currency,)).items()
    )


def _monthly(result: ClassifiedBatch) -> tuple[MonthlySales, ...]:
    groups = _group(result, lambda row: (row.currency, row.transaction_date.isoformat()[:7]))
    return tuple(MonthlySales(*key, value) for key, value in groups.items())


def _branches(result: ClassifiedBatch) -> tuple[BranchSales, ...]:
    groups = _group(result, lambda row: (row.currency, row.branch_id))
    return tuple(BranchSales(*key, value) for key, value in groups.items())


def summarize_transactions(result: ClassifiedBatch) -> tuple[CurrencySummary, ...]:
    """Compute currency summaries from verified accepted representatives only."""
    reconcile_transactions(result)
    return _summary(result)


def monthly_sales(result: ClassifiedBatch) -> tuple[MonthlySales, ...]:
    """Compute observed currency/month groups without filling missing months."""
    reconcile_transactions(result)
    return _monthly(result)


def branch_sales(result: ClassifiedBatch) -> tuple[BranchSales, ...]:
    """Compute currency/branch groups in exact Unicode lexical order."""
    reconcile_transactions(result)
    return _branches(result)


def aggregate_transactions(result: ClassifiedBatch) -> TransactionAggregates:
    """Verify the local outcomes once and compute all three pure groupings."""
    reconcile_transactions(result)
    products = product_sales(result) if result.validated.source.mode == "products" else ()
    return TransactionAggregates(_summary(result), _monthly(result), _branches(result), products)
