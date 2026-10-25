"""Exact accepted-only product metrics with currency-scoped revenue rankings."""

from collections import defaultdict
from dataclasses import dataclass, replace

from .errors import ConfigurationError, IntegrityError
from .models import ClassifiedBatch
from .money import MonetaryAverage, format_minor
from .quality import reconcile_transactions


@dataclass(frozen=True)
class ProductSales:
    """One currency/product's exact quantities, prices, revenue, and distinct keys."""

    currency: str
    product_id: str
    product_name: str
    precision: int
    accepted_line_count: int
    transaction_count: int
    unique_customer_count: int
    positive_quantity: int
    negative_quantity_magnitude: int
    net_quantity: int
    net_revenue_minor: int
    minimum_unit_price_minor: int
    maximum_unit_price_minor: int
    positive_quantity_average: MonetaryAverage
    revenue_rank: int

    @property
    def net_revenue(self) -> str:
        """Return signed exact recorded product revenue at currency precision."""
        return format_minor(self.net_revenue_minor, self.precision)

    @property
    def minimum_unit_price(self) -> str:
        """Include zero- and negative-quantity lines in the observed minimum."""
        return format_minor(self.minimum_unit_price_minor, self.precision)

    @property
    def maximum_unit_price(self) -> str:
        """Include zero- and negative-quantity lines in the observed maximum."""
        return format_minor(self.maximum_unit_price_minor, self.precision)

    @property
    def positive_quantity_weighted_unit_price(self) -> str:
        """Final HALF_UP positive-quantity price, or N/A for a zero denominator."""
        return self.positive_quantity_average.display


def product_sales(result: ClassifiedBatch) -> tuple[ProductSales, ...]:
    """Aggregate products across branches, with no cross-currency consolidation."""
    if result.validated.source.mode != "products":
        raise ConfigurationError("product_sales requires product-mode classifications")
    quality = reconcile_transactions(result)
    groups = defaultdict(list)
    for item in result.accepted:
        groups[item.row.currency, item.row.product_id].append(item.row)
    grouped = []
    for (currency, product_id), rows in groups.items():
        precision = result.validated.config.precision[currency]
        positive = sum(row.quantity for row in rows if row.quantity > 0)
        negative = sum(-row.quantity for row in rows if row.quantity < 0)
        positive_amount = sum(row.amount_minor for row in rows if row.quantity > 0)
        labels = sorted({row.product_name for row in rows if row.product_name is not None})
        grouped.append(
            ProductSales(
                currency,
                product_id,
                labels[0] if labels else product_id,
                precision,
                len(rows),
                len({row.business_key for row in rows}),
                len({(row.branch_id, row.customer_id) for row in rows}),
                positive,
                negative,
                positive - negative,
                sum(row.amount_minor for row in rows),
                min(row.unit_price_minor for row in rows),
                max(row.unit_price_minor for row in rows),
                MonetaryAverage(positive_amount, positive, precision),
                0,
            )
        )
    ordered = []
    previous_currency = None
    rank = 0
    for product in sorted(
        grouped, key=lambda row: (row.currency, -row.net_revenue_minor, row.product_id)
    ):
        rank = rank + 1 if previous_currency == product.currency else 1
        ordered.append(replace(product, revenue_rank=rank))
        previous_currency = product.currency
    for ledger in quality.currencies:
        members = [row for row in ordered if row.currency == ledger.currency]
        if (
            sum(row.accepted_line_count for row in members) != ledger.counts.accepted
            or sum(row.net_revenue_minor for row in members) != ledger.known_amount_minor.accepted
        ):
            raise IntegrityError("Products do not conserve accepted lines and revenue")
    return tuple(ordered)
