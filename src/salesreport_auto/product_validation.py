"""Product-line validation reusing the established six common business fields."""

import re
from dataclasses import replace

from .config import CurrencyConfig
from .errors import ConfigurationError, MoneyError
from .models import Issue, RawRecord, TransactionInput, ValidatedBatch, ValidatedRow
from .money import parse_decimal, to_minor_units
from .validation import _absent, sort_issues, validate_row

MAX_QUANTITY = 1_000_000_000


def _quantity(text: str) -> int:
    if not re.fullmatch(r"[+-]?[0-9]+", text):
        raise ValueError("Quantity must be signed whole-unit text")
    digits = text.lstrip("+-").lstrip("0") or "0"
    if len(digits) > 10 or int(digits) > MAX_QUANTITY:
        raise ValueError("Absolute quantity must not exceed 1,000,000,000")
    return (-1 if text.startswith("-") else 1) * int(digits)


def validate_product_row(raw: RawRecord, config: CurrencyConfig) -> ValidatedRow:
    """Check required product fields, exact price/quantity agreement, and labels.

    A valid recorded amount remains quantified even if a quantity or price fails.
    Optional label/provenance values never replace product identity.
    """
    common = validate_row(raw, config)
    if any(issue.code == "BLANK_ROW" for issue in common.issues):
        return common
    issues = list(common.issues)
    available = {}
    for field in ("line_id", "product_id", "quantity", "unit_price", "product_name"):
        cell = raw.cells.get(field)
        if cell is None or _absent(cell):
            if field != "product_name":
                issues.append(Issue("MISSING_REQUIRED_VALUE", field, "Required value is empty"))
        elif cell.kind in ("formula", "error"):
            code = "FORMULA_NOT_ALLOWED" if cell.kind == "formula" else "EXCEL_CELL_ERROR"
            issues.append(Issue(code, field, "Mapped formula/error cells cannot be trusted"))
        elif cell.kind != "text" or not isinstance(cell.value, str):
            code = "ID_NOT_TEXT" if field in ("line_id", "product_id") else "INVALID_CELL_TYPE"
            issues.append(Issue(code, field, "Product business values must be stored as text"))
        else:
            available[field] = cell.value
    quantity = None
    if "quantity" in available:
        try:
            quantity = _quantity(available["quantity"])
        except ValueError as exc:
            issues.append(Issue("INVALID_QUANTITY", "quantity", str(exc)))
    unit_price = None
    unit_price_minor = None
    if "unit_price" in available:
        try:
            unit_price = parse_decimal(available["unit_price"])
            if unit_price < 0:
                issues.append(
                    Issue("NEGATIVE_UNIT_PRICE", "unit_price", "Unit price must be nonnegative")
                )
            elif common.currency is not None:
                unit_price_minor = to_minor_units(unit_price, config.precision[common.currency])
        except MoneyError as exc:
            issues.append(Issue(exc.code, "unit_price", str(exc)))
    if quantity is not None and unit_price_minor is not None and common.amount_minor is not None:
        if common.amount_minor != quantity * unit_price_minor:
            issues.append(
                Issue(
                    "LINE_AMOUNT_MISMATCH",
                    "amount",
                    "Amount must exactly equal quantity times unit price",
                )
            )
    label = available.get("product_name")
    if label is not None and len(label) > 200:
        issues.append(
            Issue("INVALID_PRODUCT_NAME", "product_name", "Product label exceeds 200 characters")
        )
        label = None
    return replace(
        common,
        line_id=available.get("line_id"),
        product_id=available.get("product_id"),
        quantity=quantity,
        unit_price=unit_price,
        unit_price_minor=unit_price_minor,
        product_name=label,
        issues=sort_issues(issues, mode="products"),
    )


def validate_products(
    source: TransactionInput, config: CurrencyConfig | None = None
) -> ValidatedBatch:
    """Validate every product occurrence before any transaction/duplicate decisions."""
    if source.mode != "products":
        raise ConfigurationError("validate_products requires product-mode input")
    resolved = config or CurrencyConfig()
    return ValidatedBatch(
        source, resolved, tuple(validate_product_row(row, resolved) for row in source.records)
    )
