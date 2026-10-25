"""Independent field validation; provenance problems warn without rejecting sales."""

import re
from datetime import date

from .config import CurrencyConfig
from .errors import ConfigurationError, MoneyError
from .models import (
    ALL_COLUMNS,
    BUSINESS_COLUMNS,
    PRODUCT_ALL_COLUMNS,
    PROVENANCE_COLUMNS,
    Issue,
    Provenance,
    RawCell,
    RawRecord,
    ReportMode,
    TransactionInput,
    ValidatedBatch,
    ValidatedRow,
)
from .money import parse_decimal, to_minor_units


def sort_issues(issues: list[Issue], *, mode: ReportMode = "transactions") -> tuple[Issue, ...]:
    """Order issues by schema field then code, with row/group issues last."""
    columns = PRODUCT_ALL_COLUMNS if mode == "products" else ALL_COLUMNS
    order = {field: index for index, field in enumerate(columns)}
    return tuple(sorted(issues, key=lambda issue: (order.get(issue.field, 99), issue.code)))


def _absent(cell: RawCell) -> bool:
    return cell.kind not in ("formula", "error") and (
        cell.value is None or isinstance(cell.value, str) and not cell.value.strip()
    )


def _source_row_number(text: str) -> int:
    # Unlike money tokens, provenance row text has no specification length cap.
    # Chunk conversion avoids Python's process-wide string-to-int digit limit.
    digits = text.lstrip("0")
    value = 0
    for offset in range(0, len(digits), 1000):
        chunk = digits[offset : offset + 1000]
        value = value * 10 ** len(chunk) + int(chunk)
    return value


def _provenance(raw: RawRecord) -> tuple[Provenance, tuple[Issue, ...]]:
    cells = {name: raw.cells.get(name, RawCell(None)) for name in PROVENANCE_COLUMNS}
    values = dict.fromkeys(PROVENANCE_COLUMNS)
    warnings = []
    for field, cell in cells.items():
        if _absent(cell):
            continue
        value = cell.value
        valid = False
        if field != "source_row":
            valid = cell.kind == "text" and isinstance(value, str)
            if valid:
                values[field] = value
        else:
            if cell.kind == "text" and isinstance(value, str) and re.fullmatch(r"[0-9]+", value):
                values[field] = _source_row_number(value)
                valid = values[field] > 0
            elif cell.kind == "numeric" and type(value) is int:
                values[field] = value
                valid = value > 0
            elif cell.kind == "numeric" and type(value) is float and value.is_integer():
                values[field] = int(value)
                valid = values[field] > 0
        if not valid:
            values[field] = None
            warnings.append(
                Issue(
                    "INVALID_PROVENANCE", field, "Expected optional text or a positive source row"
                )
            )
    if all(_absent(cell) for cell in cells.values()):
        coverage = "absent"
    elif warnings:
        coverage = "invalid"
    elif values["source_file"] is not None and values["source_row"] is not None:
        coverage = "complete"
    else:
        coverage = "partial"
        warnings.append(
            Issue(
                "INCOMPLETE_PROVENANCE",
                "provenance",
                "Source file and source row are not both supplied",
            )
        )
    # An explicitly supplied empty sheet stays distinguishable from a missing column.
    if "source_sheet" in raw.cells and cells["source_sheet"].value == "":
        values["source_sheet"] = ""
    return Provenance(**values, coverage=coverage), sort_issues(warnings)


def validate_row(raw: RawRecord, config: CurrencyConfig) -> ValidatedRow:
    """Validate independent required fields, preserving raw data and known money."""
    provenance, warnings = _provenance(raw)
    values = dict.fromkeys(BUSINESS_COLUMNS)
    issues = []
    if all(_absent(cell) for cell in raw.cells.values()):
        return ValidatedRow(
            raw,
            **values,
            amount_minor=None,
            provenance=provenance,
            issues=(Issue("BLANK_ROW", "row", "Fully blank input record"),),
            warnings=warnings,
        )
    available = {}
    for field in BUSINESS_COLUMNS:
        cell = raw.cells[field]
        if cell.kind in ("formula", "error"):
            code = "FORMULA_NOT_ALLOWED" if cell.kind == "formula" else "EXCEL_CELL_ERROR"
            issues.append(Issue(code, field, "Mapped formula/error cells cannot be trusted"))
        elif _absent(cell):
            issues.append(Issue("MISSING_REQUIRED_VALUE", field, "Required value is empty"))
        elif cell.kind != "text" or not isinstance(cell.value, str):
            code = (
                "ID_NOT_TEXT"
                if field in ("branch_id", "transaction_id", "customer_id")
                else "INVALID_CELL_TYPE"
            )
            issues.append(Issue(code, field, "Business values must be stored as text"))
        else:
            available[field] = cell.value
    for field in ("branch_id", "transaction_id", "customer_id"):
        values[field] = available.get(field)
    if "currency" in available:
        currency = available["currency"]
        if re.fullmatch(r"[A-Z]{3}", currency) and currency in config.precision:
            values["currency"] = currency
        else:
            issues.append(
                Issue("UNSUPPORTED_CURRENCY", "currency", "Currency has no declared precision")
            )
    if "transaction_date" in available:
        text = available["transaction_date"]
        try:
            if not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", text):
                raise ValueError
            parsed = date.fromisoformat(text)
            if parsed.isoformat() != text:
                raise ValueError
            values["transaction_date"] = parsed
        except ValueError:
            issues.append(
                Issue("INVALID_DATE", "transaction_date", "Expected a valid YYYY-MM-DD date")
            )
    amount_minor = None
    if "amount" in available:
        try:
            values["amount"] = parse_decimal(available["amount"])
            if values["currency"] is not None:
                amount_minor = to_minor_units(
                    values["amount"], config.precision[values["currency"]]
                )
        except MoneyError as exc:
            issues.append(Issue(exc.code, "amount", str(exc)))
    return ValidatedRow(
        raw,
        **values,
        amount_minor=amount_minor,
        provenance=provenance,
        issues=sort_issues(issues),
        warnings=warnings,
    )


def validate_transactions(
    source: TransactionInput, config: CurrencyConfig | None = None
) -> ValidatedBatch:
    """Validate every occurrence; no duplicate removal or group decisions yet."""
    if source.mode != "transactions":
        raise ConfigurationError("validate_transactions requires transaction-mode input")
    resolved = config or CurrencyConfig()
    return ValidatedBatch(
        source, resolved, tuple(validate_row(raw, resolved) for raw in source.records)
    )
