"""Typed worksheet content derived from the existing exact transaction results."""

import json
from collections.abc import Mapping
from dataclasses import asdict, dataclass, is_dataclass, replace
from datetime import date, datetime, time
from decimal import Decimal
from typing import Literal

from .aggregation import SalesMetrics, aggregate_transactions
from .bundle import verify_master_agreement
from .bundle_models import VERIFICATION_SCOPE
from .engine import TransactionAnalysis
from .errors import IntegrityError
from .models import ReportMode
from .money import format_minor
from .quality import reconcile_transactions

SHEET_NAMES = ("Summary", "Monthly Sales", "Branches", "Data Quality", "Raw Data")
PRODUCT_SHEET_NAMES = (
    "Summary",
    "Monthly Sales",
    "Branches",
    "Products",
    "Data Quality",
    "Raw Data",
)
SPEC_VERSION = "1.0"
CHART_INTEGER_LIMIT = 999_999_999_999_999
UNKNOWN = "Unknown / not supplied"
SCOPE_STATEMENT = (
    "Metrics include accepted records only. Precise totals do not imply that all source "
    "records were accepted."
)
UPSTREAM_STATEMENT = "Upstream acceptance and monetary completeness were not supplied."
CellValue = str | int | bool | None


def text(value: object) -> str:
    """Render decoded raw metadata as text, never as a financial operand."""
    if value is None:
        return ""
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    return str(value)


def _audit_value(value: object) -> object:
    if is_dataclass(value) and not isinstance(value, type):
        return _audit_value(asdict(value))
    if isinstance(value, Mapping):
        return {key: _audit_value(item) for key, item in value.items()}
    if isinstance(value, (tuple, list)):
        return [_audit_value(item) for item in value]
    if isinstance(value, (datetime, date, time, Decimal)):
        return text(value)
    if isinstance(value, float):
        return repr(value)  # Unsupported decoded XLSX cells: audit metadata only.
    return value


def audit_json(value: object) -> str:
    """Serialize structured diagnostics deterministically, without float operands."""
    return json.dumps(_audit_value(value), ensure_ascii=False, sort_keys=True, allow_nan=False)


def _safe_cell(value: CellValue) -> CellValue:
    if type(value) is int and abs(value) > CHART_INTEGER_LIMIT:
        return str(value)
    return value


@dataclass(frozen=True)
class TablePlan:
    """A contiguous table with one text header row and an optional body."""

    name: str
    header_row: int
    end_row: int
    headers: tuple[str, ...]


@dataclass(frozen=True)
class ChartPlan:
    """One currency's exact helper range on the chart's own worksheet."""

    currency: str
    kind: Literal["monthly", "branches", "products"]
    title: str
    axis_title: str
    first_row: int
    last_row: int
    anchor: str
    category_column: int = 11
    value_column: int = 12
    currency_column: int = 10


@dataclass(frozen=True)
class SheetPlan:
    """All authoritative cells, table ranges, and isolated chart definitions."""

    title: str
    rows: tuple[tuple[CellValue, ...], ...]
    tables: tuple[TablePlan, ...]
    section_rows: tuple[int, ...]
    charts: tuple[ChartPlan, ...] = ()


@dataclass(frozen=True)
class WorkbookPlan:
    """Report-level status and deterministic worksheets for an explicit mode."""

    sheets: tuple[SheetPlan, ...]
    review_reasons: tuple[str, ...]
    omitted_charts: tuple[str, ...]
    mode: ReportMode = "transactions"

    @property
    def status(self) -> str:
        """Return publication status including report-only chart limitations."""
        return "review_required" if self.review_reasons else "completed"


def sheet_names(mode: ReportMode) -> tuple[str, ...]:
    """Return the exact visible sheet sequence required for a supported mode."""
    if mode not in ("transactions", "products"):
        raise IntegrityError("Unknown workbook mode")
    return PRODUCT_SHEET_NAMES if mode == "products" else SHEET_NAMES


class _SheetBuilder:
    def __init__(self, title: str) -> None:
        self.title = title
        self.rows: list[list[CellValue]] = [
            [title],
            [SCOPE_STATEMENT],
            [UPSTREAM_STATEMENT],
            [],
        ]
        self.tables: list[TablePlan] = []
        self.sections: list[int] = []
        self.charts: list[ChartPlan] = []

    def section(
        self, label: str, name: str, headers: tuple[str, ...], rows: list[tuple[CellValue, ...]]
    ) -> TablePlan:
        self.rows.append([label])
        self.sections.append(len(self.rows))
        self.rows.append(list(headers))
        first = len(self.rows)
        self.rows.extend([[_safe_cell(value) for value in row] for row in rows])
        table = TablePlan(name, first, len(self.rows), headers)
        self.tables.append(table)
        self.rows.append([])
        return table

    def finish(self) -> SheetPlan:
        while not self.rows[-1]:
            self.rows.pop()
        return SheetPlan(
            self.title,
            tuple(tuple(row) for row in self.rows),
            tuple(self.tables),
            tuple(self.sections),
            tuple(self.charts),
        )


_METRIC_HEADERS = (
    "accepted_record_count",
    "transaction_count",
    "unique_customer_count",
    "net_revenue",
    "average_transaction_value",
)


def _metric_values(metrics: SalesMetrics) -> tuple[CellValue, ...]:
    return (
        metrics.accepted_record_count,
        metrics.transaction_count,
        metrics.unique_customer_count,
        metrics.net_revenue,
        metrics.average_transaction_value,
    )


def _sales_sheet(analysis: TransactionAnalysis, *, monthly: bool) -> tuple[SheetPlan, list[str]]:
    builder = _SheetBuilder("Monthly Sales" if monthly else "Branches")
    groups = analysis.aggregates.monthly if monthly else analysis.aggregates.branches
    kind = "monthly" if monthly else "branches"
    category = "month" if monthly else "branch_id"
    categories = [group.month if monthly else group.branch_id for group in groups]
    table = builder.section(
        "Accepted transactions by calendar month" if monthly else "Accepted transactions by branch",
        "MonthlyMetrics" if monthly else "BranchMetrics",
        ("currency", category, *_METRIC_HEADERS),
        [
            (group.currency, label, *_metric_values(group.metrics))
            for group, label in zip(groups, categories, strict=True)
        ],
    )
    # Helpers line up with primary table rows, so each currency is contiguous.
    helper_header = ["chart_currency", f"chart_{category}", "chart_net_revenue_minor"]
    builder.rows[table.header_row - 1].extend([None, None, *helper_header])
    blocks: dict[str, list[tuple[int, int, int]]] = {}
    for index, (group, label) in enumerate(zip(groups, categories, strict=True)):
        row_number = table.header_row + 1 + index
        builder.rows[row_number - 1].extend(
            [
                None,
                None,
                group.currency,
                label,
                _safe_cell(group.metrics.net_revenue_minor),
            ]
        )
        blocks.setdefault(group.currency, []).append(
            (row_number, group.metrics.net_revenue_minor, group.metrics.precision)
        )
    omitted = []
    for index, (currency, values) in enumerate(blocks.items()):
        if any(abs(amount) > CHART_INTEGER_LIMIT for _, amount, _ in values):
            omitted.append(f"{builder.title}: {currency}")
            continue
        scale = 10 ** values[0][2]
        builder.charts.append(
            ChartPlan(
                currency,
                kind,
                f"{'Monthly' if monthly else 'Branch'} net revenue — {currency}",
                f"Net revenue ({currency} minor units; {scale} = 1 {currency})",
                values[0][0],
                values[-1][0],
                f"N{6 + index * 17}",
            )
        )
    return builder.finish(), omitted


def _summary(
    analysis: TransactionAnalysis, plan_status: str, reasons: tuple[str, ...], version: str
) -> SheetPlan:
    builder = _SheetBuilder("Summary")
    quality = analysis.quality
    source = analysis.classified.validated.source
    builder.section(
        "Accepted product lines and transactions by currency"
        if analysis.mode == "products"
        else "Accepted transactions by currency",
        "SummaryMetrics",
        ("currency", "precision", *_METRIC_HEADERS),
        [
            (group.currency, group.metrics.precision, *_metric_values(group.metrics))
            for group in analysis.aggregates.summary
        ],
    )
    context = [
        ("status", plan_status),
        ("review_reasons", audit_json(reasons)),
        ("mode", analysis.mode),
        (
            "record_grain",
            "Product line occurrences"
            if analysis.mode == "products"
            else "Transaction occurrences",
        ),
        ("application_version", version),
        ("spec_version", SPEC_VERSION),
        ("input_file", source.path.name),
        ("input_sha256", source.sha256),
        ("selected_sheet", source.selected_sheet),
        ("ignored_sheets", audit_json(source.ignored_sheets)),
        *((f"local_{name}_count", value) for name, value in asdict(quality.counts).items()),
        ("local_unknown_currency_count", quality.unknown_currency_count),
        ("local_unquantified_count", quality.unquantified_count),
        ("local_monetary_complete", quality.local_monetary_complete),
        ("local_count_balanced", quality.count_balanced),
        ("local_known_amounts_balanced", quality.known_amounts_balanced),
        ("provenance_coverage", audit_json(quality.provenance)),
        ("provenance_verification", "Unverified standalone references"),
        ("upstream_status", UNKNOWN),
        *(
            (f"upstream_{name}_count", UNKNOWN)
            for name in ("input", "accepted", "rejected", "redundant")
        ),
        ("upstream_monetary_complete", UNKNOWN),
        ("upstream_acceptance_complete", UNKNOWN),
        ("upstream_unknown_currency_count", UNKNOWN),
        ("upstream_unquantified_count", UNKNOWN),
        ("upstream_count_balanced", UNKNOWN),
        ("upstream_known_amounts_balanced", UNKNOWN),
        ("revenue_semantics", "Signed recorded amounts; no gross/net/tax/refund-link inference"),
        ("average_rounding", "Final ROUND_HALF_UP to currency precision; ties away from zero"),
        ("customer_identity", "Distinct (branch_id, customer_id) pairs"),
    ]
    upstream = analysis.upstream
    if upstream is not None:
        report = upstream.reconciliation
        replacements = {
            "provenance_verification": (
                "Internally consistent with supplied bundle; original sources not verified"
            ),
            "upstream_status": "review_required" if upstream.review_reasons else "completed",
            **{f"upstream_{key}_count": value for key, value in asdict(report.counts).items()},
            "upstream_monetary_complete": report.monetary_complete,
            "upstream_acceptance_complete": upstream.acceptance_complete,
            "upstream_unknown_currency_count": report.unknown_currency_count,
            "upstream_unquantified_count": report.unquantified_count,
            "upstream_count_balanced": True,
            "upstream_known_amounts_balanced": True,
        }
        context = [(key, replacements.get(key, value)) for key, value in context]
        context.extend(
            [
                ("bundle_directory", str(upstream.directory)),
                ("bundle_tool", "sheetmerge"),
                ("bundle_tool_version", upstream.tool_version),
                ("bundle_manifest_version", upstream.manifest_version),
                ("bundle_manifest_snapshot_sha256", upstream.manifest_sha256),
                ("declared_configuration_sha256", upstream.configuration_sha256),
                ("bundle_verification_scope", VERIFICATION_SCOPE),
            ]
        )
    builder.section("Report context and scope", "SummaryContext", ("field", "value"), context)
    config = analysis.classified.validated.config
    builder.section(
        "Effective currency precision",
        "SummaryPrecision",
        ("currency", "precision", "origin"),
        [(code, digits, config.origins[code]) for code, digits in sorted(config.precision.items())],
    )
    return builder.finish()


def _quality(
    analysis: TransactionAnalysis, reasons: tuple[str, ...], omitted: tuple[str, ...]
) -> SheetPlan:
    builder = _SheetBuilder("Data Quality")
    quality = analysis.quality
    fields = (
        "input",
        "accepted",
        "rejected",
        "redundant",
        "unknown_currency_count",
        "unquantified_count",
        "monetary_complete",
        "count_balanced",
        "known_amounts_balanced",
    )
    builder.section(
        "Local imported occurrences",
        "LocalQuality",
        ("scope", *fields),
        [
            (
                "local",
                *asdict(quality.counts).values(),
                quality.unknown_currency_count,
                quality.unquantified_count,
                quality.local_monetary_complete,
                quality.count_balanced,
                quality.known_amounts_balanced,
            )
        ],
    )
    builder.section(
        "Upstream quality — not supplied"
        if analysis.upstream is None
        else "Verified upstream declarations",
        "UpstreamQuality",
        ("scope", *fields),
        [("upstream", *(UNKNOWN for _ in fields))]
        if analysis.upstream is None
        else [
            (
                "upstream",
                *asdict(analysis.upstream.reconciliation.counts).values(),
                analysis.upstream.reconciliation.unknown_currency_count,
                analysis.upstream.reconciliation.unquantified_count,
                analysis.upstream.reconciliation.monetary_complete,
                True,
                True,
            )
        ],
    )
    labels = ("input", "accepted", "rejected", "redundant")
    ledger_headers = (
        "scope",
        "currency",
        "precision",
        *(f"{s}_count" for s in labels),
        *(f"{s}_known_amount" for s in labels),
        *(f"{s}_unquantified_count" for s in labels),
        "known_amounts_balanced",
        "monetary_complete",
    )
    builder.section(
        "Local currency ledgers (input occurrences include exclusions)",
        "LocalLedgers",
        ledger_headers,
        [
            (
                "local",
                ledger.currency,
                ledger.precision,
                *asdict(ledger.counts).values(),
                *(
                    format_minor(amount, ledger.precision)
                    for amount in asdict(ledger.known_amount_minor).values()
                ),
                *asdict(ledger.unquantified_counts).values(),
                ledger.known_amount_minor.balanced,
                ledger.monetary_complete,
            )
            for ledger in quality.currencies
        ],
    )
    builder.section(
        "Provenance coverage — references are unverified"
        if analysis.upstream is None
        else "Provenance coverage — internally consistent bundle references",
        "ProvenanceCoverage",
        ("coverage", "count"),
        list(asdict(quality.provenance).items()),
    )
    if analysis.upstream is not None:
        _upstream_quality_sections(builder, analysis, ledger_headers)
    source = analysis.classified.validated.source
    builder.section(
        "Input worksheet selection (informational)",
        "ImportContext",
        ("field", "value"),
        [
            ("selected_sheet", source.selected_sheet),
            ("ignored_sheets", audit_json(source.ignored_sheets)),
        ],
    )
    builder.section(
        "Report review reasons",
        "ReportReasons",
        ("code", "context"),
        [
            (
                code,
                audit_json(omitted)
                if code == "CHART_PRECISION_LIMIT"
                else "Upstream source quality"
                if code.startswith("UPSTREAM_")
                else "Local import quality",
            )
            for code in reasons
        ],
    )
    details = []
    for item in sorted(analysis.classified.rows, key=lambda item: item.row.raw.location):
        row = item.row
        if item.status == "accepted" and not row.warnings:
            continue
        ref = row.raw.location
        raw = row.raw.raw_values
        details.append(
            (
                "local",
                ref.import_file,
                ref.import_sheet,
                ref.import_row,
                *(text(raw.get(field)) for field in ("source_file", "source_sheet", "source_row")),
                item.status,
                audit_json([i.code for i in item.issues]),
                audit_json(item.issues),
                audit_json(row.warnings),
                audit_json(item.duplicate_of),
                audit_json(raw),
            )
        )
    if analysis.upstream is not None:
        for row in analysis.upstream.exclusions:
            details.append(
                (
                    "upstream",
                    "",
                    "",
                    "",
                    row.source.source_file,
                    row.source.source_sheet,
                    row.values["source_row"],
                    row.status,
                    audit_json(row.reason_codes),
                    audit_json(row.reasons),
                    "[]",
                    audit_json(row.duplicate_of),
                    audit_json(row.raw_values),
                )
            )
    builder.section(
        "Record-level exclusions and warnings",
        "QualityDetails",
        (
            "scope",
            "import_file",
            "import_sheet",
            "import_row",
            "source_file",
            "source_sheet",
            "source_row",
            "status",
            "reason_codes",
            "issues",
            "warnings",
            "duplicate_of",
            "raw_values",
        ),
        details,
    )
    return builder.finish()


def _upstream_quality_sections(
    builder: _SheetBuilder, analysis: TransactionAnalysis, headers: tuple[str, ...]
) -> None:
    upstream = analysis.upstream
    builder.section(
        "Upstream declared currency ledgers",
        "UpstreamLedgers",
        headers,
        [
            (
                "upstream",
                ledger.currency,
                ledger.precision,
                *asdict(ledger.counts).values(),
                *(
                    format_minor(amount, ledger.precision)
                    for amount in asdict(ledger.known_amount_minor).values()
                ),
                *asdict(ledger.unquantified_counts).values(),
                True,
                ledger.monetary_complete,
            )
            for ledger in upstream.reconciliation.currencies
        ],
    )
    builder.section(
        "Bundle verification scope and supplied configuration",
        "BundleVerification",
        ("field", "value"),
        [
            ("scope", VERIFICATION_SCOPE),
            ("tool_version", upstream.tool_version),
            ("manifest_version", upstream.manifest_version),
            ("manifest_snapshot_sha256", upstream.manifest_sha256),
            ("declared_configuration_sha256", upstream.configuration_sha256),
            ("configuration", audit_json(upstream.configuration)),
            ("policies", audit_json(upstream.policies)),
            ("upstream_acceptance_complete", upstream.acceptance_complete),
        ],
    )
    builder.section(
        "Declared source inventory — files were not opened",
        "BundleSources",
        ("source_file", "sha256", "profile"),
        [(row.source_file, row.sha256, row.profile) for row in upstream.inputs],
    )
    builder.section(
        "Verified artifact digests against supplied manifest",
        "BundleArtifacts",
        ("member", "sha256"),
        list(sorted(upstream.output_sha256.items())),
    )
    builder.section(
        "Upstream reason occurrences (codes may overlap)",
        "UpstreamReasons",
        ("code", "occurrence_count"),
        list(upstream.reason_counts.items()),
    )
    builder.section(
        "Upstream canonical exclusions — never included in sales metrics",
        "UpstreamExclusions",
        (
            *analysis.classified.validated.source.headers,
            "status",
            "reason_codes",
            "reasons",
            "duplicate_of",
            "raw_values",
        ),
        [
            (
                *(row.values[key] for key in analysis.classified.validated.source.headers),
                row.status,
                audit_json(row.reason_codes),
                audit_json(row.reasons),
                audit_json(row.duplicate_of),
                audit_json(row.raw_values),
            )
            for row in upstream.exclusions
        ],
    )


def _raw_data(analysis: TransactionAnalysis) -> SheetPlan:
    builder = _SheetBuilder("Raw Data")
    source = analysis.classified.validated.source
    config = analysis.classified.validated.config
    appended = (
        "report_import_file",
        "report_import_sheet",
        "report_import_row",
        "report_status",
        "report_reason_codes",
        "report_issues",
        "report_warnings",
        "report_duplicate_of",
        "report_amount_minor",
        "report_currency_precision",
        "report_normalized_date",
        "report_raw_cell_kinds",
    )
    if source.mode == "products":
        appended += ("report_quantity", "report_unit_price_minor")
    rows = []
    for item in analysis.classified.rows:
        row = item.row
        ref = row.raw.location
        rows.append(
            (
                *(text(row.raw.cells[header].value) for header in source.headers),
                ref.import_file,
                ref.import_sheet,
                ref.import_row,
                item.status,
                audit_json([issue.code for issue in item.issues]),
                audit_json(item.issues),
                audit_json(row.warnings),
                audit_json(item.duplicate_of),
                text(row.amount_minor),
                config.precision[row.currency] if row.currency is not None else "",
                text(row.transaction_date),
                audit_json({key: cell.kind for key, cell in row.raw.cells.items()}),
                *(
                    (row.quantity if row.quantity is not None else "", text(row.unit_price_minor))
                    if source.mode == "products"
                    else ()
                ),
            )
        )
    builder.section(
        "Every imported occurrence in physical order",
        "ImportedRecords",
        (*source.headers, *appended),
        rows,
    )
    return builder.finish()


def _products_sheet(analysis: TransactionAnalysis) -> tuple[SheetPlan, list[str]]:
    builder = _SheetBuilder("Products")
    products = analysis.aggregates.products
    table = builder.section(
        "Accepted product lines — ranked within each currency",
        "ProductMetrics",
        (
            "currency",
            "revenue_rank",
            "product_id",
            "product_name",
            "accepted_line_count",
            "transaction_count",
            "unique_customer_count",
            "positive_quantity",
            "negative_quantity_magnitude",
            "net_quantity",
            "net_revenue",
            "minimum_unit_price",
            "maximum_unit_price",
            "positive_quantity_weighted_unit_price",
        ),
        [
            (
                row.currency,
                row.revenue_rank,
                row.product_id,
                row.product_name,
                row.accepted_line_count,
                row.transaction_count,
                row.unique_customer_count,
                row.positive_quantity,
                row.negative_quantity_magnitude,
                row.net_quantity,
                row.net_revenue,
                row.minimum_unit_price,
                row.maximum_unit_price,
                row.positive_quantity_weighted_unit_price,
            )
            for row in products
        ],
    )
    builder.rows[table.header_row - 1].extend(
        [None, None, "chart_currency", "chart_product_id", "chart_net_revenue_minor"]
    )
    blocks = {}
    for index, row in enumerate(products):
        number = table.header_row + 1 + index
        builder.rows[number - 1].extend(
            [
                None,
                None,
                row.currency,
                row.product_id,
                _safe_cell(row.net_revenue_minor),
            ]
        )
        if row.revenue_rank <= 10:
            blocks.setdefault(row.currency, []).append(
                (number, row.net_revenue_minor, row.precision)
            )
    omitted = []
    for index, (currency, values) in enumerate(blocks.items()):
        if any(abs(amount) > CHART_INTEGER_LIMIT for _, amount, _ in values):
            omitted.append(f"Products: {currency}")
            continue
        builder.charts.append(
            ChartPlan(
                currency,
                "products",
                f"Top products by net revenue — {currency}",
                f"Net revenue ({currency} minor units; {10 ** values[0][2]} = 1 {currency})",
                values[0][0],
                values[-1][0],
                f"U{6 + index * 17}",
                category_column=18,
                value_column=19,
                currency_column=17,
            )
        )
    return builder.finish(), omitted


def _verify_rollup_partitions(analysis: TransactionAnalysis) -> None:
    summary = {group.currency: group.metrics for group in analysis.aggregates.summary}
    accepted_counts = {
        ledger.currency: ledger.counts.accepted for ledger in analysis.quality.currencies
    }
    if (
        sum(metrics.accepted_record_count for metrics in summary.values())
        != analysis.quality.counts.accepted
    ):
        raise IntegrityError("Summary does not conserve accepted occurrences")
    for currency, metrics in summary.items():
        if metrics.accepted_record_count != accepted_counts[currency]:
            raise IntegrityError("Summary disagrees with accepted currency ledger")
        for grouping in (analysis.aggregates.monthly, analysis.aggregates.branches):
            members = [row.metrics for row in grouping if row.currency == currency]
            for field in ("accepted_record_count", "transaction_count", "net_revenue_minor"):
                if sum(getattr(row, field) for row in members) != getattr(metrics, field):
                    raise IntegrityError("Monthly/branch metrics do not partition the summary")
        if analysis.mode == "products":
            members = [row for row in analysis.aggregates.products if row.currency == currency]
            if (
                sum(row.accepted_line_count for row in members) != metrics.accepted_record_count
                or sum(row.net_revenue_minor for row in members) != metrics.net_revenue_minor
            ):
                raise IntegrityError("Product lines and revenue do not partition the summary")


def build_report_plan(analysis: TransactionAnalysis, *, application_version: str) -> WorkbookPlan:
    """Verify existing core results and map them into cells; never use Excel totals."""
    if analysis.upstream is not None:
        verify_master_agreement(analysis.upstream, analysis.classified, analysis.quality)
    if reconcile_transactions(analysis.classified) != analysis.quality:
        raise IntegrityError("Analysis quality disagrees with classified records")
    if aggregate_transactions(analysis.classified) != analysis.aggregates:
        raise IntegrityError("Analysis aggregates disagree with classified records")
    _verify_rollup_partitions(analysis)
    monthly, monthly_omitted = _sales_sheet(analysis, monthly=True)
    branches, branch_omitted = _sales_sheet(analysis, monthly=False)
    products, product_omitted = (
        _products_sheet(analysis) if analysis.mode == "products" else (None, [])
    )
    omitted = tuple(monthly_omitted + branch_omitted + product_omitted)
    present = set(analysis.quality.review_reasons)
    if analysis.upstream is not None:
        present.update(analysis.upstream.review_reasons)
    if omitted:
        present.add("CHART_PRECISION_LIMIT")
    reasons = tuple(
        code
        for code in (
            "NO_ACCEPTED_RECORDS",
            "LOCAL_REJECTIONS",
            "LOCAL_UNQUANTIFIED_AMOUNTS",
            "UPSTREAM_REJECTIONS",
            "UPSTREAM_UNQUANTIFIED_AMOUNTS",
            "INVALID_PROVENANCE",
            "INCOMPLETE_PROVENANCE",
            "PRODUCT_LABEL_CONFLICT",
            "CHART_PRECISION_LIMIT",
        )
        if code in present
    )
    status = "review_required" if reasons else "completed"
    sheets = (
        _summary(analysis, status, reasons, application_version),
        monthly,
        branches,
        *((products,) if products is not None else ()),
        _quality(analysis, reasons, omitted),
        _raw_data(analysis),
    )
    if analysis.upstream is not None:
        sheets = tuple(
            replace(
                sheet,
                rows=(
                    *sheet.rows[:2],
                    (
                        "Supplied bundle is internally consistent; authenticity and business "
                        "completeness are not established.",
                    ),
                    *sheet.rows[3:],
                ),
            )
            for sheet in sheets
        )
    return WorkbookPlan(sheets, reasons, omitted, analysis.classified.validated.source.mode)
