"""SalesReport Auto: exact standalone transaction/product analysis and Excel reports."""

from .aggregation import (
    BranchSales,
    CurrencySummary,
    MonthlySales,
    SalesMetrics,
    TransactionAggregates,
    aggregate_transactions,
    branch_sales,
    monthly_sales,
    summarize_transactions,
)
from .bundle import load_bundle, verify_master_agreement
from .bundle_models import SheetMergeBundle, UpstreamExclusion, UpstreamReconciliation
from .classification import classify_transactions
from .config import CurrencyConfig, InputLimits
from .engine import ProductAnalysis, TransactionAnalysis, analyze_products, analyze_transactions
from .errors import (
    BundleError,
    ConfigurationError,
    InputError,
    IntegrityError,
    MoneyError,
    ReportError,
)
from .models import (
    BUSINESS_COLUMNS,
    PRODUCT_ALL_COLUMNS,
    PRODUCT_COLUMNS,
    PROVENANCE_COLUMNS,
    ClassifiedBatch,
    ClassifiedRow,
    ImportRef,
    Issue,
    Provenance,
    RawCell,
    RawRecord,
    TransactionInput,
    ValidatedBatch,
    ValidatedRow,
)
from .money import MonetaryAverage, format_minor, parse_decimal, parse_money, to_minor_units
from .product_aggregation import ProductSales, product_sales
from .product_classification import classify_products
from .product_validation import validate_product_row, validate_products
from .publication import ReportResult, generate_report, publish_report, validate_output_path
from .quality import (
    CurrencyLedger,
    OutcomeAmounts,
    OutcomeCounts,
    ProvenanceCounts,
    QualityReport,
    reconcile_transactions,
)
from .readers import read_products, read_sales_input, read_transactions
from .report_model import WorkbookPlan, build_report_plan
from .validation import validate_row, validate_transactions
from .version import VERSION
from .workbook import verify_workbook, write_workbook

__version__ = VERSION

__all__ = [
    "BundleError",
    "SheetMergeBundle",
    "UpstreamExclusion",
    "UpstreamReconciliation",
    "load_bundle",
    "verify_master_agreement",
    "BUSINESS_COLUMNS",
    "PROVENANCE_COLUMNS",
    "PRODUCT_COLUMNS",
    "PRODUCT_ALL_COLUMNS",
    "ProductAnalysis",
    "ProductSales",
    "BranchSales",
    "ClassifiedBatch",
    "ClassifiedRow",
    "ConfigurationError",
    "CurrencyConfig",
    "CurrencyLedger",
    "CurrencySummary",
    "ImportRef",
    "InputError",
    "InputLimits",
    "IntegrityError",
    "Issue",
    "MonetaryAverage",
    "MoneyError",
    "MonthlySales",
    "OutcomeAmounts",
    "OutcomeCounts",
    "Provenance",
    "ProvenanceCounts",
    "QualityReport",
    "ReportError",
    "ReportResult",
    "RawCell",
    "RawRecord",
    "SalesMetrics",
    "TransactionAggregates",
    "TransactionAnalysis",
    "TransactionInput",
    "ValidatedBatch",
    "ValidatedRow",
    "WorkbookPlan",
    "aggregate_transactions",
    "analyze_transactions",
    "analyze_products",
    "branch_sales",
    "build_report_plan",
    "classify_transactions",
    "classify_products",
    "format_minor",
    "generate_report",
    "monthly_sales",
    "parse_decimal",
    "parse_money",
    "publish_report",
    "product_sales",
    "read_transactions",
    "read_products",
    "read_sales_input",
    "reconcile_transactions",
    "summarize_transactions",
    "to_minor_units",
    "validate_row",
    "validate_product_row",
    "validate_products",
    "validate_transactions",
    "validate_output_path",
    "verify_workbook",
    "write_workbook",
]
