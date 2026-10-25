"""Standalone transaction reporting CLI with explicit review/failure statuses."""

import argparse
import json
import re
import sys
from collections.abc import Sequence
from pathlib import Path

from .config import CurrencyConfig
from .errors import ConfigurationError, InputError, IntegrityError, ReportError
from .publication import generate_report
from .version import VERSION


class _Parser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        self.print_usage(sys.stderr)
        self.exit(2, f"salesreport-auto: error: {message}\n")


def _precision_definitions(values: list[str]) -> CurrencyConfig:
    pairs = []
    for value in values:
        match = re.fullmatch(r"([A-Z]{3})=([0-9])", value)
        if match is None:
            raise ConfigurationError(
                "--currency-precision requires CODE=N, uppercase CODE and N=0..9"
            )
        pairs.append((match[1], int(match[2])))
    return CurrencyConfig(pairs)


def main(argv: Sequence[str] | None = None) -> int:
    """Generate an explicit-mode workbook; return 0 completed, 1 review, or 2 error."""
    parser = _Parser(
        prog="salesreport-auto",
        allow_abbrev=False,
        description=("Create a verified standalone sales Excel report without overwriting files."),
        epilog=(
            "Required columns: branch_id, transaction_id, transaction_date (YYYY-MM-DD), "
            "customer_id, amount (plain decimal text), currency. Optional: source_file, "
            "source_sheet, source_row. Products additionally require line_id, product_id, "
            "quantity (signed whole units), unit_price (nonnegative), and optionally product_name. "
            "Product line amount must equal quantity times unit price exactly. "
            "XLSX business cells must be text. Transactions select exact Master or require one "
            "worksheet; products require exactly one worksheet. "
            "Defaults: USD=2, PLN=2, JPY=0. Exit codes: "
            "0 completed, 1 published for review, 2 failure. "
            "--bundle verifies a SheetMerge version-1 bundle in transaction mode only; "
            "its currency precision is authoritative. "
            "No authenticity/business completeness is implied."
        ),
    )
    parser.add_argument("--version", action="version", version=f"SalesReport Auto {VERSION}")
    parser.add_argument(
        "--mode", required=True, choices=("transactions", "products"), help="Explicit report mode"
    )
    parser.add_argument("--input", required=True, type=Path, help="One CSV or text-cell XLSX input")
    parser.add_argument(
        "--bundle", type=Path, help="Optional SheetMerge directory; transactions only"
    )
    parser.add_argument(
        "--output", required=True, type=Path, help="New .xlsx file; parent must exist"
    )
    parser.add_argument(
        "--currency-precision",
        action="append",
        default=[],
        metavar="CODE=N",
        help="Extend/override currency precision (repeat for different codes)",
    )
    args = parser.parse_args(argv)
    try:
        config = _precision_definitions(args.currency_precision)
        result = generate_report(
            args.input, args.output, mode=args.mode, config=config, bundle=args.bundle
        )
    except (ConfigurationError, InputError, IntegrityError, ReportError, OSError) as exc:
        print(f"salesreport-auto: error: {exc}", file=sys.stderr)
        for note in getattr(exc, "__notes__", ()):
            print(f"salesreport-auto: warning: {note}", file=sys.stderr)
        return 2
    if result.cleanup_warning:
        print(
            f"salesreport-auto: warning: Report published. {result.cleanup_warning}",
            file=sys.stderr,
        )
    print(json.dumps(result.to_summary(), ensure_ascii=True, sort_keys=True, allow_nan=False))
    return 1 if result.status == "review_required" else 0
