"""Tiny independent producers of the inspected SheetMerge version-1 file schemas."""

import csv
import hashlib
import json
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook

BUSINESS = ("branch_id", "transaction_id", "transaction_date", "customer_id", "amount", "currency")
MASTER = (*BUSINESS, "source_file", "source_sheet", "source_row")
EXCLUSIONS = (
    *MASTER,
    "status",
    "reason_codes",
    "reasons",
    "duplicate_of_file",
    "duplicate_of_sheet",
    "duplicate_of_row",
    "raw_values",
)
LABELS = ("input", "accepted", "rejected", "redundant")
OUTPUTS = (
    "master.csv",
    "master.xlsx",
    "exclusions.csv",
    "audit.jsonl",
    "reconciliation.csv",
    "reconciliation.json",
)


def _amount(minor, precision):
    digits = str(abs(minor)).zfill(precision + 1)
    return ("-" if minor < 0 else "") + (
        digits if precision == 0 else digits[:-precision] + "." + digits[-precision:]
    )


def _write_csv(path, headers, rows):
    with path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(
            stream, fieldnames=headers, quoting=csv.QUOTE_ALL, lineterminator="\n"
        )
        writer.writeheader()
        writer.writerows(rows)


@dataclass
class SyntheticBundle:
    root: Path

    def manifest(self):
        return json.loads((self.root / "manifest.json").read_text(encoding="utf-8"))

    def write_manifest(self, value):
        (self.root / "manifest.json").write_text(json.dumps(value), encoding="utf-8")

    def resign(self, *, sync_reconciliation=True):
        manifest = self.manifest()
        manifest["output_sha256"] = {
            name: hashlib.sha256((self.root / name).read_bytes()).hexdigest() for name in OUTPUTS
        }
        if sync_reconciliation:
            manifest["reconciliation"] = json.loads(
                (self.root / "reconciliation.json").read_text(encoding="utf-8")
            )
        self.write_manifest(manifest)

    def change_reconciliation(self, mutation):
        data = json.loads((self.root / "reconciliation.json").read_text(encoding="utf-8"))
        mutation(data)
        (self.root / "reconciliation.json").write_text(json.dumps(data), encoding="utf-8")
        self.resign()

    def exclusions(self):
        with (self.root / "exclusions.csv").open(encoding="utf-8", newline="") as stream:
            return list(csv.DictReader(stream))

    def write_exclusions(self, rows):
        _write_csv(self.root / "exclusions.csv", EXCLUSIONS, rows)
        self.resign()


def make_bundle(root, *, variant="review", precision=None):
    root.mkdir()
    precision = {"USD": 2} if precision is None else precision
    code = "USD" if "USD" in precision else next(iter(precision))
    p = precision[code]
    scale = 10**p
    base = dict(
        zip(
            BUSINESS,
            ("001", "000001", "2024-01-02", "000101", _amount(10 * scale, p), code),
            strict=True,
        )
    )
    descriptors = []
    if variant != "header":
        descriptors.append(
            (
                "rejected" if variant == "conflict" else "accepted",
                base,
                10 * scale,
                "CONFLICTING_TRANSACTION" if variant == "conflict" else None,
                None,
            )
        )
    if variant in ("review", "duplicates"):
        descriptors.append(("redundant", base.copy(), 10 * scale, "EXACT_DUPLICATE", 2))
    if variant == "review":
        descriptors.extend(
            [
                (
                    "rejected",
                    {
                        **base,
                        "transaction_id": "bad-date",
                        "transaction_date": "",
                        "amount": _amount(3 * scale, p),
                    },
                    3 * scale,
                    "INVALID_DATE",
                    None,
                ),
                (
                    "rejected",
                    {**base, "transaction_id": "bad-money", "amount": ""},
                    None,
                    "INVALID_AMOUNT",
                    None,
                ),
                (
                    "rejected",
                    {
                        **base,
                        "transaction_id": "bad-currency",
                        "currency": "",
                        "amount": _amount(2 * scale, p),
                    },
                    None,
                    "UNSUPPORTED_CURRENCY",
                    None,
                ),
            ]
        )
    if variant == "conflict":
        descriptors.append(
            (
                "rejected",
                {**base, "amount": _amount(20 * scale, p)},
                20 * scale,
                "CONFLICTING_TRANSACTION",
                None,
            )
        )
    records = []
    for row_number, (status, values, minor, reason, winner) in enumerate(descriptors, 2):
        canonical = {
            **values,
            "source_file": "original.csv",
            "source_sheet": "",
            "source_row": str(row_number),
        }
        source = {"source_file": "original.csv", "source_sheet": "", "source_row": row_number}
        raw = values.copy()
        if reason == "INVALID_DATE":
            raw["transaction_date"] = "bad"
        if reason == "INVALID_AMOUNT":
            raw["amount"] = "NaN"
        if reason == "UNSUPPORTED_CURRENCY":
            raw["currency"] = "EUR" if code != "EUR" else "ZZZ"
        field = {
            "INVALID_DATE": "transaction_date",
            "INVALID_AMOUNT": "amount",
            "UNSUPPORTED_CURRENCY": "currency",
        }.get(reason, "transaction")
        issues = (
            []
            if reason is None
            else [{"code": reason, "field": field, "message": "Synthetic upstream reason"}]
        )
        duplicate = None if winner is None else {**source, "source_row": winner}
        records.append((canonical, status, minor, raw, source, issues, duplicate))
    masters = [record[0] for record in records if record[1] == "accepted"]
    _write_csv(root / "master.csv", MASTER, masters)
    book = Workbook()
    book.active.title = "Master"
    for number, values in enumerate([dict(zip(MASTER, MASTER, strict=True)), *masters], 1):
        for column, name in enumerate(MASTER, 1):
            cell = book.active.cell(number, column, values[name])
            cell.data_type = "s"
    book.save(root / "master.xlsx")
    book.close()
    excluded = []
    audit = []
    for canonical, status, minor, raw, source, issues, duplicate in records:
        if status != "accepted":
            excluded.append(
                {
                    **canonical,
                    "status": status,
                    "reason_codes": json.dumps([issue["code"] for issue in issues]),
                    "reasons": json.dumps(issues),
                    "duplicate_of_file": "" if duplicate is None else duplicate["source_file"],
                    "duplicate_of_sheet": "",
                    "duplicate_of_row": "" if duplicate is None else str(duplicate["source_row"]),
                    "raw_values": json.dumps(raw),
                }
            )
        audit.append(
            {
                "source": source,
                "status": status,
                "values": {key: canonical[key] or None for key in BUSINESS},
                "amount_minor": None if minor is None else str(minor),
                "issues": issues,
                "duplicate_of": duplicate,
                "raw_values": raw,
                "raw_cells": {
                    key: {
                        "value": value,
                        "kind": "text",
                        "numeric_text": None,
                        "invalid_excel_date": False,
                        "original_text": None,
                    }
                    for key, value in raw.items()
                },
            }
        )
    _write_csv(root / "exclusions.csv", EXCLUSIONS, excluded)
    (root / "audit.jsonl").write_text(
        "".join(json.dumps(row) + "\n" for row in audit), encoding="utf-8"
    )
    status_counts = Counter(row[1] for row in records)
    unknown_count = sum(row[0]["currency"] not in precision for row in records)
    unquantified = sum(row[2] is None for row in records)
    currencies = []
    for currency, digits in sorted(precision.items()):
        subset = [row for row in records if row[0]["currency"] == currency]
        counts = {
            "input": len(subset),
            **{status: sum(row[1] == status for row in subset) for status in LABELS[1:]},
        }
        totals = {
            label: sum(
                row[2]
                for row in subset
                if row[2] is not None and (label == "input" or row[1] == label)
            )
            for label in LABELS
        }
        missing = {
            label: sum(row[2] is None for row in subset if label == "input" or row[1] == label)
            for label in LABELS
        }
        currencies.append(
            {
                "currency": currency,
                "precision": digits,
                "counts": counts,
                "amounts": {label: _amount(total, digits) for label, total in totals.items()},
                "minor_units": {label: str(total) for label, total in totals.items()},
                "unquantified_counts": missing,
                "known_amounts_balanced": True,
                "monetary_complete": missing["input"] == 0,
            }
        )
    report = {
        "counts": {
            "input": len(records),
            **{status: status_counts[status] for status in LABELS[1:]},
        },
        "count_equation": "input = accepted + rejected + redundant",
        "count_balanced": True,
        "known_amounts_balanced": True,
        "monetary_complete": unquantified == 0,
        "unknown_currency_count": unknown_count,
        "unquantified_count": unquantified,
        "currencies": currencies,
    }
    (root / "reconciliation.json").write_text(json.dumps(report), encoding="utf-8")
    csv_headers = (
        "currency",
        *(f"{label}_count" for label in LABELS),
        *(f"{label}_known_amount" for label in LABELS),
        *(f"{label}_unquantified_count" for label in LABELS),
        "known_amounts_balanced",
        "monetary_complete",
    )
    _write_csv(
        root / "reconciliation.csv",
        csv_headers,
        [
            {
                "currency": row["currency"],
                **{f"{label}_count": str(row["counts"][label]) for label in LABELS},
                **{f"{label}_known_amount": row["amounts"][label] for label in LABELS},
                **{
                    f"{label}_unquantified_count": str(row["unquantified_counts"][label])
                    for label in LABELS
                },
                "known_amounts_balanced": "true",
                "monetary_complete": str(row["monetary_complete"]).lower(),
            }
            for row in currencies
        ],
    )
    manifest = {
        "tool": "sheetmerge",
        "tool_version": "0.1.0",
        "manifest_version": 1,
        "business_key": ["branch_id", "transaction_id"],
        "configuration_sha256": "a" * 64,
        "configuration": {
            "currency_precision": precision,
            "profiles": [
                {
                    "name": "synthetic",
                    "format": "csv",
                    "patterns": ["*.csv"],
                    "columns": {key: key for key in BUSINESS},
                }
            ],
            "limits": {"max_file_bytes": 20971520},
        },
        "inputs": [{"source_file": "original.csv", "sha256": "1" * 64, "profile": "synthetic"}],
        "policies": {
            "duplicates": "keep earliest file/sheet/row; report others as redundant",
            "conflicts": "reject all occurrences of a conflicting key",
            "invalid_candidates": "reject all occurrences of a key with invalid candidates",
            "blank_rows": "reject with BLANK_ROW",
        },
        "reconciliation": report,
        "output_sha256": {},
    }
    fixture = SyntheticBundle(root)
    fixture.write_manifest(manifest)
    fixture.resign()
    return fixture
