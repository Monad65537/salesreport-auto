"""Actual SheetMerge manifest-v1, reconciliation JSON, and exclusions CSV schemas."""

import csv
import io
import re
from dataclasses import dataclass
from types import MappingProxyType

from .bundle_io import strict_json
from .bundle_models import (
    BUNDLE_OUTPUTS,
    SourceInput,
    SourceRef,
    UpstreamExclusion,
    UpstreamReconciliation,
    freeze_json,
)
from .config import CurrencyConfig
from .errors import BundleError, ConfigurationError
from .models import ALL_COLUMNS, Issue
from .money import format_minor
from .quality import CurrencyLedger, OutcomeAmounts, OutcomeCounts
from .readers import _CSV_READER_LOCK

LABELS = ("input", "accepted", "rejected", "redundant")
EXCLUSION_COLUMNS = (
    *ALL_COLUMNS,
    "status",
    "reason_codes",
    "reasons",
    "duplicate_of_file",
    "duplicate_of_sheet",
    "duplicate_of_row",
    "raw_values",
)


def _fail(message: str) -> None:
    raise BundleError("INVALID_BUNDLE", message)


def _object(value: object, required: tuple[str, ...], where: str) -> dict:
    if not isinstance(value, dict) or not set(required) <= value.keys():
        _fail(f"{where}: expected object with required fields {', '.join(required)}")
    return value


def _string(value: object, where: str, *, nonempty: bool = True) -> str:
    if not isinstance(value, str) or nonempty and not value.strip():
        _fail(f"{where}: expected text")
    return value


def _integer(value: object, where: str) -> int:
    if type(value) is not int or value < 0:
        _fail(f"{where}: expected nonnegative integer")
    return value


def _boolean(value: object, expected: bool, where: str) -> None:
    if type(value) is not bool or value != expected:
        _fail(f"{where}: inconsistent or nonboolean flag")


def _digest(value: object, where: str) -> str:
    if not isinstance(value, str) or not re.fullmatch(r"[0-9a-f]{64}", value):
        _fail(f"{where}: expected lowercase SHA-256 hexadecimal digest")
    return value


@dataclass(frozen=True)
class Manifest:
    """Validated supplied declarations, not executable source configuration."""

    tool_version: str
    config: CurrencyConfig
    configuration: dict
    configuration_sha256: str
    inputs: tuple[SourceInput, ...]
    policies: dict
    reconciliation: dict
    output_sha256: dict[str, str]


def parse_manifest(data: bytes) -> Manifest:
    """Validate required version-1 fields while allowing additive object metadata."""
    value = _object(
        strict_json(data),
        (
            "tool",
            "tool_version",
            "manifest_version",
            "business_key",
            "configuration",
            "configuration_sha256",
            "inputs",
            "policies",
            "reconciliation",
            "output_sha256",
        ),
        "manifest",
    )
    if (
        value["tool"] != "sheetmerge"
        or type(value["manifest_version"]) is not int
        or value["manifest_version"] != 1
    ):
        _fail("Unsupported tool or manifest_version; require sheetmerge manifest version 1")
    if value["business_key"] != ["branch_id", "transaction_id"]:
        _fail("Unsupported business_key")
    version = _string(value["tool_version"], "tool_version")
    configuration = _object(
        value["configuration"], ("currency_precision", "profiles", "limits"), "configuration"
    )
    if not isinstance(configuration["profiles"], list) or not isinstance(
        configuration["limits"], dict
    ):
        _fail("Configuration profiles/limits must be an array/object")
    try:
        config = CurrencyConfig.from_bundle(configuration["currency_precision"])
    except ConfigurationError as exc:
        raise BundleError("INVALID_BUNDLE", str(exc)) from exc
    config_hash = _digest(value["configuration_sha256"], "configuration_sha256")
    if not isinstance(value["inputs"], list):
        _fail("inputs must be an array")
    inputs = []
    seen = set()
    for item in value["inputs"]:
        item = _object(item, ("source_file", "sha256", "profile"), "inputs entry")
        name = _string(item["source_file"], "inputs.source_file")
        if name in seen:
            _fail("Duplicate original source filename in inputs inventory")
        seen.add(name)
        inputs.append(
            SourceInput(
                name,
                _digest(item["sha256"], "inputs.sha256"),
                _string(item["profile"], "inputs.profile"),
            )
        )
    policies = _object(
        value["policies"],
        ("duplicates", "conflicts", "invalid_candidates", "blank_rows"),
        "policies",
    )
    for name in ("duplicates", "conflicts", "invalid_candidates", "blank_rows"):
        _string(policies[name], f"policies.{name}")
    hashes = _object(value["output_sha256"], BUNDLE_OUTPUTS, "output_sha256")
    if set(hashes) != set(BUNDLE_OUTPUTS):
        _fail("output_sha256 keys must be exactly the six known artifact basenames")
    for name, digest in hashes.items():
        _digest(digest, f"output_sha256.{name}")
    return Manifest(
        version,
        config,
        configuration,
        config_hash,
        tuple(inputs),
        policies,
        value["reconciliation"],
        hashes,
    )


def _counts(value: object, where: str) -> OutcomeCounts:
    data = _object(value, LABELS, where)
    counts = OutcomeCounts(*(_integer(data[name], f"{where}.{name}") for name in LABELS))
    if not counts.balanced:
        _fail(f"{where}: count partition does not balance")
    return counts


def parse_reconciliation(value: object, config: CurrencyConfig) -> UpstreamReconciliation:
    """Validate actual nested counts, monetary strings, precision, and coverage flags."""
    data = _object(
        value,
        (
            "counts",
            "count_equation",
            "count_balanced",
            "known_amounts_balanced",
            "monetary_complete",
            "unknown_currency_count",
            "unquantified_count",
            "currencies",
        ),
        "reconciliation",
    )
    counts = _counts(data["counts"], "counts")
    if data["count_equation"] != "input = accepted + rejected + redundant":
        _fail("Incorrect count_equation")
    _boolean(data["count_balanced"], True, "count_balanced")
    _boolean(data["known_amounts_balanced"], True, "known_amounts_balanced")
    unknown = _integer(data["unknown_currency_count"], "unknown_currency_count")
    unquantified = _integer(data["unquantified_count"], "unquantified_count")
    if not 0 <= unknown <= unquantified <= counts.input:
        _fail("Global unknown/unquantified counts exceed their input scope")
    _boolean(data["monetary_complete"], unquantified == 0, "monetary_complete")
    if not isinstance(data["currencies"], list):
        _fail("currencies must be an array")
    ledgers = []
    seen = set()
    for item in data["currencies"]:
        item = _object(
            item,
            (
                "currency",
                "precision",
                "counts",
                "amounts",
                "minor_units",
                "unquantified_counts",
                "known_amounts_balanced",
                "monetary_complete",
            ),
            "currency ledger",
        )
        code = _string(item["currency"], "currency")
        precision = _integer(item["precision"], "precision")
        if code not in config.precision or code in seen or precision != config.precision[code]:
            _fail("Currency ledger inventory or precision disagrees with configuration")
        seen.add(code)
        row_counts = _counts(item["counts"], f"{code}.counts")
        row_unknown = _counts(item["unquantified_counts"], f"{code}.unquantified_counts")
        if row_unknown.accepted != 0 or any(
            getattr(row_unknown, key) > getattr(row_counts, key) for key in LABELS
        ):
            _fail(
                "Unquantified currency counts exceed corresponding counts or include accepted rows"
            )
        amounts = _object(item["amounts"], LABELS, f"{code}.amounts")
        minors = _object(item["minor_units"], LABELS, f"{code}.minor_units")
        totals = []
        for label in LABELS:
            token = minors[label]
            if (
                not isinstance(token, str)
                or not re.fullmatch(r"-?(?:0|[1-9][0-9]*)", token)
                or token == "-0"
            ):
                _fail("minor_units values must be canonical signed integer strings")
            try:
                minor = int(token)
            except ValueError as exc:
                raise BundleError("INVALID_BUNDLE", "Unsupported integer metadata size") from exc
            if not isinstance(amounts[label], str) or amounts[label] != format_minor(
                minor, precision
            ):
                _fail("Currency amount strings disagree with exact minor units/precision")
            totals.append(minor)
        money = OutcomeAmounts(*totals)
        if not money.balanced:
            _fail("Currency known-money partition does not balance")
        _boolean(item["known_amounts_balanced"], True, f"{code}.known_amounts_balanced")
        _boolean(item["monetary_complete"], row_unknown.input == 0, f"{code}.monetary_complete")
        ledgers.append(CurrencyLedger(code, precision, row_counts, money, row_unknown))
    if seen != set(config.precision):
        _fail("Currency ledgers must cover every configured currency, including unused currencies")
    if sum(row.counts.input for row in ledgers) + unknown != counts.input:
        _fail("Currency input counts and unknown currency count do not reconcile globally")
    if sum(row.unquantified_counts.input for row in ledgers) + unknown != unquantified:
        _fail("Currency unquantified counts do not reconcile globally")
    for label in LABELS:
        total = sum(getattr(row.counts, label) for row in ledgers)
        if total > getattr(counts, label) or label == "accepted" and total != counts.accepted:
            _fail("Currency outcome counts disagree with global counts")
    return UpstreamReconciliation(
        counts, tuple(sorted(ledgers, key=lambda row: row.currency)), unknown, unquantified
    )


def _source_ref(file: str, sheet: str, row: str) -> SourceRef:
    _string(file, "source_file")
    if not re.fullmatch(r"[0-9]+", row):
        _fail("Source/duplicate row must be positive decimal-digit text")
    number = int(row)
    if number <= 0:
        _fail("Source/duplicate row must be positive")
    return SourceRef(file, sheet, number)


def parse_exclusions(data: bytes) -> tuple[UpstreamExclusion, ...]:
    """Parse exclusion strings and structured reasons without sales-field validation."""
    try:
        with _CSV_READER_LOCK:
            previous = csv.field_size_limit()
            try:
                csv.field_size_limit(max(previous, len(data)))
                reader = csv.reader(io.StringIO(data.decode("utf-8"), newline=""), strict=True)
                if tuple(next(reader, ())) != EXCLUSION_COLUMNS:
                    _fail("exclusions.csv headers must match the SheetMerge schema exactly")
                records = []
                for row in reader:
                    if len(row) != len(EXCLUSION_COLUMNS):
                        _fail("Malformed exclusions.csv record width")
                    values = dict(zip(EXCLUSION_COLUMNS, row, strict=True))
                    status = values["status"]
                    if status not in ("rejected", "redundant"):
                        _fail("Unsupported upstream exclusion status")
                    codes = strict_json(values["reason_codes"])
                    reasons_data = strict_json(values["reasons"])
                    if (
                        not isinstance(codes, list)
                        or not codes
                        or any(not isinstance(code, str) or not code.strip() for code in codes)
                    ):
                        _fail("reason_codes must be a nonempty array of nonempty strings")
                    if not isinstance(reasons_data, list) or not reasons_data:
                        _fail("reasons must be a nonempty array")
                    reasons = []
                    for reason in reasons_data:
                        reason = _object(reason, ("code", "field", "message"), "reason")
                        reasons.append(
                            Issue(
                                _string(reason["code"], "reason.code"),
                                _string(reason["field"], "reason.field", nonempty=False),
                                _string(reason["message"], "reason.message", nonempty=False),
                            )
                        )
                    if list(codes) != [issue.code for issue in reasons]:
                        _fail("reason_codes and reasons disagree")
                    raw = _object(strict_json(values["raw_values"]), (), "raw_values")
                    source = _source_ref(
                        values["source_file"], values["source_sheet"], values["source_row"]
                    )
                    duplicate = None
                    if status == "redundant":
                        duplicate = _source_ref(
                            values["duplicate_of_file"],
                            values["duplicate_of_sheet"],
                            values["duplicate_of_row"],
                        )
                    elif any(
                        values[name]
                        for name in ("duplicate_of_file", "duplicate_of_sheet", "duplicate_of_row")
                    ):
                        _fail("Rejected exclusions cannot supply duplicate references")
                    records.append(
                        UpstreamExclusion(
                            source,
                            status,
                            MappingProxyType({name: values[name] for name in ALL_COLUMNS}),
                            tuple(codes),
                            tuple(reasons),
                            duplicate,
                            freeze_json(raw),
                        )
                    )
                return tuple(sorted(records, key=lambda row: row.source))
            finally:
                csv.field_size_limit(previous)
    except (ValueError, UnicodeError, csv.Error) as exc:
        if isinstance(exc, BundleError):
            raise
        raise BundleError("INVALID_BUNDLE", "Malformed exclusions.csv data") from exc
