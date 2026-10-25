from dataclasses import replace
from decimal import Decimal, localcontext

import pytest

from salesreport_auto import (
    BUSINESS_COLUMNS,
    CurrencyConfig,
    IntegrityError,
    OutcomeAmounts,
    OutcomeCounts,
    analyze_transactions,
    branch_sales,
    classify_transactions,
    monthly_sales,
    read_transactions,
    reconcile_transactions,
    summarize_transactions,
    validate_transactions,
)


def test_minimal_input_raw_preservation_and_unknown_upstream(write_input, record, transport):
    row = record(branch_id=" 001 ", transaction_id="=1+1", customer_id="Łódź-NA-000001")
    path = write_input([row], format=transport)
    before = path.read_bytes(), path.stat().st_mtime_ns
    result = analyze_transactions(path)
    item = result.classified.accepted[0]
    assert dict(item.row.raw.raw_values) == row
    assert item.row.branch_id == " 001 "
    assert item.row.transaction_id == "=1+1"
    assert item.row.customer_id == "Łódź-NA-000001"
    assert item.row.provenance.coverage == "absent"
    assert result.upstream is None
    assert result.quality.status == "completed"
    assert result.quality.provenance.absent == 1
    assert result.quality.counts == OutcomeCounts(1, 1, 0, 0)
    assert result.aggregates.summary[0].metrics.net_revenue == "12.34"
    assert before == (path.read_bytes(), path.stat().st_mtime_ns)
    assert result.classified.validated.source.headers == BUSINESS_COLUMNS
    with pytest.raises(TypeError):
        item.row.raw.cells["amount"] = None


def test_s03_month_branch_customer_denominators(write_input, record, transport):
    rows = [
        record(branch_id="A", transaction_id="T1", customer_id="C1", amount="10.00"),
        record(branch_id="A", transaction_id="T2", customer_id="C1", amount="-2.00"),
        record(branch_id="A", transaction_id="T3", customer_id="C2", amount="0.00"),
        record(
            branch_id="B",
            transaction_id="T1",
            customer_id="C1",
            amount="3.01",
            transaction_date="2024-02-01",
        ),
    ]
    result = analyze_transactions(write_input(rows, format=transport))
    summary = result.aggregates.summary[0]
    metrics = summary.metrics
    assert summary.currency == "USD"
    assert (
        metrics.accepted_record_count,
        metrics.transaction_count,
        metrics.unique_customer_count,
    ) == (4, 4, 3)
    assert metrics.net_revenue == "11.01"
    assert metrics.average_transaction_value == "2.75"
    assert metrics.average.numerator_minor == 1101
    assert metrics.average.denominator == 4
    jan, feb = result.aggregates.monthly
    assert (jan.month, jan.metrics.transaction_count, jan.metrics.unique_customer_count) == (
        "2024-01",
        3,
        2,
    )
    assert (jan.metrics.net_revenue, jan.metrics.average_transaction_value) == ("8.00", "2.67")
    assert (feb.month, feb.metrics.net_revenue) == ("2024-02", "3.01")
    assert [(g.branch_id, g.metrics.net_revenue) for g in result.aggregates.branches] == [
        ("A", "8.00"),
        ("B", "3.01"),
    ]
    assert (
        sum(g.metrics.net_revenue_minor for g in result.aggregates.branches)
        == metrics.net_revenue_minor
    )
    assert summarize_transactions(result.classified) == result.aggregates.summary
    assert monthly_sales(result.classified) == result.aggregates.monthly
    assert branch_sales(result.classified) == result.aggregates.branches


@pytest.mark.parametrize("amount,expected", [("0.01", "0.01"), ("-0.01", "-0.01")])
def test_s04_transaction_average_half_ties(write_input, record, transport, amount, expected):
    result = analyze_transactions(
        write_input(
            [
                record(transaction_id="1", amount="0.00"),
                record(transaction_id="2", amount=amount),
            ],
            format=transport,
        )
    )
    assert result.aggregates.summary[0].metrics.average_transaction_value == expected


def test_s05_integer_aggregation_beyond_row_size_and_decimal_context(
    write_input, record, transport
):
    rows = [record(transaction_id=f"T{i:02}", amount="0.10") for i in range(10)]
    with localcontext() as context:
        context.prec = 2
        result = analyze_transactions(write_input(rows, format=transport))
        assert result.aggregates.summary[0].metrics.net_revenue == "1.00"
        huge = "9" * 36 + ".99"
        result = analyze_transactions(
            write_input(
                [
                    record(transaction_id="1", amount=huge),
                    record(transaction_id="2", amount=huge),
                ],
                format=transport,
            )
        )
        assert result.aggregates.summary[0].metrics.net_revenue_minor == 2 * (10**38 - 1)
        assert result.aggregates.summary[0].metrics.average_transaction_value == huge
        assert not result.quality.review_required  # Chart limitations belong to a later phase.


@pytest.mark.parametrize(
    "value,code",
    [
        ("1.231", "EXCESS_FRACTIONAL_PRECISION"),
        ("NaN", "INVALID_AMOUNT"),
        ("Infinity", "INVALID_AMOUNT"),
        ("1e2", "INVALID_AMOUNT"),
        ("1,23", "INVALID_AMOUNT"),
        (" 1.00", "INVALID_AMOUNT"),
        ("", "MISSING_REQUIRED_VALUE"),
        ("1" + "0" * 36, "AMOUNT_TOO_LARGE"),
        ("1" * 513, "INVALID_AMOUNT"),
    ],
)
def test_s06_unquantified_invalid_money(write_input, record, transport, value, code):
    result = analyze_transactions(write_input([record(amount=value)], format=transport))
    assert result.quality.counts == OutcomeCounts(1, 0, 1, 0)
    assert code in {issue.code for issue in result.classified.rejected[0].issues}
    assert result.classified.rejected[0].row.raw.raw_values["amount"] == value
    assert result.quality.unquantified_count == 1
    assert not result.quality.local_monetary_complete
    assert result.aggregates.summary == ()


def test_s06_exact_extra_zeroes_and_signed_zero(write_input, record, transport):
    result = analyze_transactions(
        write_input(
            [
                record(transaction_id="1", amount="1.2300"),
                record(transaction_id="2", amount="-0.00"),
            ],
            format=transport,
        )
    )
    assert [item.row.amount_minor for item in result.classified.accepted] == [123, 0]
    assert result.aggregates.summary[0].metrics.net_revenue == "1.23"


@pytest.mark.parametrize(
    "text,valid",
    [
        ("2024-02-29", True),
        ("0001-01-01", True),
        ("9999-12-31", True),
        ("2023-02-29", False),
        ("2024-2-29", False),
        ("2024-02-29T00:00:00", False),
        ("2024-02-29Z", False),
        ("2024/02/29", False),
        ("2024-02-29 ", False),
    ],
)
def test_s07_dates_and_known_rejected_money(write_input, record, transport, text, valid):
    result = analyze_transactions(write_input([record(transaction_date=text)], format=transport))
    assert result.quality.counts.accepted == int(valid)
    assert result.quality.unquantified_count == 0
    assert result.quality.local_monetary_complete
    ledger = next(group for group in result.quality.currencies if group.currency == "USD")
    assert ledger.known_amount_minor == (
        OutcomeAmounts(1234, 1234, 0, 0) if valid else OutcomeAmounts(1234, 0, 1234, 0)
    )
    if not valid:
        assert "INVALID_DATE" in {issue.code for issue in result.classified.rejected[0].issues}
        assert result.quality.review_required


def test_s08_mixed_currencies_and_explicit_precision(write_input, record, transport):
    rows = [
        record(transaction_id="1", currency="USD", amount="1.23"),
        record(transaction_id="2", currency="PLN", amount="2.34"),
        record(transaction_id="3", currency="JPY", amount="3"),
        record(transaction_id="4", currency="EUR", amount="4.56"),
        record(transaction_id="5", currency="JPY", amount="0.01"),
    ]
    path = write_input(rows, format=transport)
    result = analyze_transactions(path)
    assert [(s.currency, s.metrics.net_revenue) for s in result.aggregates.summary] == [
        ("JPY", "3"),
        ("PLN", "2.34"),
        ("USD", "1.23"),
    ]
    assert result.quality.unknown_currency_count == 1
    assert result.quality.unquantified_count == 2
    unknown = result.classified.rejected[0].row
    assert unknown.amount == Decimal("4.56") and unknown.amount_minor is None
    custom = analyze_transactions(path, config=CurrencyConfig({"EUR": 2, "JPY": 2}))
    assert custom.quality.counts.accepted == 5
    assert custom.aggregates.summary[0].currency == "EUR"
    assert custom.quality.unknown_currency_count == 0
    assert not custom.quality.review_required


@pytest.mark.parametrize("currency", ["usd", " USD", "US", "123", "EURO"])
def test_currency_codes_are_not_normalized(write_input, record, currency):
    result = analyze_transactions(write_input([record(currency=currency)]))
    assert result.quality.unknown_currency_count == 1
    assert "UNSUPPORTED_CURRENCY" in {i.code for i in result.classified.rejected[0].issues}


def test_s09_duplicates_normalize_money_ignore_provenance_and_keep_first(
    write_input, record, transport
):
    headers = [*BUSINESS_COLUMNS, "source_file", "source_sheet", "source_row"]
    rows = [
        record(amount="1.2300", source_file="z.csv", source_row="90"),
        record(amount="1.23", source_file="a.csv", source_row="2"),
        record(amount="+1.230", source_file="a.csv", source_row="3"),
        record(branch_id="002", amount="1.23", source_file="a.csv", source_row="4"),
    ]
    result = analyze_transactions(write_input(rows, headers=headers, format=transport))
    assert result.quality.counts == OutcomeCounts(4, 2, 0, 2)
    winner = result.classified.accepted[0]
    assert winner.row.raw.location.import_row == 2
    assert winner.row.provenance.source_file == "z.csv"
    assert all(item.duplicate_of == winner.row.raw.location for item in result.classified.redundant)
    assert not result.quality.review_required
    assert result.quality.provenance.complete == 4


@pytest.mark.parametrize(
    "change,expected_code",
    [
        ({"amount": "99.00"}, "CONFLICTING_TRANSACTION"),
        ({"customer_id": "other"}, "CONFLICTING_TRANSACTION"),
        ({"transaction_date": "2024-02-01"}, "CONFLICTING_TRANSACTION"),
        ({"currency": "PLN"}, "CONFLICTING_TRANSACTION"),
        ({"transaction_date": "bad"}, "KEY_HAS_INVALID_CANDIDATE"),
        ({"amount": "invalid"}, "KEY_HAS_INVALID_CANDIDATE"),
    ],
)
def test_s09_entire_readable_key_is_quarantined(
    write_input, record, transport, change, expected_code
):
    rows = [record(), record(**change), record(), record(branch_id="002")]
    result = analyze_transactions(write_input(rows, format=transport))
    assert result.quality.counts == OutcomeCounts(4, 1, 3, 0)
    assert all(
        expected_code in {i.code for i in item.issues} for item in result.classified.rejected
    )
    assert result.classified.accepted[0].row.branch_id == "002"


def test_single_invalid_candidate_and_unreadable_keys(write_input, record):
    result = analyze_transactions(
        write_input(
            [
                record(transaction_id="invalid", transaction_date="bad"),
                record(branch_id="", transaction_id="", customer_id="", transaction_date="bad"),
                record(),
            ]
        )
    )
    first, missing = result.classified.rejected
    assert {i.code for i in first.issues} == {"INVALID_DATE", "KEY_HAS_INVALID_CANDIDATE"}
    assert len(missing.issues) == 4
    assert missing.row.business_key is None
    assert result.quality.counts == OutcomeCounts(3, 1, 2, 0)
    assert result.quality.local_monetary_complete


def test_s10_blank_partial_and_multiline_physical_rows(write_input, record):
    path = write_input(
        [
            record(customer_id="customer\r\nnext"),
            [""] * 6,
            record(transaction_id="two", customer_id=""),
        ]
    )
    with path.open("a", encoding="utf-8", newline="") as stream:
        stream.write("\n")
    result = analyze_transactions(path)
    assert [item.row.raw.location.import_row for item in result.classified.rows] == [2, 4, 5, 6]
    assert result.classified.accepted[0].row.customer_id == "customer\r\nnext"
    assert result.quality.counts == OutcomeCounts(4, 1, 3, 0)
    assert result.quality.unquantified_count == 2
    assert sum("BLANK_ROW" in {i.code for i in row.issues} for row in result.classified.rows) == 2


def test_s12_header_only_does_not_invent_currency_metrics(write_input, transport):
    result = analyze_transactions(write_input([], format=transport))
    assert result.quality.counts == OutcomeCounts(0, 0, 0, 0)
    assert result.quality.local_monetary_complete
    assert result.quality.review_reasons == ("NO_ACCEPTED_RECORDS",)
    assert (
        result.aggregates.summary == result.aggregates.monthly == result.aggregates.branches == ()
    )
    assert result.upstream is None


def test_s15_independent_count_and_known_money_ledgers(write_input, record, transport):
    rows = [
        record(amount="10.00"),
        record(amount="10.00"),
        record(transaction_id="bad-date", transaction_date="invalid", amount="3.00"),
        record(transaction_id="bad-money", amount="NaN"),
        record(transaction_id="unknown", amount="5.00", currency="EUR"),
        [""] * 6,
    ]
    result = analyze_transactions(write_input(rows, format=transport))
    quality = result.quality
    assert quality.counts == OutcomeCounts(6, 1, 4, 1)
    assert quality.unknown_currency_count == 2
    assert quality.unquantified_count == 3
    usd = next(c for c in quality.currencies if c.currency == "USD")
    assert usd.counts == OutcomeCounts(4, 1, 2, 1)
    assert usd.known_amount_minor == OutcomeAmounts(2300, 1000, 300, 1000)
    assert usd.unquantified_counts == OutcomeCounts(1, 0, 1, 0)
    assert quality.count_balanced and quality.known_amounts_balanced
    assert quality.review_reasons == ("LOCAL_REJECTIONS", "LOCAL_UNQUANTIFIED_AMOUNTS")
    assert result.aggregates.summary[0].metrics.net_revenue == "10.00"


def test_distinct_customers_not_added_across_months_and_lexical_order(write_input, record):
    result = analyze_transactions(
        write_input(
            [
                record(branch_id="2", transaction_id="z", transaction_date="2024-03-01"),
                record(branch_id="10", transaction_id="b", transaction_date="2024-01-01"),
                record(branch_id="10", transaction_id="a", transaction_date="2024-03-01"),
            ]
        )
    )
    assert result.aggregates.summary[0].metrics.unique_customer_count == 2
    assert [m.month for m in result.aggregates.monthly] == ["2024-01", "2024-03"]
    assert [b.branch_id for b in result.aggregates.branches] == ["10", "2"]
    assert sum(m.metrics.unique_customer_count for m in result.aggregates.monthly) == 3


def test_classification_and_reconciliation_detect_corrupted_results(write_input, record):
    validated = validate_transactions(read_transactions(write_input([record(), record()])))
    # Physical location, not passed tuple order or provenance, determines winner.
    classified = classify_transactions(replace(validated, rows=tuple(reversed(validated.rows))))
    assert classified.accepted[0].row.raw.location.import_row == 2
    assert reconcile_transactions(classified).counts == OutcomeCounts(2, 1, 0, 1)
    corruptions = [
        replace(classified, rows=classified.rows[:-1]),
        replace(classified, rows=(*classified.rows, classified.rows[0])),
        replace(
            classified, rows=(replace(classified.rows[0], status="rejected"), classified.rows[1])
        ),
        replace(
            classified, rows=(classified.rows[0], replace(classified.rows[1], duplicate_of=None))
        ),
    ]
    for broken in corruptions:
        with pytest.raises(IntegrityError):
            reconcile_transactions(broken)
    altered = replace(validated.rows[0], amount_minor=999)
    altered_batch = replace(validated, rows=(altered, validated.rows[1]))
    with pytest.raises(IntegrityError, match="minor-unit"):
        reconcile_transactions(classify_transactions(altered_batch))


def test_metrics_are_stable_across_repeated_runs_and_permutation(write_input, record):
    rows = [record(transaction_id="3"), record(transaction_id="1"), record(transaction_id="2")]
    path = write_input(rows)
    first = analyze_transactions(path)
    assert first == analyze_transactions(path)
    reordered = analyze_transactions(write_input(list(reversed(rows))))
    assert first.aggregates == reordered.aggregates
    assert first.quality == reordered.quality


def test_conflicting_and_invalid_candidate_reasons_can_coexist(write_input, record, transport):
    result = analyze_transactions(
        write_input(
            [
                record(),
                record(amount="13.00"),
                record(transaction_date="invalid"),
            ],
            format=transport,
        )
    )
    assert result.quality.counts == OutcomeCounts(3, 0, 3, 0)
    for item in result.classified.rows:
        assert {"CONFLICTING_TRANSACTION", "KEY_HAS_INVALID_CANDIDATE"} <= {
            issue.code for issue in item.issues
        }
    assert result.quality.local_monetary_complete
    assert result.quality.review_reasons == ("NO_ACCEPTED_RECORDS", "LOCAL_REJECTIONS")


def test_rollup_partitions_are_exact_for_each_currency(write_input, record, transport):
    rows = [
        record(
            branch_id=branch,
            transaction_id=str(index),
            currency=currency,
            amount=amount,
            transaction_date=f"2024-{month:02}-01",
        )
        for index, (branch, currency, month, amount) in enumerate(
            [
                ("A", "USD", 1, "0.10"),
                ("A", "USD", 2, "-1.23"),
                ("B", "USD", 2, "0.00"),
                ("A", "JPY", 1, "123"),
                ("B", "JPY", 3, "-10"),
                ("A", "PLN", 1, "3.00"),
            ]
        )
    ]
    result = analyze_transactions(write_input(rows, format=transport))
    for summary in result.aggregates.summary:
        for grouped in (result.aggregates.monthly, result.aggregates.branches):
            groups = [group for group in grouped if group.currency == summary.currency]
            assert (
                sum(g.metrics.net_revenue_minor for g in groups)
                == summary.metrics.net_revenue_minor
            )
            assert (
                sum(g.metrics.transaction_count for g in groups)
                == summary.metrics.transaction_count
            )
