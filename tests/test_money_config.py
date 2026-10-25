from decimal import Decimal, localcontext

import pytest

from salesreport_auto import (
    ConfigurationError,
    CurrencyConfig,
    InputLimits,
    MonetaryAverage,
    MoneyError,
    format_minor,
    parse_decimal,
    parse_money,
    to_minor_units,
)


def test_explicit_config_is_immutable_and_tracks_origins():
    definitions = {"EUR": 2, "USD": 3}
    config = CurrencyConfig(definitions)
    definitions["EUR"] = 9
    assert dict(config.precision) == {"USD": 3, "PLN": 2, "JPY": 0, "EUR": 2}
    assert config.origins["USD"] == "user"
    assert config.origins["PLN"] == "default"
    with pytest.raises(TypeError):
        config.precision["USD"] = 4


@pytest.mark.parametrize(
    "definitions",
    [
        [("USD", 2), ("USD", 2)],
        {"usd": 2},
        {"EU": 2},
        {"123": 2},
        {"EUR": True},
        {"EUR": -1},
        {"EUR": 10},
        {"EUR": "2"},
        ["EUR=2"],
        3,
    ],
)
def test_bad_precision_definitions(definitions):
    with pytest.raises(ConfigurationError):
        CurrencyConfig(definitions)


@pytest.mark.parametrize(
    "limits",
    [
        {"max_records": 0},
        {"max_records": True},
        {"max_records": 100_001},
        {"max_file_bytes": 21 * 1024 * 1024},
        {"max_xlsx_uncompressed_bytes": -1},
    ],
)
def test_limits_must_be_bounded_positive_integers(limits):
    with pytest.raises(ConfigurationError):
        InputLimits(**limits)


@pytest.mark.parametrize(
    "text,precision,expected",
    [
        ("0.10", 2, 10),
        ("-0.01", 2, -1),
        ("+12.3400", 2, 1234),
        ("-0.00", 2, 0),
        ("12", 0, 12),
        ("1.234", 3, 1234),
        ("0.000000001", 9, 1),
        ("00001.00", 2, 100),
        ("9" * 38, 0, int("9" * 38)),
    ],
)
def test_money_is_exact_independent_of_context(text, precision, expected):
    with localcontext() as context:
        context.prec = 2
        assert parse_money(text, precision) == expected
    assert type(expected) is int


@pytest.mark.parametrize(
    "text,code",
    [
        ("1.231", "EXCESS_FRACTIONAL_PRECISION"),
        ("0.001", "EXCESS_FRACTIONAL_PRECISION"),
        ("NaN", "INVALID_AMOUNT"),
        ("Infinity", "INVALID_AMOUNT"),
        ("1e2", "INVALID_AMOUNT"),
        ("1,23", "INVALID_AMOUNT"),
        ("12.34 ", "INVALID_AMOUNT"),
        (" 12.34", "INVALID_AMOUNT"),
        ("", "INVALID_AMOUNT"),
        (".1", "INVALID_AMOUNT"),
        ("1.", "INVALID_AMOUNT"),
        ("(1.00)", "INVALID_AMOUNT"),
        ("$1.00", "INVALID_AMOUNT"),
        ("١.٢", "INVALID_AMOUNT"),
        ("9" * 39, "AMOUNT_TOO_LARGE"),
        ("0" * 513, "INVALID_AMOUNT"),
    ],
)
def test_invalid_money_is_never_rounded(text, code):
    precision = 0 if len(text) == 39 else 2
    with pytest.raises(MoneyError) as error:
        parse_money(text, precision)
    assert error.value.code == code


def test_token_boundary_zeros_and_decimal_contract():
    assert parse_money("0" * 512, 2) == 0
    assert parse_decimal("1.2300") == Decimal("1.2300")
    for bad in (Decimal("NaN"), Decimal("Infinity"), 0.1):
        with pytest.raises(MoneyError):
            to_minor_units(bad, 2)
    for precision in (True, -1, 10):
        with pytest.raises(ValueError):
            parse_money("1.00", precision)


@pytest.mark.parametrize(
    "numerator,denominator,precision,display",
    [
        (1, 2, 2, "0.01"),
        (-1, 2, 2, "-0.01"),
        (800, 3, 2, "2.67"),
        (1101, 4, 2, "2.75"),
        (0, 1, 0, "0"),
        (1, 2, 0, "1"),
        (123, 0, 2, "N/A"),
        (-1, 3, 2, "0.00"),
    ],
)
def test_final_half_up_rounding_preserves_exact_ratio(numerator, denominator, precision, display):
    with localcontext() as context:
        context.prec = 2
        average = MonetaryAverage(numerator, denominator, precision)
        assert average.display == display
        assert (average.numerator_minor, average.denominator) == (numerator, denominator)
    assert format_minor(10**45, 2) == "1" + "0" * 43 + ".00"


def test_average_and_formatter_reject_invalid_core_arguments():
    with pytest.raises(ValueError):
        MonetaryAverage(1, -1, 2)
    with pytest.raises(TypeError):
        MonetaryAverage(0.1, 1, 2)
    with pytest.raises(TypeError):
        format_minor(0.1, 2)
