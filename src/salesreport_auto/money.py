"""Context-independent money parsing, integer conversion, and final rounding."""

import re
from dataclasses import dataclass
from decimal import Decimal

from .errors import MoneyError

MAX_TOKEN_LENGTH = 512
MAX_MINOR_DIGITS = 38


def _check_precision(precision: int) -> None:
    if type(precision) is not int or not 0 <= precision <= 9:
        raise ValueError("Precision must be an integer from 0 to 9")


def parse_decimal(text: str) -> Decimal:
    """Parse finite plain decimal text, without whitespace or exponent syntax."""
    if (
        not isinstance(text, str)
        or len(text) > MAX_TOKEN_LENGTH
        or re.fullmatch(r"[+-]?[0-9]+(?:\.[0-9]+)?", text) is None
    ):
        raise MoneyError("INVALID_AMOUNT", "Expected at most 512 characters of plain decimal text")
    return Decimal(text)


def to_minor_units(amount: Decimal, precision: int) -> int:
    """Convert a finite Decimal exactly; reject rounding and row-size overflow."""
    _check_precision(precision)
    if not isinstance(amount, Decimal) or not amount.is_finite():
        raise MoneyError("INVALID_AMOUNT", "Expected a finite Decimal")
    if amount.is_zero():
        return 0
    sign, digits, exponent = amount.as_tuple()
    shift = exponent + precision
    if shift < 0:
        remove = -shift
        if remove >= len(digits) or any(digits[-remove:]):
            raise MoneyError(
                "EXCESS_FRACTIONAL_PRECISION", "Amount requires fractional minor units"
            )
        digits = digits[:-remove]
        shift = 0
    if len(digits) + shift > MAX_MINOR_DIGITS:
        raise MoneyError("AMOUNT_TOO_LARGE", "Nonzero minor-unit amount exceeds 38 digits")
    coefficient = int("".join(map(str, digits)))
    return (-1 if sign else 1) * coefficient * 10**shift


def parse_money(text: str, precision: int) -> int:
    """Parse a monetary token and return exact Python integer minor units."""
    return to_minor_units(parse_decimal(text), precision)


def format_minor(amount: int, precision: int) -> str:
    """Format an arbitrarily large signed integer without Decimal/float math."""
    _check_precision(precision)
    if type(amount) is not int:
        raise TypeError("Minor units must be a Python integer")
    sign = "-" if amount < 0 else ""
    digits = str(abs(amount)).zfill(precision + 1)
    return sign + digits if precision == 0 else f"{sign}{digits[:-precision]}.{digits[-precision:]}"


@dataclass(frozen=True)
class MonetaryAverage:
    """Exact minor-unit ratio with one HALF_UP rounding at display time."""

    numerator_minor: int
    denominator: int
    precision: int

    def __post_init__(self) -> None:
        _check_precision(self.precision)
        if type(self.numerator_minor) is not int:
            raise TypeError("Average numerator must be a Python integer")
        if type(self.denominator) is not int or self.denominator < 0:
            raise ValueError("Average denominator must be a nonnegative integer")

    @property
    def rounded_minor(self) -> int | None:
        """Return HALF_UP minor units; no denominator means no average."""
        if self.denominator == 0:
            return None
        quotient, remainder = divmod(abs(self.numerator_minor), self.denominator)
        quotient += 2 * remainder >= self.denominator
        return -quotient if self.numerator_minor < 0 else quotient

    @property
    def display(self) -> str:
        """Return a fixed-precision display string, or N/A."""
        rounded = self.rounded_minor
        return "N/A" if rounded is None else format_minor(rounded, self.precision)
