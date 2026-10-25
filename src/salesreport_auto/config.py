"""Explicit standalone currency precision and bounded local-file budgets."""

import re
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from types import MappingProxyType

from .errors import ConfigurationError

DEFAULT_PRECISION: Mapping[str, int] = MappingProxyType({"USD": 2, "PLN": 2, "JPY": 0})


@dataclass(frozen=True, init=False)
class CurrencyConfig:
    """Resolve defaults plus API overrides; iterable pairs detect duplicates."""

    precision: Mapping[str, int]
    origins: Mapping[str, str]

    def __init__(
        self, definitions: Mapping[str, int] | Iterable[tuple[str, int]] | None = None
    ) -> None:
        precision = dict(DEFAULT_PRECISION)
        origins = dict.fromkeys(precision, "default")
        entries = definitions.items() if isinstance(definitions, Mapping) else definitions
        seen = set()
        if entries is not None:
            try:
                for code, digits in entries:
                    if not isinstance(code, str) or not re.fullmatch(r"[A-Z]{3}", code):
                        raise ConfigurationError(
                            "Currency codes must be three uppercase ASCII letters"
                        )
                    if code in seen:
                        raise ConfigurationError(f"Repeated currency definition: {code}")
                    if type(digits) is not int or not 0 <= digits <= 9:
                        raise ConfigurationError(
                            f"{code}: precision must be an integer from 0 to 9"
                        )
                    seen.add(code)
                    precision[code] = digits
                    origins[code] = "user"
            except (TypeError, ValueError) as exc:
                if isinstance(exc, ConfigurationError):
                    raise
                raise ConfigurationError(
                    "Definitions must be (currency, integer precision) pairs"
                ) from exc
        object.__setattr__(self, "precision", MappingProxyType(precision))
        object.__setattr__(self, "origins", MappingProxyType(origins))

    @classmethod
    def from_bundle(cls, definitions: Mapping[str, int]) -> "CurrencyConfig":
        """Validate a complete bundle precision map without merging standalone defaults."""
        if not isinstance(definitions, Mapping) or not definitions:
            raise ConfigurationError("Bundle currency precision must be a nonempty object")
        validated = cls(definitions)
        result = object.__new__(cls)
        object.__setattr__(
            result,
            "precision",
            MappingProxyType({code: validated.precision[code] for code in definitions}),
        )
        object.__setattr__(
            result, "origins", MappingProxyType(dict.fromkeys(definitions, "bundle"))
        )
        return result


@dataclass(frozen=True)
class InputLimits:
    """MVP limits; callers may provide stricter budgets, never larger ones."""

    max_file_bytes: int = 20 * 1024 * 1024
    max_records: int = 100_000
    max_xlsx_uncompressed_bytes: int = 200 * 1024 * 1024

    def __post_init__(self) -> None:
        for name, maximum in (
            ("max_file_bytes", 20 * 1024 * 1024),
            ("max_records", 100_000),
            ("max_xlsx_uncompressed_bytes", 200 * 1024 * 1024),
        ):
            value = getattr(self, name)
            if type(value) is not int or not 0 < value <= maximum:
                raise ConfigurationError(f"{name} must be a positive integer at most {maximum}")
