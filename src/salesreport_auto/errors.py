"""Public failures for configuration, structural inputs, money, and invariants."""


class ConfigurationError(ValueError):
    """A currency definition or input limit is invalid."""


class InputError(ValueError):
    """A structural file error prevents a trustworthy input result."""


class BundleError(InputError):
    """An invalid bundle or selected-master mismatch must never fall back to standalone."""

    def __init__(self, code: str, message: str) -> None:
        self.code = code
        super().__init__(f"{code}: {message}")


class IntegrityError(ValueError):
    """Records, classifications, or financial conservation are inconsistent."""


class ReportError(ValueError):
    """A workbook cannot be rendered, verified, or safely published."""


class MoneyError(ValueError):
    """A monetary token cannot be represented under the declared policy."""

    def __init__(self, code: str, message: str) -> None:
        self.code = code
        super().__init__(message)
