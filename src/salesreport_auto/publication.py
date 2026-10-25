"""Verified, no-clobber report publication using a same-directory staging file."""

import hashlib
import os
import tempfile
from dataclasses import asdict, dataclass, replace
from pathlib import Path

from .bundle_io import verify_bundle_unchanged
from .config import CurrencyConfig, InputLimits
from .engine import TransactionAnalysis, analyze_products, analyze_transactions
from .errors import ReportError
from .models import ReportMode
from .report_model import WorkbookPlan, build_report_plan
from .workbook import verify_workbook, write_workbook


def _reject_links(path: Path) -> None:
    for component in (path, *path.parents):
        if component.is_symlink() or component.is_junction():
            raise ReportError(f"Path contains a symbolic link or junction: {path}")


def validate_output_path(input_path: str | Path, output: str | Path) -> Path:
    """Reject aliases, links, existing targets, and absent output parent folders."""
    source = Path(input_path).absolute()
    target = Path(output).absolute()
    if target.suffix.lower() != ".xlsx":
        raise ReportError("Output must be an .xlsx file path")
    _reject_links(source)
    _reject_links(target)
    if source.resolve() == target.resolve() or (
        source.exists() and target.exists() and source.samefile(target)
    ):
        raise ReportError("Output must not alias or overwrite the input file")
    if target.exists():
        raise ReportError(f"Output already exists; choose a new file: {target}")
    if not target.parent.is_dir():
        raise ReportError(f"Output parent directory must already exist: {target.parent}")
    return target


def _verify_input_unchanged(analysis: TransactionAnalysis) -> None:
    source = analysis.classified.validated.source
    _reject_links(source.path)
    # Recheck identity/size/timestamps as well as the original snapshot digest.
    with source.path.open("rb") as stream:
        before = os.fstat(stream.fileno())
        digest = hashlib.sha256()
        size = 0
        while chunk := stream.read(1024 * 1024):
            size += len(chunk)
            if size > 20 * 1024 * 1024:
                raise ReportError("Input changed or exceeds its original size budget")
            digest.update(chunk)
        after = os.fstat(stream.fileno())
    current = source.path.stat()
    signatures = {
        (entry.st_dev, entry.st_ino, entry.st_size, entry.st_mtime_ns)
        for entry in (before, after, current)
    }
    if len(signatures) != 1 or digest.hexdigest() != source.sha256:
        raise ReportError("Input changed after analysis; no report was published")


@dataclass(frozen=True)
class ReportResult:
    """Published status/results, plus any warning about an owned staging leftover."""

    output: Path
    analysis: TransactionAnalysis
    plan: WorkbookPlan
    cleanup_warning: str | None = None

    @property
    def status(self) -> str:
        """Return completed or review_required, including chart precision limits."""
        return self.plan.status

    @property
    def review_reasons(self) -> tuple[str, ...]:
        """Return ordered, unique report reasons without mutating core quality."""
        return self.plan.review_reasons

    def to_summary(self) -> dict[str, object]:
        """Return the specification's JSON-safe standalone success summary."""
        return {
            "status": self.status,
            "mode": self.analysis.mode,
            "output": str(self.output),
            "local_counts": asdict(self.analysis.quality.counts),
            "local_unquantified_count": self.analysis.quality.unquantified_count,
            "local_monetary_complete": self.analysis.quality.local_monetary_complete,
            "upstream": None
            if self.analysis.upstream is None
            else {
                "counts": asdict(self.analysis.upstream.reconciliation.counts),
                "unquantified_count": self.analysis.upstream.reconciliation.unquantified_count,
                "monetary_complete": self.analysis.upstream.reconciliation.monetary_complete,
            },
            "review_reasons": list(self.review_reasons),
            "currencies": [
                {
                    "code": group.currency,
                    "precision": group.metrics.precision,
                    "net_revenue": group.metrics.net_revenue,
                    "transaction_count": group.metrics.transaction_count,
                }
                for group in self.analysis.aggregates.summary
            ],
        }


def _publish_no_replace(staging: Path, target: Path) -> None:
    # Creating a hard link is atomic and fails if the destination exists, on both
    # Windows/NTFS and POSIX. Never fall back to a replacing rename or partial copy.
    os.link(staging, target)


def publish_report(analysis: TransactionAnalysis, output: str | Path) -> ReportResult:
    """Render, reopen/verify, and atomically publish without replacing any file.

    The filesystem must support same-directory hard links. Failure is closed: no
    rename/copy fallback that could overwrite an existing target is attempted.
    A cleanup denial after publication returns success with cleanup_warning;
    before publication it adds a note without hiding the original failure.
    """
    from .version import VERSION

    source = analysis.classified.validated.source.path
    target = validate_output_path(source, output)
    if analysis.upstream is not None and target.resolve().is_relative_to(
        analysis.upstream.directory.resolve()
    ):
        raise ReportError("Output must be outside the supplied bundle directory")
    staging: Path | None = None
    result: ReportResult | None = None
    failure: BaseException | None = None
    try:
        plan = build_report_plan(analysis, application_version=VERSION)
        if analysis.upstream is not None:
            verify_bundle_unchanged(analysis.upstream)
        _verify_input_unchanged(analysis)
        descriptor, name = tempfile.mkstemp(
            prefix=f".{target.stem}-", suffix=".xlsx", dir=target.parent
        )
        staging = Path(name)
        os.close(descriptor)
        write_workbook(plan, staging)
        verify_workbook(staging, plan)
        _verify_input_unchanged(analysis)
        validate_output_path(source, target)
        if analysis.upstream is not None:
            verify_bundle_unchanged(analysis.upstream)
        _publish_no_replace(staging, target)
        result = ReportResult(target, analysis, plan)
    except ReportError as exc:
        failure = exc
        raise
    except Exception as exc:
        # Errors contain no raw customer records; callers get a stable failure type.
        failure = ReportError(f"Cannot publish report ({type(exc).__name__})")
        raise failure from exc
    except BaseException as exc:
        failure = exc
        raise
    finally:
        if staging is not None:
            try:
                staging.unlink(missing_ok=True)
            except OSError as exc:
                warning = f"Cannot remove owned temporary file {staging} ({type(exc).__name__})"
                if result is not None:
                    # Publication has committed; reporting failure here would lie
                    # about the complete report already at the final path.
                    result = replace(result, cleanup_warning=warning)
                elif failure is not None:
                    failure.add_note(warning)
    assert result is not None  # A normal return is possible only after publication committed.
    return result


def generate_report(
    input_path: str | Path,
    output: str | Path,
    *,
    mode: ReportMode = "transactions",
    bundle: str | Path | None = None,
    config: CurrencyConfig | None = None,
    limits: InputLimits | None = None,
) -> ReportResult:
    """Analyze the explicit standalone input mode and safely publish one Excel report."""
    if mode not in ("transactions", "products"):
        raise ReportError("Mode must be transactions or products")
    if bundle is not None and mode != "transactions":
        raise ReportError("--bundle is supported only in transaction mode")
    if bundle is not None and Path(output).absolute().resolve().is_relative_to(
        Path(bundle).absolute().resolve()
    ):
        raise ReportError("Output must be outside the supplied bundle directory")
    validate_output_path(input_path, output)
    analyze = analyze_products if mode == "products" else analyze_transactions
    analysis = analyze(
        input_path,
        config=config,
        limits=limits,
        **({"bundle": bundle} if bundle is not None else {}),
    )
    return publish_report(analysis, output)
