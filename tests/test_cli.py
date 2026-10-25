import json
import os
import subprocess
import sys
import tomllib
from pathlib import Path

import pytest

from salesreport_auto import ReportError, cli
from salesreport_auto.cli import main


def _arguments(source, output, *extra):
    return ["--mode", "transactions", "--input", str(source), "--output", str(output), *extra]


def test_cli_clean_summary_schema_and_exact_values(write_input, record, tmp_path, capsys):
    source = write_input([record(), record()])
    output = tmp_path / "report.xlsx"
    assert main(_arguments(source, output)) == 0
    captured = capsys.readouterr()
    assert captured.err == ""
    result = json.loads(captured.out)
    assert set(result) == {
        "status",
        "mode",
        "output",
        "local_counts",
        "local_unquantified_count",
        "local_monetary_complete",
        "upstream",
        "review_reasons",
        "currencies",
    }
    assert result["status"] == "completed"
    assert result["local_counts"] == {"input": 2, "accepted": 1, "rejected": 0, "redundant": 1}
    assert result["currencies"] == [
        {"code": "USD", "precision": 2, "net_revenue": "12.34", "transaction_count": 1}
    ]
    assert result["upstream"] is None and result["review_reasons"] == []
    assert output.is_file()


@pytest.mark.parametrize("case", ["invalid", "conflict", "empty", "provenance", "chart_limit"])
def test_cli_review_reports_publish_with_exit_one(write_input, record, tmp_path, capsys, case):
    headers = None
    if case == "invalid":
        rows = [record(amount="invalid")]
    elif case == "conflict":
        rows = [record(), record(amount="99.00")]
    elif case == "empty":
        rows = []
    elif case == "provenance":
        rows = [record(source_row="2")]
        headers = [*rows[0]]
    else:
        rows = [record(amount="9999999999999999", currency="JPY")]
    source = write_input(rows, headers=headers)
    output = tmp_path / "review.xlsx"
    assert main(_arguments(source, output)) == 1
    captured = capsys.readouterr()
    summary = json.loads(captured.out)
    assert captured.err == ""
    assert summary["status"] == "review_required" and summary["review_reasons"]
    assert output.is_file()
    if case == "chart_limit":
        assert summary["review_reasons"] == ["CHART_PRECISION_LIMIT"]


def test_cli_precision_declarations_extend_and_override_defaults(
    write_input, record, tmp_path, capsys
):
    source = write_input(
        [
            record(currency="EUR", transaction_id="1", amount="1.234"),
            record(currency="JPY", transaction_id="2", amount="0.01"),
        ]
    )
    assert (
        main(
            _arguments(
                source,
                tmp_path / "custom.xlsx",
                "--currency-precision",
                "EUR=3",
                "--currency-precision",
                "JPY=2",
            )
        )
        == 0
    )
    summary = json.loads(capsys.readouterr().out)
    assert summary["currencies"] == [
        {"code": "EUR", "precision": 3, "net_revenue": "1.234", "transaction_count": 1},
        {"code": "JPY", "precision": 2, "net_revenue": "0.01", "transaction_count": 1},
    ]


@pytest.mark.parametrize(
    "definitions", [["EUR=2", "EUR=2"], ["eur=2"], ["EUR=10"], ["EUR=x"], ["EUR=-1"]]
)
def test_cli_bad_precision_is_fatal_without_report(
    write_input, record, tmp_path, capsys, definitions
):
    source = write_input([record()])
    args = []
    for value in definitions:
        args.extend(["--currency-precision", value])
    output = tmp_path / "bad.xlsx"
    assert main(_arguments(source, output, *args)) == 2
    captured = capsys.readouterr()
    assert captured.out == ""
    assert "salesreport-auto: error:" in captured.err
    assert not output.exists()


@pytest.mark.parametrize("option", ["--help", "--version"])
def test_help_and_version_need_no_report_arguments(capsys, option):
    with pytest.raises(SystemExit) as exited:
        main([option])
    assert exited.value.code == 0
    captured = capsys.readouterr()
    assert captured.out and captured.err == ""
    if option == "--help":
        assert "--bundle verifies a SheetMerge" in " ".join(captured.out.split())
        assert "USD=2" in captured.out
    else:
        assert captured.out.strip() == "SalesReport Auto 0.1.0"


@pytest.mark.parametrize("extra", [["--unknown"], ["--out", "other.xlsx"]])
def test_unsupported_options_fail_clearly(write_input, record, tmp_path, capsys, extra):
    source = write_input([record()])
    output = tmp_path / "unsupported.xlsx"
    with pytest.raises(SystemExit) as exited:
        main(_arguments(source, output, *extra))
    assert exited.value.code == 2
    captured = capsys.readouterr()
    assert captured.out == "" and "salesreport-auto: error:" in captured.err
    assert not output.exists()


def test_unknown_mode_and_missing_arguments_fail(write_input, record, tmp_path, capsys):
    source = write_input([record()])
    arguments = _arguments(source, tmp_path / "products.xlsx")
    arguments[1] = "unsupported"
    for args in (arguments, []):
        with pytest.raises(SystemExit) as exited:
            main(args)
        assert exited.value.code == 2
        assert "salesreport-auto: error:" in capsys.readouterr().err


def test_structural_error_and_collision_have_no_success_summary(
    write_input, record, tmp_path, capsys
):
    source = tmp_path / "malformed.csv"
    source.write_text("bad headers", encoding="utf-8")
    output = tmp_path / "failed.xlsx"
    assert main(_arguments(source, output)) == 2
    captured = capsys.readouterr()
    assert not output.exists() and captured.out == "" and captured.err
    source = write_input([record()])
    output.write_bytes(b"existing")
    assert main(_arguments(source, output)) == 2
    captured = capsys.readouterr()
    assert captured.out == "" and "already exists" in captured.err
    assert output.read_bytes() == b"existing"


def test_python_module_entry_point_and_declared_console_script(write_input, record, tmp_path):
    source = write_input([record()])
    project = Path(__file__).parents[1]
    environment = os.environ.copy()
    environment["PYTHONPATH"] = str(project / "src")
    output = tmp_path / "module.xlsx"
    run = subprocess.run(
        [sys.executable, "-m", "salesreport_auto", *_arguments(source, output)],
        env=environment,
        capture_output=True,
        text=True,
        check=False,
    )
    assert run.returncode == 0 and run.stderr == ""
    assert json.loads(run.stdout)["status"] == "completed"
    assert output.is_file()
    config = tomllib.loads((project / "pyproject.toml").read_text(encoding="utf-8"))
    assert config["project"]["scripts"] == {"salesreport-auto": "salesreport_auto.cli:main"}


def test_cli_writer_failure_is_stderr_only(write_input, record, tmp_path, capsys, monkeypatch):
    source = write_input([record()])

    def fail(*args, **kwargs):
        raise ReportError("Synthetic report failure")

    monkeypatch.setattr(cli, "generate_report", fail)
    output = tmp_path / "failed-writer.xlsx"
    assert main(_arguments(source, output)) == 2
    captured = capsys.readouterr()
    assert captured.out == ""
    assert captured.err.strip() == "salesreport-auto: error: Synthetic report failure"
    assert not output.exists()
