# -*- coding: utf-8 -*-

"""
unit tests for eparse cli
"""

import json

import pytest
from click.testing import CliRunner

from eparse.cli import main

kwargs = {"obj": {}, "catch_exceptions": False}
MULTI_SHEET_FIXTURE = "tests/files/multi_sheet_tables.xlsx"


def test_main():
    runner = CliRunner()
    result = runner.invoke(main, ["--help"])
    assert result.exit_code == 0
    assert "Usage" in result.output


def test_scan():
    runner = CliRunner()
    result = runner.invoke(main, ["-f", "tests/", "scan"], **kwargs)
    assert result.exit_code == 0
    assert "eparse_unit_test_data" in result.output


def test_parse():
    runner = CliRunner()
    result = runner.invoke(
        main, ["-f", "tests/", "parse", "--exclude-nested"], **kwargs
    )
    assert result.exit_code == 0
    assert "eparse_unit_test_data" in result.output


def test_query():
    runner = CliRunner()
    result = runner.invoke(main, ["-i", "sqlite3:///tests/test.db", "query"], **kwargs)
    assert result.exit_code == 0
    assert result.output == ""


def test_migrate():
    runner = CliRunner()
    result = runner.invoke(
        main,
        ["-i", "sqlite3:///tests/test.db", "migrate", "-m", "migration_000102_000200"],
        **kwargs,
    )
    assert result.exit_code == 1
    assert "duplicate column name: timestamp" in result.output


def test_parse_invalid_filter_pattern_exit_code():
    runner = CliRunner()
    result = runner.invoke(
        main,
        ["-f", MULTI_SHEET_FIXTURE, "parse", "--sheet-filter", "["],
        **kwargs,
    )
    assert result.exit_code != 0
    assert "Traceback" not in result.output
    assert "filter error" in result.output

    result = runner.invoke(
        main,
        ["-f", MULTI_SHEET_FIXTURE, "parse", "--range-filter", "("],
        **kwargs,
    )
    assert result.exit_code != 0
    assert "Traceback" not in result.output
    assert "filter error" in result.output


def test_parse_json_export(tmp_path):
    out = tmp_path / "out.json"
    runner = CliRunner()
    result = runner.invoke(
        main,
        [
            "-f",
            MULTI_SHEET_FIXTURE,
            "-o",
            str(out),
            "parse",
            "--sheet-filter",
            "Sheet[12]",
            "--range-filter",
            "A1:.*",
            "--type",
            "json",
        ],
        **kwargs,
    )
    assert result.exit_code == 0

    records = json.loads(out.read_text())
    sheets = {r["sheet"] for r in records}
    table_names = {r["name"] for r in records}
    assert sheets == {"Sheet1", "Sheet2"}
    assert table_names == {"ID", "Key"}
    # 2 tables x 3 rows x 2 cols of cell-level records
    assert len(records) == 12


def test_parse_csv_export(tmp_path):
    out = tmp_path / "out.csv"
    runner = CliRunner()
    result = runner.invoke(
        main,
        [
            "-f",
            MULTI_SHEET_FIXTURE,
            "-o",
            str(out),
            "parse",
            "--sheet-filter",
            "Sheet3",
            "--type",
            "csv",
        ],
        **kwargs,
    )
    assert result.exit_code == 0

    lines = out.read_text().splitlines()
    header = lines[0].split(",")
    assert "sheet" in header
    assert "excel_RC" in header
    assert all("Sheet1" not in line and "Sheet2" not in line for line in lines[1:])
    assert any("Sheet3" in line for line in lines[1:])


def test_outputs():
    runner = CliRunner()
    result = runner.invoke(main, ["-o", "null:///", "scan"], **kwargs)
    assert result.exit_code == 0
    assert result.output == ""
    result = runner.invoke(main, ["-o", "stdout:///", "scan"], **kwargs)
    assert result.exit_code == 0
    assert result.output == ""
    result = runner.invoke(main, ["-o", "sqlite3:///:memory:", "scan"], **kwargs)
    assert result.exit_code == 0
    assert result.output == ""
    result = runner.invoke(main, ["-o", "test", "scan"], **kwargs)
    assert result.exit_code == 1
    assert "test is not a recognized endpoint" in result.output
    with pytest.raises(ValueError):
        result = runner.invoke(main, ["-d", "-o", "test", "scan"], **kwargs)
