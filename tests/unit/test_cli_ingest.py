"""CLI ``sb-report ingest`` (demostración de F4) y modelos de salida de DT4-10."""

import dataclasses

import pytest

from small_business_reporting_automation import cli, models
from small_business_reporting_automation.models import (
    EnvironmentReport, FileOutcome, IngestedFile, IngestResult, ProcessingEvent, SourceRecord,
)

from pathlib import Path

REPO = Path(__file__).resolve().parents[2]


def test_cli_ingest_demo_data(capsys):
    assert cli.main(["ingest", "--config", str(REPO / "config/demo.toml"), "--input-dir", str(REPO / "data/raw")]) == 0
    out = capsys.readouterr().out
    assert "4 archivos aceptados, 0 rechazados, 177 registros de origen" in out


def test_cli_ingest_reports_rejections_and_keeps_going(tmp_path, capsys):
    (tmp_path / "centro_ventas_2026-11.csv").write_bytes(b'a;b\n"x"y;1\n')
    (tmp_path / "norte_ventas_2026-11.xlsx").write_bytes(b"no es un zip")
    assert cli.main(["ingest", "--config", str(REPO / "config/demo.toml"), "--input-dir", str(tmp_path)]) == 0
    out = capsys.readouterr().out
    assert "csv_malformed" in out and "zip_unreadable" in out and "2 rechazados" in out


def test_cli_ingest_missing_dir_is_controlled(tmp_path, capsys):
    assert cli.main(["ingest", "--config", str(REPO / "config/demo.toml"), "--input-dir", str(tmp_path / "no")]) == 2
    assert "no se puede leer" in capsys.readouterr().err


def test_cli_ingest_bad_config_is_controlled(tmp_path, capsys):
    bad = tmp_path / "c.toml"
    bad.write_text("config_version = 99\n")
    assert cli.main(["ingest", "--config", str(bad), "--input-dir", str(tmp_path)]) == 2


@pytest.mark.parametrize(
    "cls", [SourceRecord, IngestedFile, FileOutcome, ProcessingEvent, IngestResult, EnvironmentReport]
)
def test_output_models_are_frozen_slots(cls):
    assert cls.__dataclass_params__.frozen and "__slots__" in vars(cls)


def test_environment_report_fields_exact():
    assert [f.name for f in dataclasses.fields(EnvironmentReport)] == [
        "python_version", "openpyxl_version", "expat_version", "expat_version_info",
        "openpyxl_lxml", "openpyxl_defusedxml", "expat_min_ok", "expat_reparse_deferral",
    ]


def test_existing_issue_codes_untouched():
    assert len(models.IssueCode) == 26
