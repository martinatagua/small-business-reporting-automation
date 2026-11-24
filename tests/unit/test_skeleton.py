from importlib.metadata import entry_points

import pytest

import small_business_reporting_automation as pkg
from small_business_reporting_automation import cli


def test_package_imports():
    assert pkg.__version__


def test_entry_point_declared():
    eps = [e for e in entry_points(group="console_scripts") if e.name == "sb-report"]
    assert len(eps) == 1
    assert eps[0].value == "small_business_reporting_automation.cli:main"
    assert eps[0].load() is cli.main


def test_main_reports_pending(capsys):
    assert cli.main([]) == 0
    assert "F4" in capsys.readouterr().out


def test_help_exits_zero(capsys):
    with pytest.raises(SystemExit) as exc:
        cli.main(["--help"])
    assert exc.value.code == 0
    assert "sb-report" in capsys.readouterr().out
