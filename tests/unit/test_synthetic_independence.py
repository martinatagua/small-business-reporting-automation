"""T3-01, T3-02, T3-03, T3-13, T3-14: independencia del generador y contraste con F2."""

import ast
import csv
import fnmatch
import re
import sys
import tomllib
from decimal import Decimal
from pathlib import Path

import small_business_reporting_automation.synthetic as synthetic
from small_business_reporting_automation.config import parse_config
from small_business_reporting_automation.models import IssueCode, canonical_text

REPO = Path(__file__).resolve().parents[2]
SRC = Path(synthetic.__file__)
TREE = ast.parse(SRC.read_text(encoding="utf-8"))
ALLOWED = {"random", "hashlib", "json", "csv", "io", "datetime", "decimal", "dataclasses", "pathlib", "os", "openpyxl"}
FIXTURE = REPO / "tests" / "fixtures" / "manual_oracle"


def test_t3_01_generator_independence():
    imported = []
    for node in ast.walk(TREE):
        if isinstance(node, ast.Import):
            imported += [a.name.split(".")[0] for a in node.names]
        elif isinstance(node, ast.ImportFrom):
            assert node.level == 0, "import relativo prohibido"
            imported.append((node.module or "").split(".")[0])
    assert set(imported) <= ALLOWED, set(imported) - ALLOWED
    for name in imported:
        assert name == "openpyxl" or name in sys.stdlib_module_names
    forbidden = {"normalize", "rules", "pipeline", "export", "config", "models", "pandas", "importlib",
                 "small_business_reporting_automation", "ingest"}
    assert not forbidden & set(imported)
    for node in ast.walk(TREE):
        if isinstance(node, ast.Call):
            func = node.func
            assert not (isinstance(func, ast.Name) and func.id in {"__import__", "exec", "eval", "compile"})


def test_t3_02_rng_is_local():
    imports = {a.name for n in ast.walk(TREE) if isinstance(n, ast.Import) for a in n.names}
    assert "time" not in imports and "uuid" not in imports and "secrets" not in imports
    for node in ast.walk(TREE):
        if isinstance(node, ast.Attribute):
            if isinstance(node.value, ast.Name) and node.value.id == "random":
                assert node.attr == "Random", f"RNG global: random.{node.attr}"
            assert node.attr not in {"now", "today", "utcnow", "urandom", "uuid4", "perf_counter"}
            assert not (isinstance(node.value, ast.Name) and node.value.id in {"time", "uuid", "secrets"})
    # el único uso de ``random.Random`` está en la derivación por archivo
    calls = [n for n in ast.walk(TREE) if isinstance(n, ast.Attribute) and n.attr == "Random"]
    assert len(calls) >= 1
    # sin iteración directa sobre ``set``
    for node in ast.walk(TREE):
        if isinstance(node, (ast.For, ast.comprehension)):
            it = node.iter
            assert not (isinstance(it, ast.Call) and getattr(it.func, "id", "") == "set")
            assert not isinstance(it, ast.Set)


def test_t3_03_issue_code_literals():
    codes = {c.value for c in IssueCode}
    assert len(codes) == 26
    pattern = re.compile(r"^[CFR]\d{2}_[A-Z_]+$")
    literals = {n.value for n in ast.walk(TREE) if isinstance(n, ast.Constant) and isinstance(n.value, str)
                and pattern.match(n.value)}
    assert literals <= codes
    assert set(synthetic.ISSUE_CODES) == codes and len(synthetic.ISSUE_CODES) == 26


# --------------------------------------------------------------------------------------------------
# T3-13: fixture del oráculo manual
# --------------------------------------------------------------------------------------------------


def test_t3_13_manual_oracle_fixture_complete():
    expected = list(csv.DictReader((FIXTURE / "expected.csv").read_text(encoding="utf-8").splitlines()))
    codes = set()
    for row in expected:
        codes |= {c for c in row["expected_issue_codes"].split(";") if c}
    assert {f"R{n:02d}" for n in range(1, 15)} <= {c[:3] for c in codes}
    assert "F11_UNMAPPED_COLUMN" in codes
    valid_codes = {c.value for c in IssueCode}
    assert codes <= valid_codes
    # el perfil propio es válido, mapea reported_total y branch_in_file y declara dos formatos de fecha
    cfg = parse_config((FIXTURE / "config.toml").read_text(encoding="utf-8"))
    profile = cfg.sources[0]
    assert "reported_total" in profile.columns and "branch_in_file" in profile.columns
    assert len(profile.date_formats) == 2
    # la tabla describe exactamente los registros de sales.csv (cabecera = registro 1)
    sales = list(csv.reader((FIXTURE / "sales.csv").read_text(encoding="utf-8").splitlines()))
    record_rows = [int(r["source_row"]) for r in expected if r["level"] == "record"]
    assert record_rows == list(range(2, len(sales) + 1))
    assert len(record_rows) >= 10
    # escrito a mano: no lo produce el generador ni lleva manifiesto
    assert "manual_oracle" not in SRC.read_text(encoding="utf-8")
    assert not (FIXTURE / "manifest.json").exists() and (FIXTURE / "README.md").exists()
    assert "a mano" in (FIXTURE / "README.md").read_text(encoding="utf-8")


# --------------------------------------------------------------------------------------------------
# T3-14: literales del generador = config/demo.toml
# --------------------------------------------------------------------------------------------------


def test_t3_14_generator_literals_match_demo_config():
    raw = tomllib.loads((REPO / "config" / "demo.toml").read_text(encoding="utf-8"))
    run = raw["run"]
    assert synthetic.PERIOD == (run["period_start"], run["period_end"])
    assert synthetic.TOLERANCE == Decimal(run["total_tolerance"])
    assert synthetic.TARGET_CONFIG_VERSION == raw["config_version"] == 1
    profiles = {s["id"]: s for s in raw["sources"]}
    from datetime import date

    patterns = {"N": "%d/%m/%Y", "C": "%d-%m-%Y", "SX": "%Y-%m-%d", "SC": "%d/%m/%Y"}
    unmapped = {"N": set(), "C": {"Observaciones"}, "SX": {"Cajero"}, "SC": set()}
    for key, spec in synthetic.FILES.items():
        prof = profiles[spec["source_id"]]
        assert prof["branch_id"] == spec["branch"] and prof["type"] == spec["kind"]
        assert fnmatch.fnmatchcase(spec["name"], prof["match"])
        assert prof["header_row"] == spec["header_row"]
        assert patterns[key] in prof["date_formats"]
        assert date(2026, 10, 15).strftime(patterns[key])
        if spec["kind"] == "xlsx":
            assert prof["sheet"] == spec["sheet"]
        else:
            assert prof["encoding"] == spec["encoding"] and prof["delimiter"] == spec["delimiter"]
        if key == "C":
            assert (prof["decimal_separator"], prof["thousands_separator"]) == (",", ".")
        else:
            assert prof["decimal_separator"] == "."
        aliases = {canonical_text(a): f for f, al in prof["columns"].items() for a in al}
        headers = list(spec["headers"])
        mapped = {canonical_text(h) for h in headers} & set(aliases)
        not_mapped = {h for h in headers if canonical_text(h) not in aliases}
        assert not_mapped == unmapped[key], (key, not_mapped)
        required = {"transaction_id", "sale_date", "product_sku", "quantity", "unit_price"}
        assert required <= {aliases[m] for m in mapped}
        # el perfil no lista alias que no estén en el archivo salvo campos opcionales ausentes
        assert {aliases[m] for m in mapped} == set(prof["columns"])
    assert parse_config((REPO / "config" / "demo.toml").read_text(encoding="utf-8")).sample.seed == 20261001
