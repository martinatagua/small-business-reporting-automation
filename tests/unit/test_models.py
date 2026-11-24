"""TM-01 … TM-07 (F2). Las expectativas son literales independientes de la implementación."""

import ast
import dataclasses
from pathlib import Path

import pytest

from small_business_reporting_automation import models
from small_business_reporting_automation.models import (
    CANONICAL_FIELDS,
    ISSUE_SPECS,
    MAPPABLE_FIELDS,
    REQUIRED_FIELDS,
    REQUIRED_MAPPED_FIELDS,
    SEVERITY_TO_STATUS,
    STATUS_LABELS_ES,
    Issue,
    IssueCode,
    IssueLevel,
    IssueSpec,
    KpiEffect,
    RecordStatus,
    Severity,
    canonical_text,
)

SRC = Path(models.__file__).parent

EXPECTED_CODES = [
    "C01_INVALID_CONFIG",
    "F01_UNREADABLE_FILE",
    "F02_ENCODING_ERROR",
    "F03_UNSUPPORTED_TYPE",
    "F04_LIMIT_EXCEEDED",
    "F05_NO_PROFILE",
    "F06_AMBIGUOUS_PROFILE",
    "F07_SHEET_NOT_FOUND",
    "F08_REQUIRED_COLUMN_MISSING",
    "F09_COLUMN_COLLISION",
    "F10_NO_DATA_ROWS",
    "F11_UNMAPPED_COLUMN",
    "R01_MISSING_REQUIRED",
    "R02_INVALID_DATE",
    "R03_AMBIGUOUS_DATE",
    "R04_INVALID_NUMBER",
    "R05_NON_INTEGER_QUANTITY",
    "R06_ZERO_QUANTITY",
    "R07_NEGATIVE_QUANTITY",
    "R08_ZERO_PRICE",
    "R09_NEGATIVE_PRICE",
    "R10_DATE_OUT_OF_PERIOD",
    "R11_TOTAL_MISMATCH",
    "R12_UNKNOWN_SKU",
    "R13_DUPLICATE_SUSPECT",
    "R14_BRANCH_CONFLICT",
]

EXPECTED_SEVERITIES = {
    "C01_INVALID_CONFIG": "error",
    "F01_UNREADABLE_FILE": "error",
    "F02_ENCODING_ERROR": "error",
    "F03_UNSUPPORTED_TYPE": "error",
    "F04_LIMIT_EXCEEDED": "error",
    "F05_NO_PROFILE": "error",
    "F06_AMBIGUOUS_PROFILE": "error",
    "F07_SHEET_NOT_FOUND": "error",
    "F08_REQUIRED_COLUMN_MISSING": "error",
    "F09_COLUMN_COLLISION": "error",
    "F10_NO_DATA_ROWS": "warning",
    "F11_UNMAPPED_COLUMN": "info",
    "R01_MISSING_REQUIRED": "error",
    "R02_INVALID_DATE": "error",
    "R03_AMBIGUOUS_DATE": "error",
    "R04_INVALID_NUMBER": "error",
    "R05_NON_INTEGER_QUANTITY": "error",
    "R06_ZERO_QUANTITY": "error",
    "R07_NEGATIVE_QUANTITY": "scope",
    "R08_ZERO_PRICE": "warning",
    "R09_NEGATIVE_PRICE": "error",
    "R10_DATE_OUT_OF_PERIOD": "scope",
    "R11_TOTAL_MISMATCH": "error",
    "R12_UNKNOWN_SKU": "warning",
    "R13_DUPLICATE_SUSPECT": "warning",
    "R14_BRANCH_CONFLICT": "error",
}


def test_tm01_exact_26_codes_in_order():
    assert len(EXPECTED_CODES) == 26
    assert [c.value for c in IssueCode] == EXPECTED_CODES
    assert [c.name for c in IssueCode] == EXPECTED_CODES
    assert list(ISSUE_SPECS) == list(IssueCode)


def test_tm02_severities_and_pending_decisions():
    assert {c.value: s.severity.value for c, s in ISSUE_SPECS.items()} == EXPECTED_SEVERITIES
    # P-01, P-02 y P-03 aprobados por el propietario (2026-11-20): ningún código queda pendiente.
    pending = {c.value: s.pending_decision for c, s in ISSUE_SPECS.items() if s.pending_decision is not None}
    assert pending == {}
    for code in ("R06_ZERO_QUANTITY", "R10_DATE_OUT_OF_PERIOD", "R11_TOTAL_MISMATCH"):
        assert ISSUE_SPECS[IssueCode(code)].pending_decision is None
    assert all(s.code is c for c, s in ISSUE_SPECS.items())
    with pytest.raises(TypeError):
        ISSUE_SPECS[IssueCode.R01_MISSING_REQUIRED] = None  # type: ignore[index]
    with pytest.raises(dataclasses.FrozenInstanceError):
        ISSUE_SPECS[IssueCode.R06_ZERO_QUANTITY].pending_decision = "P-01"  # type: ignore[misc]
    assert issubclass(IssueSpec, object) and IssueSpec.__dataclass_params__.frozen


def test_tm03_enums_and_labels():
    assert [s.value for s in Severity] == ["error", "scope", "warning", "info"]
    assert [s.value for s in RecordStatus] == ["valid", "observed", "out_of_scope"]
    assert [k.value for k in KpiEffect] == [
        "included",
        "included_warning",
        "excluded_observed",
        "excluded_out_of_scope",
    ]
    assert [l.value for l in IssueLevel] == ["file", "record"]
    assert dict(STATUS_LABELS_ES) == {
        "valid": "válido",
        "observed": "observado",
        "out_of_scope": "fuera de alcance",
    }
    assert dict(SEVERITY_TO_STATUS) == {"error": "observed", "scope": "out_of_scope"}
    assert "warning" not in SEVERITY_TO_STATUS and "info" not in SEVERITY_TO_STATUS


def test_tm04_field_constants():
    assert CANONICAL_FIELDS == (
        "branch_id", "transaction_id", "sale_date", "product_sku", "product_name",
        "quantity", "unit_price", "reported_total", "branch_in_file", "line_amount",
    )
    assert REQUIRED_FIELDS == (
        "branch_id", "transaction_id", "sale_date", "product_sku", "quantity", "unit_price",
    )
    assert MAPPABLE_FIELDS == (
        "transaction_id", "sale_date", "product_sku", "product_name",
        "quantity", "unit_price", "reported_total", "branch_in_file",
    )
    assert REQUIRED_MAPPED_FIELDS == (
        "transaction_id", "sale_date", "product_sku", "quantity", "unit_price",
    )
    assert "branch_id" not in MAPPABLE_FIELDS and "line_amount" not in MAPPABLE_FIELDS


def _record_issue(**kw):
    base = dict(
        level=IssueLevel.RECORD,
        record_id="norte-xlsx:a.xlsx:Ventas:5",
        issue_code=IssueCode.R01_MISSING_REQUIRED,
        severity=Severity.ERROR,
        resulting_status=RecordStatus.OBSERVED,
        message="x",
    )
    base.update(kw)
    return Issue(**base)


def test_tm05_issue_fields_order_and_defaults():
    assert [f.name for f in dataclasses.fields(Issue)] == [
        "level", "record_id", "source_id", "source_file", "source_sheet", "source_row", "field",
        "issue_code", "detail_code", "severity", "resulting_status", "raw_value", "message",
        "related_record_ids",
    ]
    i = _record_issue()
    assert i.related_record_ids == ()
    assert isinstance(i.related_record_ids, tuple)


def test_tm05_issue_invariants():
    # nivel file sin record_id ni resulting_status
    ok = Issue(level=IssueLevel.FILE, issue_code=IssueCode.F10_NO_DATA_ROWS, severity=Severity.WARNING, message="m")
    assert ok.record_id is None and ok.resulting_status is None
    with pytest.raises(ValueError):
        Issue(level=IssueLevel.FILE, record_id="r", issue_code=IssueCode.F10_NO_DATA_ROWS,
              severity=Severity.WARNING, message="m")
    with pytest.raises(ValueError):
        Issue(level=IssueLevel.FILE, issue_code=IssueCode.F03_UNSUPPORTED_TYPE, severity=Severity.ERROR,
              resulting_status=RecordStatus.OBSERVED, message="m")
    # detail_code solo con F01/F04
    for code in (IssueCode.F01_UNREADABLE_FILE, IssueCode.F04_LIMIT_EXCEEDED):
        Issue(level=IssueLevel.FILE, issue_code=code, severity=Severity.ERROR, detail_code="x", message="m")
    with pytest.raises(ValueError):
        Issue(level=IssueLevel.FILE, issue_code=IssueCode.F03_UNSUPPORTED_TYPE, severity=Severity.ERROR,
              detail_code="x", message="m")
    with pytest.raises(ValueError):
        _record_issue(detail_code="x")
    # tupla obligatoria
    with pytest.raises(TypeError):
        _record_issue(related_record_ids=["a"])
    assert _record_issue(related_record_ids=("a", "b")).related_record_ids == ("a", "b")
    # C01 no es una incidencia de registro ni de archivo
    for lvl in IssueLevel:
        with pytest.raises(ValueError):
            Issue(level=lvl, issue_code=IssueCode.C01_INVALID_CONFIG, severity=Severity.ERROR, message="m")
    # coherencia de severidad y de familia código/nivel
    with pytest.raises(ValueError):
        _record_issue(severity=Severity.WARNING)
    with pytest.raises(ValueError):
        Issue(level=IssueLevel.RECORD, issue_code=IssueCode.F10_NO_DATA_ROWS, severity=Severity.WARNING, message="m")


def test_tm05_issue_frozen_and_slots():
    i = _record_issue()
    with pytest.raises(dataclasses.FrozenInstanceError):
        i.message = "otro"  # type: ignore[misc]
    assert not hasattr(i, "__dict__")
    assert Issue.__dataclass_params__.frozen and "__slots__" in vars(Issue)


def test_canonical_text_contract():
    assert canonical_text("Nº Ticket") == "no ticket"  # §12.1; NFKD: º → o
    assert canonical_text("N. Ticket") == "n ticket"
    assert canonical_text("  Descripción ") == "descripcion"
    assert canonical_text("PRECIO--Unitario__x") == "precio unitario x"
    assert canonical_text("ＦＥＣＨＡ") == "fecha"  # NFKD de ancho completo
    assert canonical_text("Straße") == "strasse"  # casefold
    assert canonical_text("---") == ""


@pytest.mark.parametrize("module", ["config.py", "models.py"])
def test_tm06_no_pandas_openpyxl_imports(module):
    tree = ast.parse((SRC / module).read_text(encoding="utf-8"))
    imported = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            imported |= {a.name.split(".")[0] for a in node.names}
        elif isinstance(node, ast.ImportFrom) and node.module:
            imported.add(node.module.split(".")[0])
    assert not imported & {"pandas", "openpyxl", "numpy"}


def test_tm07_skeleton_tests_still_present():
    text = (Path(__file__).parent / "test_skeleton.py").read_text(encoding="utf-8")
    for name in ("test_package_imports", "test_entry_point_declared", "test_main_reports_pending", "test_help_exits_zero"):
        assert f"def {name}" in text
