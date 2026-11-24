"""TC-P01 … TC-P08 y TC-N01 … TC-N31 (F2).

Las configuraciones negativas parten de un TOML mínimo válido (texto literal) al que se cambia una línea.
Los números de ``[limits]`` de ``config/demo.toml`` son provisionales (P-12): no se fijan aquí como
umbrales; solo se comprueba su coherencia con lo que devuelve ``tomllib``.
"""

import ast
import dataclasses
import tomllib
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

import small_business_reporting_automation.config as config_module
from small_business_reporting_automation.config import (
    AppConfig,
    ConfigError,
    ConfigProblem,
    Limits,
    RunSettings,
    SampleSettings,
    SourceProfile,
    load_config,
    parse_config,
)
from small_business_reporting_automation.models import IssueCode

REPO = Path(__file__).resolve().parents[2]
DEMO = REPO / "config" / "demo.toml"

LIMIT_KEYS = [
    "max_file_mb", "max_zip_entries", "max_total_uncompressed_mb", "max_entry_uncompressed_mb",
    "max_declared_compression_ratio", "max_shared_strings", "max_shared_strings_chars", "max_sheets",
    "max_columns", "max_cells_per_file", "max_text_chars", "max_records_per_file", "max_records_per_run",
]

RUN = """[run]
period_start = 2026-10-01
period_end = 2026-10-31
null_tokens = ["", "N/A"]
total_tolerance = "0.01"
duplicate_policy = "warn"
"""

LIMITS = "[limits]\n" + "".join(f"{k} = 10\n" for k in LIMIT_KEYS)

COLUMNS = """[sources.columns]
transaction_id = ["ticket"]
sale_date = ["fecha"]
product_sku = ["sku"]
quantity = ["cantidad"]
unit_price = ["precio"]
"""

XLSX = (
    """[[sources]]
id = "a-xlsx"
branch_id = "A"
match = "a_*.xlsx"
type = "xlsx"
sheet = "Ventas"
header_row = 1
date_formats = ["%d/%m/%Y"]
accept_native_dates = true
accept_serial_dates = false
decimal_separator = "."
thousands_separator = ""
"""
    + COLUMNS
)

CSV = (
    """[[sources]]
id = "b-csv"
branch_id = "B"
match = "b_*.csv"
type = "csv"
encoding = "utf-8"
delimiter = ";"
header_row = 1
date_formats = ["%Y-%m-%d"]
decimal_separator = ","
thousands_separator = "."
"""
    + COLUMNS
)


def build(*, version="config_version = 1\n", run=RUN, limits=LIMITS, sources=(XLSX,)):
    return version + run + limits + "".join(sources)


MIN = build()


def sub(block: str, old: str, new: str) -> str:
    assert block.count(old) == 1, f"{old!r} debe aparecer exactamente una vez"
    return block.replace(old, new)


def drop(block: str, key: str) -> str:
    lines = block.splitlines(keepends=True)
    kept = [l for l in lines if not l.startswith(f"{key} =")]
    assert len(kept) == len(lines) - 1, key
    return "".join(kept)


def errors(text: str) -> list[tuple[str, str]]:
    with pytest.raises(ConfigError) as exc:
        parse_config(text)
    assert exc.value.code == IssueCode.C01_INVALID_CONFIG
    assert exc.value.problems
    for pr in exc.value.problems:
        assert isinstance(pr, ConfigProblem) and pr.path is not None and pr.reason and pr.message
    return [(pr.path, pr.reason) for pr in exc.value.problems]


# ---------------------------------------------------------------------------------------------------
# Positivas
# ---------------------------------------------------------------------------------------------------


def test_tc_p01_demo_loads():
    cfg = load_config(DEMO)
    raw = tomllib.loads(DEMO.read_text(encoding="utf-8"))
    assert isinstance(cfg, AppConfig) and cfg.config_version == 1
    assert [s.id for s in cfg.sources] == ["norte-xlsx", "centro-csv", "sur-xlsx", "sur-csv"]
    assert [s.type for s in cfg.sources] == ["xlsx", "csv", "xlsx", "csv"]
    assert [s.branch_id for s in cfg.sources] == ["NORTE", "CENTRO", "SUR", "SUR"]
    assert cfg.run.period_start == date(2026, 10, 1) and type(cfg.run.period_start) is date
    assert cfg.run.period_end == date(2026, 10, 31)
    assert cfg.run.null_tokens == ("", "N/A", "NA", "n/a", "-", "null")
    assert cfg.run.total_tolerance == Decimal("0.01") and isinstance(cfg.run.total_tolerance, Decimal)
    assert cfg.run.duplicate_policy == "warn"
    assert cfg.run.catalog_path == "data/reference/products.csv"
    assert [f.name for f in dataclasses.fields(Limits)] == LIMIT_KEYS
    assert dataclasses.asdict(cfg.limits) == raw["limits"]  # provisionales (P-12): sin fijar valores
    norte, centro, sur_x, sur_c = cfg.sources
    assert (norte.sheet, norte.header_row, norte.accept_native_dates, norte.accept_serial_dates) == ("Ventas", 3, True, False)
    assert norte.encoding is None and norte.delimiter is None
    assert (centro.encoding, centro.delimiter, centro.decimal_separator, centro.thousands_separator) == ("cp1252", ";", ",", ".")
    assert centro.sheet is None and centro.accept_native_dates is None
    assert sur_x.branch_aliases == ("sur", "sucursal sur") and norte.branch_aliases == ()
    assert sur_c.date_formats == ("%d/%m/%Y", "%m/%d/%Y")
    assert norte.columns["transaction_id"] == ("no ticket", "n ticket")
    assert list(sur_x.columns) == [
        "transaction_id", "sale_date", "product_sku", "product_name", "quantity", "unit_price", "branch_in_file",
    ]


def test_tc_p02_minimal_valid():
    cfg = parse_config(MIN)
    assert cfg.run.catalog_path is None
    s = cfg.sources[0]
    assert s.branch_aliases == () and s.currency_symbols == ()
    assert dict(s.columns) == {
        "transaction_id": ("ticket",), "sale_date": ("fecha",), "product_sku": ("sku",),
        "quantity": ("cantidad",), "unit_price": ("precio",),
    }
    assert cfg.limits.max_file_mb == 10


def test_tc_p03_same_alias_across_profiles_is_valid():
    cfg = parse_config(build(sources=(XLSX, sub(CSV, 'id = "b-csv"', 'id = "other"'))))
    assert len(cfg.sources) == 2
    assert cfg.sources[0].columns["sale_date"] == cfg.sources[1].columns["sale_date"] == ("fecha",)
    # el mismo alias dos veces en el MISMO campo no es colisión entre campos
    dup = sub(XLSX, 'sale_date = ["fecha"]', 'sale_date = ["fecha", "Fecha"]')
    assert parse_config(build(sources=(dup,))).sources[0].columns["sale_date"] == ("fecha", "Fecha")


def test_tc_p04_limit_relations_at_boundary():
    limits = sub(sub(LIMITS, "max_entry_uncompressed_mb = 10", "max_entry_uncompressed_mb = 7"),
                 "max_total_uncompressed_mb = 10", "max_total_uncompressed_mb = 7")
    limits = sub(sub(limits, "max_records_per_file = 10", "max_records_per_file = 3"),
                 "max_records_per_run = 10", "max_records_per_run = 3")
    cfg = parse_config(build(limits=limits))
    assert cfg.limits.max_entry_uncompressed_mb == cfg.limits.max_total_uncompressed_mb == 7
    assert cfg.limits.max_records_per_file == cfg.limits.max_records_per_run == 3


@pytest.mark.parametrize("fmt", ["%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%y", "%Y%%%m%%%d", "%d %m %Y"])
def test_tc_p05_valid_date_formats(fmt):
    cfg = parse_config(build(sources=(sub(XLSX, '["%d/%m/%Y"]', f'["{fmt}"]'),)))
    assert cfg.sources[0].date_formats == (fmt,)


@pytest.mark.parametrize("dec,thou", [(".", ""), (",", ""), (",", ".")])
def test_tc_p06_separator_combinations(dec, thou):
    src = sub(sub(XLSX, 'decimal_separator = "."', f'decimal_separator = "{dec}"'),
              'thousands_separator = ""', f'thousands_separator = "{thou}"')
    s = parse_config(build(sources=(src,))).sources[0]
    assert (s.decimal_separator, s.thousands_separator) == (dec, thou)


def test_tc_p07_immutability():
    cfg = load_config(DEMO)
    for obj, attr in [
        (cfg, "config_version"), (cfg.run, "period_start"), (cfg.limits, "max_columns"),
        (cfg.sources[0], "id"), (ConfigProblem("p", "r", "m"), "path"),
    ]:
        with pytest.raises(dataclasses.FrozenInstanceError):
            setattr(obj, attr, None)
        assert not hasattr(obj, "__dict__")  # slots
    assert isinstance(cfg.sources, tuple)
    assert isinstance(cfg.run.null_tokens, tuple)
    assert isinstance(cfg.sources[0].date_formats, tuple)
    assert isinstance(cfg.sources[0].columns["sale_date"], tuple)
    with pytest.raises(TypeError):
        cfg.sources[0].columns["x"] = ("y",)  # type: ignore[index]
    with pytest.raises(TypeError):
        del cfg.sources[0].columns["sale_date"]  # type: ignore[attr-defined]
    for cls in (AppConfig, RunSettings, Limits, SourceProfile, ConfigProblem):
        assert cls.__dataclass_params__.frozen and "__slots__" in vars(cls)


def test_tc_p08_all_limits_at_minimum():
    limits = "[limits]\n" + "".join(f"{k} = 1\n" for k in LIMIT_KEYS)
    cfg = parse_config(build(limits=limits))
    assert all(getattr(cfg.limits, k) == 1 for k in LIMIT_KEYS)


# ---------------------------------------------------------------------------------------------------
# Negativas: lectura, UTF-8 y sintaxis
# ---------------------------------------------------------------------------------------------------


def test_tc_n01_toml_syntax():
    assert errors("config_version = = 1") == [("", "toml_syntax")]
    assert errors(MIN + "\n[[sources") == [("", "toml_syntax")]


def test_tc_n02_duplicate_key():
    assert errors(MIN.replace("config_version = 1\n", "config_version = 1\nconfig_version = 1\n", 1)) == [("", "toml_syntax")]
    assert errors(sub(MIN, 'duplicate_policy = "warn"\n', 'duplicate_policy = "warn"\nduplicate_policy = "warn"\n')) == [("", "toml_syntax")]


def test_tc_n03_not_utf8(tmp_path):
    f = tmp_path / "c.toml"
    f.write_bytes(MIN.replace("Ventas", "Ventás").encode("cp1252"))
    with pytest.raises(ConfigError) as exc:
        load_config(f)
    assert [(p.reason) for p in exc.value.problems] == ["invalid_utf8"]
    assert exc.value.code == IssueCode.C01_INVALID_CONFIG
    f.write_bytes(b"\xef\xbb\xbf" + MIN.encode())  # BOM: no es TOML válido para tomllib
    with pytest.raises(ConfigError) as exc2:
        load_config(f)
    assert exc2.value.problems[0].reason == "toml_syntax"


def test_tc_n04_missing_file_and_directory(tmp_path):
    with pytest.raises(ConfigError) as exc:
        load_config(tmp_path / "nope.toml")
    assert [p.reason for p in exc.value.problems] == ["file_not_found"]
    with pytest.raises(ConfigError) as exc:
        load_config(tmp_path)
    assert [p.reason for p in exc.value.problems] == ["file_unreadable"]


def test_load_config_accepts_str_and_path(tmp_path):
    f = tmp_path / "c.toml"
    f.write_text(MIN, encoding="utf-8")
    assert load_config(f) == load_config(str(f))


# ---------------------------------------------------------------------------------------------------
# Negativas: claves
# ---------------------------------------------------------------------------------------------------


@pytest.mark.parametrize(
    "text,path",
    [
        ("extra = 1\n" + MIN, "extra"),
        (sub(MIN, 'duplicate_policy = "warn"\n', 'duplicate_policy = "warn"\nextra = 1\n'), "run.extra"),
        (sub(MIN, "max_file_mb = 10\n", "max_file_mb = 10\nextra = 1\n"), "limits.extra"),
        (sub(MIN, 'id = "a-xlsx"\n', 'id = "a-xlsx"\nextra = 1\n'), "sources[0].extra"),
        (sub(MIN, 'quantity = ["cantidad"]', 'quantity = ["cantidad"]\ncantidad = ["x"]'), "sources[0].columns.cantidad"),
        (sub(MIN, 'quantity = ["cantidad"]', 'quantity = ["cantidad"]\nbranch_id = ["x"]'), "sources[0].columns.branch_id"),
        (sub(MIN, 'quantity = ["cantidad"]', 'quantity = ["cantidad"]\nline_amount = ["x"]'), "sources[0].columns.line_amount"),
    ],
)
def test_tc_n05_unknown_keys(text, path):
    assert (path, "unknown_key") in errors(text)


def _missing_cases():
    cases = [("config_version", build(version="")), ("run", build(run="")), ("limits", build(limits="")),
             ("sources", build(sources=()))]
    for k in ("period_start", "period_end", "null_tokens", "total_tolerance", "duplicate_policy"):
        cases.append((f"run.{k}", build(run=drop(RUN, k))))
    for k in LIMIT_KEYS:
        cases.append((f"limits.{k}", build(limits=drop(LIMITS, k))))
    for k in ("id", "branch_id", "match", "type", "header_row", "date_formats", "decimal_separator", "thousands_separator"):
        cases.append((f"sources[0].{k}", build(sources=(drop(XLSX, k),))))
    for k in ("sheet", "accept_native_dates", "accept_serial_dates"):
        cases.append((f"sources[0].{k}", build(sources=(drop(XLSX, k),))))
    for k in ("encoding", "delimiter"):
        cases.append((f"sources[0].{k}", build(sources=(drop(CSV, k),))))
    cases.append(("sources[0].columns", build(sources=(XLSX.replace(COLUMNS, ""),))))
    return cases


@pytest.mark.parametrize("path,text", _missing_cases(), ids=[c[0] for c in _missing_cases()])
def test_tc_n06_missing_required_key(path, text):
    if path == "sources":
        pass
    got = errors(text)
    assert (path, "missing_key") in got


def test_tc_n06_no_defaults_every_limit_required():
    assert len([c for c in _missing_cases() if c[0].startswith("limits.")]) == 13


# ---------------------------------------------------------------------------------------------------
# Negativas: tipos
# ---------------------------------------------------------------------------------------------------


@pytest.mark.parametrize(
    "text,path",
    [
        (sub(MIN, "period_start = 2026-10-01", 'period_start = "2026-10-01"'), "run.period_start"),
        (sub(MIN, "period_start = 2026-10-01", "period_start = 2026-10-01T00:00:00"), "run.period_start"),
        (sub(MIN, "period_end = 2026-10-31", "period_end = 2026-10-31T10:00:00+01:00"), "run.period_end"),
        (sub(MIN, "period_end = 2026-10-31", "period_end = 10:00:00"), "run.period_end"),
        (sub(MIN, "header_row = 1", 'header_row = "3"'), "sources[0].header_row"),
        (sub(MIN, "header_row = 1", "header_row = true"), "sources[0].header_row"),
        (sub(MIN, "header_row = 1", "header_row = 1.0"), "sources[0].header_row"),
        (sub(MIN, 'null_tokens = ["", "N/A"]', 'null_tokens = ""'), "run.null_tokens"),
        (sub(MIN, 'null_tokens = ["", "N/A"]', "null_tokens = [1]"), "run.null_tokens[0]"),
        (sub(MIN, 'sale_date = ["fecha"]', 'sale_date = "fecha"'), "sources[0].columns.sale_date"),
        (sub(MIN, "accept_native_dates = true", 'accept_native_dates = "true"'), "sources[0].accept_native_dates"),
        (sub(MIN, "accept_serial_dates = false", "accept_serial_dates = 0"), "sources[0].accept_serial_dates"),
        (sub(MIN, "config_version = 1", "config_version = true"), "config_version"),
        (sub(MIN, "config_version = 1", 'config_version = "1"'), "config_version"),
        (sub(MIN, 'date_formats = ["%d/%m/%Y"]', 'date_formats = "%d/%m/%Y"'), "sources[0].date_formats"),
        (sub(MIN, 'sheet = "Ventas"', "sheet = 3"), "sources[0].sheet"),
        (sub(MIN, 'id = "a-xlsx"', "id = 7"), "sources[0].id"),
        (MIN.replace("[run]", "run = 1\n[zz]", 1), "run"),
        (sub(MIN, 'decimal_separator = "."', "decimal_separator = 1"), "sources[0].decimal_separator"),
        (sub(MIN, 'duplicate_policy = "warn"', "duplicate_policy = true"), "run.duplicate_policy"),
        (sub(MIN, 'duplicate_policy = "warn"', 'duplicate_policy = "warn"\ncatalog_path = 5'), "run.catalog_path"),
    ],
)
def test_tc_n07_wrong_types(text, path):
    assert (path, "invalid_type") in errors(text)


def test_tc_n07_sources_wrong_shape():
    assert ("sources", "invalid_type") in errors(MIN.replace("[[sources]]", "[sources]", 1))
    assert ("sources[0]", "invalid_type") in errors(build(version="config_version = 1\nsources = [1]\n", sources=()))


# ---------------------------------------------------------------------------------------------------
# Negativas: formatos de fecha, separadores, ids, alias
# ---------------------------------------------------------------------------------------------------


@pytest.mark.parametrize(
    "fmt",
    ["%m/%Y", "%d/%Y", "%d/%m", "%d/%m/%Y %H:%M", "%d/%d/%Y", "%d/%m/%Y/%y", "%d/%m/%Y %z", "%d/%m/%Y %", "%d %B %Y",
     "%b %d %Y", "", "abc", "%j %Y", "%d%%%m"],
)
def test_tc_n08_invalid_date_formats(fmt):
    text = build(sources=(sub(XLSX, '["%d/%m/%Y"]', f"['{fmt}']"),))
    assert ("sources[0].date_formats[0]", "invalid_date_format") in errors(text)


def test_tc_n08_round_trip_check_is_executed(monkeypatch):
    # Con la lista blanca de PO-1 no hay formato orgánico que falle la ida y vuelta con las fechas de
    # referencia reales; se comprueba que el paso existe usando una referencia fuera del rango de %y.
    text = build(sources=(sub(XLSX, '["%d/%m/%Y"]', "['%d/%m/%y']"),))
    parse_config(text)  # válido con las referencias reales
    monkeypatch.setattr(config_module, "_REFERENCE_DATES", (date(2075, 1, 1),))  # %y: 75 -> 1975
    assert ("sources[0].date_formats[0]", "invalid_date_format") in errors(text)


def test_tc_n09_empty_date_formats():
    assert ("sources[0].date_formats", "empty_value") in errors(build(sources=(sub(XLSX, '["%d/%m/%Y"]', "[]"),)))


@pytest.mark.parametrize("sep", [".", ","])
def test_tc_n10_equal_separators(sep):
    src = sub(sub(XLSX, 'decimal_separator = "."', f'decimal_separator = "{sep}"'),
              'thousands_separator = ""', f'thousands_separator = "{sep}"')
    assert ("sources[0].decimal_separator", "separator_conflict") in errors(build(sources=(src,)))
    both_empty = sub(XLSX, 'decimal_separator = "."', 'decimal_separator = ""')
    assert ("sources[0].decimal_separator", "separator_conflict") in errors(build(sources=(both_empty,)))


@pytest.mark.parametrize("dec,thou", [(".", ","), ("'", ""), (",", " "), ("", "."), (";", "")])
def test_tc_n11_unsupported_separator_combinations(dec, thou):
    src = sub(sub(XLSX, 'decimal_separator = "."', f'decimal_separator = "{dec}"'),
              'thousands_separator = ""', f'thousands_separator = "{thou}"')
    assert ("sources[0].decimal_separator", "unsupported_separators") in errors(build(sources=(src,)))


def test_tc_n12_duplicate_ids():
    got = errors(build(sources=(XLSX, sub(CSV, 'id = "b-csv"', 'id = "a-xlsx"'))))
    assert got == [("sources[1].id", "duplicate_id")]


def test_tc_n13_alias_collision_within_profile():
    one = sub(XLSX, 'transaction_id = ["ticket"]', 'transaction_id = ["fecha"]')
    assert errors(build(sources=(one,))) == [("sources[0].columns.sale_date", "alias_collision")]
    # PO-2: solo difieren en mayúsculas, acentos, puntuación o espacios
    for variant in ["FECHA", "Fécha", "  fecha  ", "ＦＥＣＨＡ", "fecha!", "_fecha_"]:
        two = sub(XLSX, 'transaction_id = ["ticket"]', f'transaction_id = ["{variant}"]')
        assert ("sources[0].columns.sale_date", "alias_collision") in errors(build(sources=(two,))), variant
    # puntuación interna colapsa a espacio: "Fecha  Venta" == "fecha-venta"
    a = sub(sub(XLSX, 'transaction_id = ["ticket"]', 'transaction_id = ["Fecha  Venta"]'),
            'quantity = ["cantidad"]', 'quantity = ["fecha-venta"]')
    assert ("sources[0].columns.quantity", "alias_collision") in errors(build(sources=(a,)))


def test_tc_n13_only_the_colliding_profile_reports():
    bad = sub(CSV, 'transaction_id = ["ticket"]', 'transaction_id = ["SKU"]')
    bad = sub(bad, 'id = "b-csv"', 'id = "c"')
    got = errors(build(sources=(XLSX, bad)))
    assert got == [("sources[1].columns.product_sku", "alias_collision")]


@pytest.mark.parametrize("text,reason", [
    (build(sources=(sub(CSV, 'encoding = "utf-8"\n', ""),)), "missing_key"),
    (build(sources=(sub(CSV, 'delimiter = ";"\n', ""),)), "missing_key"),
])
def test_tc_n14_csv_missing_encoding_or_delimiter(text, reason):
    assert reason in [r for _, r in errors(text)]


@pytest.mark.parametrize("literal", ['""', '";;"', "'\"'", '"\\n"', '"\\r"'])
def test_tc_n15_invalid_delimiter(literal):
    text = build(sources=(sub(CSV, 'delimiter = ";"', f"delimiter = {literal}"),))
    assert errors(text) == [("sources[0].delimiter", "invalid_delimiter")]


@pytest.mark.parametrize("match", ["", "   "])
def test_tc_n16_empty_match(match):
    assert ("sources[0].match", "empty_value") in errors(build(sources=(sub(XLSX, 'match = "a_*.xlsx"', f'match = "{match}"'),)))


# ---------------------------------------------------------------------------------------------------
# Negativas: límites, política, tolerancia
# ---------------------------------------------------------------------------------------------------

_BAD_LIMIT_VALUES = [("0", "non_positive"), ("-1", "non_positive"), ("1.5", "invalid_type"),
                     ('"20"', "invalid_type"), ("true", "invalid_type")]


@pytest.mark.parametrize("key", LIMIT_KEYS)
@pytest.mark.parametrize("value,reason", _BAD_LIMIT_VALUES)
def test_tc_n17_limit_values(key, value, reason):
    limits = sub(LIMITS, f"{key} = 10", f"{key} = {value}")
    assert (f"limits.{key}", reason) in errors(build(limits=limits))


def test_tc_n18_entry_greater_than_total():
    limits = sub(LIMITS, "max_entry_uncompressed_mb = 10", "max_entry_uncompressed_mb = 11")
    assert errors(build(limits=limits)) == [("limits.max_entry_uncompressed_mb", "limit_relation")]


def test_tc_n19_records_per_file_greater_than_per_run():
    limits = sub(LIMITS, "max_records_per_file = 10", "max_records_per_file = 11")
    assert errors(build(limits=limits)) == [("limits.max_records_per_file", "limit_relation")]


def test_tc_n19_relation_skipped_when_dependency_invalid():
    limits = sub(LIMITS, "max_records_per_run = 10", "max_records_per_run = 0")
    assert errors(build(limits=limits)) == [("limits.max_records_per_run", "non_positive")]


def test_tc_n20_duplicate_policy():
    assert ("run.duplicate_policy", "unsupported_policy") in errors(sub(MIN, 'duplicate_policy = "warn"', 'duplicate_policy = "drop"'))


@pytest.mark.parametrize(
    "literal,reason",
    [("0.01", "invalid_type"), ("1", "invalid_type"), ('"abc"', "invalid_tolerance"), ('"NaN"', "invalid_tolerance"),
     ('"Infinity"', "invalid_tolerance"), ('"-inf"', "invalid_tolerance"), ('"1e-2"', "invalid_tolerance"),
     ('""', "invalid_tolerance"), ('"0,01"', "invalid_tolerance"), ('" 0.01"', "invalid_tolerance"),
     ('"0.01\\n"', "invalid_tolerance"), ('"١.٥"', "invalid_tolerance"), ('"-0.01"', "negative_tolerance")],
)
def test_tc_n21_tolerance(literal, reason):
    assert ("run.total_tolerance", reason) in errors(sub(MIN, 'total_tolerance = "0.01"', f"total_tolerance = {literal}"))


@pytest.mark.parametrize("value", ["0", "0.00", "+0.5", "12"])
def test_tc_n21_tolerance_valid_strings_are_exact_decimal(value):
    cfg = parse_config(sub(MIN, 'total_tolerance = "0.01"', f'total_tolerance = "{value}"'))
    assert cfg.run.total_tolerance == Decimal(value) and str(cfg.run.total_tolerance) == str(Decimal(value))


def test_tc_n21_no_float_conversion_in_source():
    tree = ast.parse(Path(config_module.__file__).read_text(encoding="utf-8"))
    names = {n.id for n in ast.walk(tree) if isinstance(n, ast.Name)}
    assert "float" not in names
    assert Decimal("0.1") == parse_config(sub(MIN, 'total_tolerance = "0.01"', 'total_tolerance = "0.1"')).run.total_tolerance


@pytest.mark.parametrize("t", ["xls", "XLSX", "csv ", "", "xlsm"])
def test_tc_n22_source_type(t):
    got = errors(build(sources=(sub(XLSX, 'type = "xlsx"', f'type = "{t}"'),)))
    assert got == [("sources[0].type", "unsupported_source_type")]


@pytest.mark.parametrize("n", [0, -1])
def test_tc_n23_header_row_below_one(n):
    assert ("sources[0].header_row", "non_positive") in errors(sub(MIN, "header_row = 1", f"header_row = {n}"))


def test_tc_n24_config_version():
    assert ("config_version", "unsupported_version") in errors(sub(MIN, "config_version = 1", "config_version = 2"))
    assert ("config_version", "unsupported_version") in errors(sub(MIN, "config_version = 1", "config_version = 0"))


def test_tc_n25_period_order():
    text = sub(sub(MIN, "period_start = 2026-10-01", "period_start = 2026-11-01"), "period_end = 2026-10-31", "period_end = 2026-10-31")
    assert errors(text) == [("run.period_start", "invalid_period")]
    equal = sub(MIN, "period_end = 2026-10-31", "period_end = 2026-10-01")
    assert parse_config(equal).run.period_start == parse_config(equal).run.period_end


def test_tc_n26_sources_missing_or_empty():
    assert ("sources", "missing_key") in errors(build(sources=()))
    assert errors(build(version="config_version = 1\nsources = []\n", sources=())) == [("sources", "no_sources")]


def test_tc_n27_keys_of_the_other_type():
    csv_with_sheet = sub(CSV, 'encoding = "utf-8"', 'encoding = "utf-8"\nsheet = "X"')
    assert ("sources[0].sheet", "key_not_allowed_for_type") in errors(build(sources=(csv_with_sheet,)))
    csv_with_accept = sub(CSV, 'encoding = "utf-8"', 'encoding = "utf-8"\naccept_native_dates = true')
    assert ("sources[0].accept_native_dates", "key_not_allowed_for_type") in errors(build(sources=(csv_with_accept,)))
    xlsx_with_encoding = sub(XLSX, 'sheet = "Ventas"', 'sheet = "Ventas"\nencoding = "utf-8"')
    assert ("sources[0].encoding", "key_not_allowed_for_type") in errors(build(sources=(xlsx_with_encoding,)))
    xlsx_with_delim = sub(XLSX, 'sheet = "Ventas"', 'sheet = "Ventas"\ndelimiter = ";"')
    assert ("sources[0].delimiter", "key_not_allowed_for_type") in errors(build(sources=(xlsx_with_delim,)))


@pytest.mark.parametrize("field", ["transaction_id", "sale_date", "product_sku", "quantity", "unit_price"])
def test_tc_n28_required_mapped_fields(field):
    text = build(sources=(drop(XLSX, field),))
    assert errors(text) == [(f"sources[0].columns.{field}", "missing_key")]


def test_tc_n28_optional_mapped_fields_stay_optional():
    # reported_total, product_name y branch_in_file son opcionales (no son F08)
    assert "product_name" not in parse_config(MIN).sources[0].columns


def test_tc_n29_empty_alias_list_and_aliases():
    assert ("sources[0].columns.sale_date", "empty_value") in errors(sub(MIN, 'sale_date = ["fecha"]', "sale_date = []"))
    assert ("sources[0].columns.sale_date[0]", "empty_value") in errors(sub(MIN, 'sale_date = ["fecha"]', 'sale_date = [""]'))
    assert ("sources[0].columns.sale_date[0]", "empty_value") in errors(sub(MIN, 'sale_date = ["fecha"]', 'sale_date = ["  "]'))
    assert ("sources[0].columns.sale_date", "empty_alias") in errors(sub(MIN, 'sale_date = ["fecha"]', 'sale_date = ["--"]'))


def test_tc_n30_deterministic_aggregation():
    text = build(
        version="config_version = 2\n",
        run=sub(sub(RUN, "period_start = 2026-10-01", "period_start = 2026-12-01"), 'duplicate_policy = "warn"', 'duplicate_policy = "x"'),
        limits=sub(LIMITS, "max_sheets = 10", "max_sheets = 0"),
        sources=(sub(XLSX, "header_row = 1", "header_row = 0"),),
    )
    expected = [
        ("config_version", "unsupported_version"),
        ("run.period_start", "invalid_period"),
        ("run.duplicate_policy", "unsupported_policy"),
        ("limits.max_sheets", "non_positive"),
        ("sources[0].header_row", "non_positive"),
    ]
    assert errors(text) == expected
    assert errors(text) == expected  # estable entre ejecuciones
    with pytest.raises(ConfigError) as exc:
        parse_config(text)
    assert isinstance(exc.value.problems, tuple) and len(exc.value.problems) == 5
    assert all(pr.message for pr in exc.value.problems)
    assert "C01_INVALID_CONFIG" in str(exc.value)


def test_tc_n31_unknown_encoding():
    assert ("sources[0].encoding", "invalid_encoding") in errors(build(sources=(sub(CSV, 'encoding = "utf-8"', 'encoding = "utf-99"'),)))
    assert ("sources[0].encoding", "empty_value") in errors(build(sources=(sub(CSV, 'encoding = "utf-8"', 'encoding = ""'),)))
    for ok in ("cp1252", "utf-8-sig", "latin-1"):
        parse_config(build(sources=(sub(CSV, 'encoding = "utf-8"', f'encoding = "{ok}"'),)))


def test_dependent_errors_are_skipped():
    # date_formats no es lista: no se revisa cada formato; separadores inválidos no generan más errores
    got = errors(sub(MIN, 'date_formats = ["%d/%m/%Y"]', "date_formats = 5"))
    assert got == [("sources[0].date_formats", "invalid_type")]
    # un perfil con type inválido no genera errores de claves específicas de tipo
    got = errors(sub(MIN, 'type = "xlsx"', 'type = "xls"'))
    assert got == [("sources[0].type", "unsupported_source_type")]


def test_no_partial_configuration_returned():
    with pytest.raises(ConfigError):
        parse_config(sub(MIN, "max_file_mb = 10", "max_file_mb = 0"))


def test_c01_catalog_path_not_empty_but_not_checked_for_existence():
    cfg = parse_config(sub(MIN, 'duplicate_policy = "warn"', 'duplicate_policy = "warn"\ncatalog_path = "no/existe.csv"'))
    assert cfg.run.catalog_path == "no/existe.csv"
    assert ("run.catalog_path", "empty_value") in errors(sub(MIN, 'duplicate_policy = "warn"', 'duplicate_policy = "warn"\ncatalog_path = ""'))


def test_source_files_and_sheets_not_checked():
    # F2 no toca el sistema de archivos de datos: match/sheet inexistentes son válidos
    assert parse_config(MIN).sources[0].match == "a_*.xlsx"


# ---------------------------------------------------------------------------------------------------
# T3-16 (b): [sample] (F3, DF3-1)
# ---------------------------------------------------------------------------------------------------


def with_sample(sample: str) -> str:
    return MIN + sample


def test_sample_valid_and_backward_compatible():
    assert parse_config(MIN).sample is None  # una configuración v1 válida antes de F3 sigue siéndolo
    for seed in (0, 1, 20261001, 2**63 - 1):
        cfg = parse_config(with_sample(f"[sample]\nseed = {seed}\n"))
        assert cfg.sample == SampleSettings(seed) and cfg.sample.seed == seed and type(cfg.sample.seed) is int
        assert cfg.config_version == 1
    demo = load_config(DEMO)
    assert demo.sample == SampleSettings(20261001)
    assert [f.name for f in dataclasses.fields(RunSettings)] == [
        "period_start", "period_end", "null_tokens", "total_tolerance", "duplicate_policy", "catalog_path"]


def test_sample_immutability():
    cfg = parse_config(with_sample("[sample]\nseed = 7\n"))
    with pytest.raises(dataclasses.FrozenInstanceError):
        cfg.sample.seed = 8  # type: ignore[misc]
    assert SampleSettings.__dataclass_params__.frozen and "__slots__" in vars(SampleSettings)
    assert not hasattr(cfg.sample, "__dict__")


@pytest.mark.parametrize(
    "text,path,reason",
    [
        ("sample = 5\n" + MIN, "sample", "invalid_type"),
        (with_sample("[sample]\n"), "sample.seed", "missing_key"),
        (with_sample("[sample]\nseed = 1\nextra = 2\n"), "sample.extra", "unknown_key"),
        (with_sample("[sample]\nseed = true\n"), "sample.seed", "invalid_type"),
        (with_sample("[sample]\nseed = 1.5\n"), "sample.seed", "invalid_type"),
        (with_sample('[sample]\nseed = "1"\n'), "sample.seed", "invalid_type"),
        (with_sample("[sample]\nseed = -1\n"), "sample.seed", "seed_out_of_range"),
        (with_sample(f"[sample]\nseed = {2**63}\n"), "sample.seed", "seed_out_of_range"),
    ],
)
def test_sample_rejections(text, path, reason):
    assert (path, reason) in errors(text)


def test_sample_aggregation_order():
    text = build(
        limits=sub(LIMITS, "max_sheets = 10", "max_sheets = 0"),
        sources=(sub(XLSX, "header_row = 1", "header_row = 0"),),
    ) + "[sample]\nseed = -5\n"
    assert errors(text) == [
        ("limits.max_sheets", "non_positive"),
        ("sample.seed", "seed_out_of_range"),
        ("sources[0].header_row", "non_positive"),
    ]
    assert errors("sample = 1\nconfig_version = 2\n" + MIN.split("\n", 1)[1])[:2] == [
        ("config_version", "unsupported_version"), ("sample", "invalid_type")]
