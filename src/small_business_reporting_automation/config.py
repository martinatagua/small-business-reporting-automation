"""Carga y validación estricta de la configuración TOML (F2, C01).

Solo biblioteca estándar. No comprueba existencia de archivos fuente, hojas, rutas relativas ni globs
(fases posteriores). No implementa el mapeo de encabezados (F5). No hay valores por defecto: todas las
claves salvo las opcionales documentadas deben figurar en el archivo (P-12 sigue pendiente para ``[limits]``).

Un problema de lectura, UTF-8 o sintaxis produce un ``ConfigError`` con un único problema. Si el TOML
parsea, se acumulan todos los problemas de estructura y validación en orden determinista (raíz → ``run`` →
``limits`` → ``sample`` → perfiles → comprobaciones entre perfiles) y nunca se devuelve una configuración parcial.

``config_version = 1``: la tabla ``[sample]`` (F3, DF3-1) es una ampliación ADITIVA y opcional. Toda
configuración v1 sin ``[sample]`` sigue siendo válida y produce ``AppConfig.sample is None``. Un archivo con
``[sample]`` lo rechaza el código anterior a F3 (clave desconocida).
"""

from __future__ import annotations

import codecs
import re
import tomllib
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from os import PathLike
from pathlib import Path
from types import MappingProxyType
from typing import Any, Mapping

from .models import MAPPABLE_FIELDS, REQUIRED_MAPPED_FIELDS, IssueCode, canonical_text

# ---------------------------------------------------------------------------------------------------
# Estructuras públicas
# ---------------------------------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class ConfigProblem:
    """Un problema de configuración. ``reason`` es un identificador interno estable (no un ``detail_code``)."""

    path: str
    reason: str
    message: str


class ConfigError(Exception):
    """C01: configuración inválida. Aborta la ejecución (salida 2 en la CLI, F9)."""

    code = IssueCode.C01_INVALID_CONFIG

    def __init__(self, problems: tuple[ConfigProblem, ...]) -> None:
        self.problems = tuple(problems)
        lines = [f"Configuración inválida ({IssueCode.C01_INVALID_CONFIG}):"]
        lines += [f"  - {p.path or '(raíz)'}: {p.message} [{p.reason}]" for p in self.problems]
        super().__init__("\n".join(lines))


@dataclass(frozen=True, slots=True)
class RunSettings:
    period_start: date
    period_end: date
    null_tokens: tuple[str, ...]
    total_tolerance: Decimal
    duplicate_policy: str
    catalog_path: str | None


@dataclass(frozen=True, slots=True)
class Limits:
    max_file_mb: int
    max_zip_entries: int
    max_total_uncompressed_mb: int
    max_entry_uncompressed_mb: int
    max_declared_compression_ratio: int
    max_shared_strings: int
    max_shared_strings_chars: int
    max_sheets: int
    max_columns: int
    max_cells_per_file: int
    max_text_chars: int
    max_records_per_file: int
    max_records_per_run: int


@dataclass(frozen=True, slots=True)
class SourceProfile:
    id: str
    branch_id: str
    match: str
    type: str
    header_row: int
    date_formats: tuple[str, ...]
    decimal_separator: str
    thousands_separator: str
    columns: Mapping[str, tuple[str, ...]]
    branch_aliases: tuple[str, ...]
    currency_symbols: tuple[str, ...]
    # Solo XLSX (``None`` en CSV)
    sheet: str | None
    accept_native_dates: bool | None
    accept_serial_dates: bool | None
    # Solo CSV (``None`` en XLSX)
    encoding: str | None
    delimiter: str | None


@dataclass(frozen=True, slots=True)
class SampleSettings:
    """``[sample]`` (DF3-1): semilla por defecto de ``sb-report generate-sample``."""

    seed: int


@dataclass(frozen=True, slots=True)
class AppConfig:
    config_version: int
    run: RunSettings
    limits: Limits
    sources: tuple[SourceProfile, ...]
    sample: SampleSettings | None = None


# ---------------------------------------------------------------------------------------------------
# Esquema
# ---------------------------------------------------------------------------------------------------

_ROOT_KEYS = ("config_version", "run", "limits", "sources")
_ROOT_OPTIONAL = ("sample",)
_SAMPLE_KEYS = ("seed",)
_SEED_MAX = 2**63 - 1  # E-7; tomllib no limita el tamaño de los enteros
_RUN_KEYS = ("period_start", "period_end", "null_tokens", "total_tolerance", "duplicate_policy")
_RUN_OPTIONAL = ("catalog_path",)
_LIMIT_KEYS = (
    "max_file_mb",
    "max_zip_entries",
    "max_total_uncompressed_mb",
    "max_entry_uncompressed_mb",
    "max_declared_compression_ratio",
    "max_shared_strings",
    "max_shared_strings_chars",
    "max_sheets",
    "max_columns",
    "max_cells_per_file",
    "max_text_chars",
    "max_records_per_file",
    "max_records_per_run",
)
_COMMON_KEYS = (
    "id",
    "branch_id",
    "match",
    "type",
    "header_row",
    "date_formats",
    "decimal_separator",
    "thousands_separator",
    "columns",
)
_COMMON_OPTIONAL = ("branch_aliases", "currency_symbols")
_XLSX_KEYS = ("sheet", "accept_native_dates", "accept_serial_dates")
_CSV_KEYS = ("encoding", "delimiter")
_SOURCE_TYPES = ("xlsx", "csv")
# §7.3 / §11.3 (G1–G3): únicas combinaciones (decimal, miles) con gramática definida.
_SEPARATOR_COMBINATIONS = ((".", ""), (",", ""), (",", "."))
_TOLERANCE_RE = re.compile(r"[+-]?[0-9]+(\.[0-9]+)?")
_DATE_DIRECTIVES = ("%d", "%m", "%Y", "%y", "%%")
_REFERENCE_DATES = (date(2026, 3, 14), date(1999, 12, 31))

_TYPE_NAMES = {str: "cadena", int: "entero", bool: "booleano", date: "fecha", list: "lista", dict: "tabla"}


def _join(parent: str, key: str) -> str:
    return f"{parent}.{key}" if parent else key


class _Problems:
    def __init__(self) -> None:
        self.items: list[ConfigProblem] = []

    def add(self, path: str, reason: str, message: str) -> None:
        self.items.append(ConfigProblem(path, reason, message))

    def __bool__(self) -> bool:
        return bool(self.items)


def _is(value: Any, kind: type) -> bool:
    # Tipo exacto: ``bool`` es subclase de ``int`` y ``datetime`` lo es de ``date``.
    return type(value) is kind


def _expect(p: _Problems, path: str, value: Any, kind: type) -> bool:
    if _is(value, kind):
        return True
    p.add(path, "invalid_type", f"debe ser de tipo {_TYPE_NAMES[kind]}, no {type(value).__name__}")
    return False


def _non_empty_str(p: _Problems, path: str, value: Any) -> bool:
    if not _expect(p, path, value, str):
        return False
    if not value.strip():
        p.add(path, "empty_value", "no puede estar vacía ni contener solo espacios")
        return False
    return True


def _check_keys(
    p: _Problems, table: Mapping[str, Any], path: str, required: tuple[str, ...], allowed: tuple[str, ...]
) -> None:
    for key in table:
        if key not in allowed:
            p.add(_join(path, key), "unknown_key", f"clave desconocida {key!r}")
    for key in required:
        if key not in table:
            p.add(_join(path, key), "missing_key", f"falta la clave obligatoria {key!r}")


def _str_list(p: _Problems, path: str, value: Any, *, allow_empty_items: bool, allow_empty_list: bool) -> bool:
    if not _expect(p, path, value, list):
        return False
    ok = True
    if not value and not allow_empty_list:
        p.add(path, "empty_value", "la lista no puede estar vacía")
        ok = False
    for i, item in enumerate(value):
        item_path = f"{path}[{i}]"
        if not _expect(p, item_path, item, str):
            ok = False
        elif not allow_empty_items and not item.strip():
            p.add(item_path, "empty_value", "no puede estar vacía ni contener solo espacios")
            ok = False
    return ok


# ---------------------------------------------------------------------------------------------------
# Validadores
# ---------------------------------------------------------------------------------------------------


def _check_date_format(fmt: str) -> str | None:
    """Devuelve un motivo de rechazo en español, o ``None`` si el formato es válido (PO-1)."""
    counts = {"d": 0, "m": 0, "y": 0}
    i = 0
    while i < len(fmt):
        if fmt[i] != "%":
            i += 1
            continue
        directive = fmt[i : i + 2]
        if directive not in _DATE_DIRECTIVES:
            return f"directiva no admitida {directive!r}; solo %d, %m, %Y, %y y %%"
        if directive == "%d":
            counts["d"] += 1
        elif directive == "%m":
            counts["m"] += 1
        elif directive in ("%Y", "%y"):
            counts["y"] += 1
        i += 2
    for name, label in (("d", "día"), ("m", "mes"), ("y", "año")):
        if counts[name] != 1:
            return f"debe contener exactamente un {label} (encontrados: {counts[name]})"
    for ref in _REFERENCE_DATES:
        try:
            if datetime.strptime(ref.strftime(fmt), fmt).date() != ref:
                return "no es reversible (ida y vuelta) con fechas de referencia"
        except ValueError:
            return "no es reversible (ida y vuelta) con fechas de referencia"
    return None


def _check_run(p: _Problems, run: Any, path: str = "run") -> dict[str, Any] | None:
    if not _expect(p, path, run, dict):
        return None
    _check_keys(p, run, path, _RUN_KEYS, _RUN_KEYS + _RUN_OPTIONAL)
    before = len(p.items)
    vals: dict[str, Any] = {}
    dates_ok = True
    for key in ("period_start", "period_end"):
        if key in run:
            if _expect(p, _join(path, key), run[key], date):
                vals[key] = run[key]
            else:
                dates_ok = False
        else:
            dates_ok = False
    if dates_ok and vals["period_start"] > vals["period_end"]:
        p.add(_join(path, "period_start"), "invalid_period", "period_start no puede ser posterior a period_end")
    if "null_tokens" in run:
        npath = _join(path, "null_tokens")
        if _str_list(p, npath, run["null_tokens"], allow_empty_items=True, allow_empty_list=True):
            vals["null_tokens"] = tuple(run["null_tokens"])
    if "total_tolerance" in run:
        tpath = _join(path, "total_tolerance")
        tol = run["total_tolerance"]
        if not _is(tol, str):
            p.add(tpath, "invalid_type", f"debe ser una cadena (nunca un número TOML), no {type(tol).__name__}")
        elif not _TOLERANCE_RE.fullmatch(tol):
            p.add(tpath, "invalid_tolerance", "cadena decimal no admitida; formato [+-]dígitos[.dígitos]")
        else:
            value = Decimal(tol)  # directamente desde la cadena, sin float
            if value < 0:
                p.add(tpath, "negative_tolerance", "la tolerancia no puede ser negativa")
            else:
                vals["total_tolerance"] = value
    if "duplicate_policy" in run:
        dpath = _join(path, "duplicate_policy")
        if _expect(p, dpath, run["duplicate_policy"], str):
            if run["duplicate_policy"] != "warn":
                p.add(dpath, "unsupported_policy", "el único valor admitido en el MVP es 'warn'")
            else:
                vals["duplicate_policy"] = "warn"
    vals["catalog_path"] = None
    if "catalog_path" in run:
        if _non_empty_str(p, _join(path, "catalog_path"), run["catalog_path"]):
            vals["catalog_path"] = run["catalog_path"]
    return vals if len(p.items) == before else None


def _check_limits(p: _Problems, limits: Any, path: str = "limits") -> dict[str, int] | None:
    if not _expect(p, path, limits, dict):
        return None
    _check_keys(p, limits, path, _LIMIT_KEYS, _LIMIT_KEYS)
    before = len(p.items)
    vals: dict[str, int] = {}
    for key in _LIMIT_KEYS:
        if key not in limits:
            continue
        kpath = _join(path, key)
        if not _expect(p, kpath, limits[key], int):
            continue
        if limits[key] <= 0:
            p.add(kpath, "non_positive", "debe ser un entero estrictamente positivo")
            continue
        vals[key] = limits[key]
    for small, large in (
        ("max_entry_uncompressed_mb", "max_total_uncompressed_mb"),
        ("max_records_per_file", "max_records_per_run"),
    ):
        if small in vals and large in vals and vals[small] > vals[large]:
            p.add(_join(path, small), "limit_relation", f"{small} no puede ser mayor que {large}")
    return vals if len(p.items) == before else None


def _check_sample(p: _Problems, sample: Any, path: str = "sample") -> dict[str, int] | None:
    if not _expect(p, path, sample, dict):
        return None
    _check_keys(p, sample, path, _SAMPLE_KEYS, _SAMPLE_KEYS)
    before = len(p.items)
    vals: dict[str, int] = {}
    if "seed" in sample:
        spath = _join(path, "seed")
        if _expect(p, spath, sample["seed"], int):  # tipo exacto: ``bool`` no es una semilla
            if not 0 <= sample["seed"] <= _SEED_MAX:
                p.add(spath, "seed_out_of_range", "la semilla debe estar en [0, 2^63 - 1]")
            else:
                vals["seed"] = sample["seed"]
    return vals if len(p.items) == before else None


def _check_columns(p: _Problems, columns: Any, path: str) -> dict[str, tuple[str, ...]] | None:
    if not _expect(p, path, columns, dict):
        return None
    _check_keys(p, columns, path, REQUIRED_MAPPED_FIELDS, MAPPABLE_FIELDS)
    before = len(p.items)
    vals: dict[str, tuple[str, ...]] = {}
    canon_owner: dict[str, str] = {}
    reported: set[tuple[str, str]] = set()
    for field in MAPPABLE_FIELDS:  # orden determinista
        if field not in columns:
            continue
        fpath = _join(path, field)
        if not _str_list(p, fpath, columns[field], allow_empty_items=False, allow_empty_list=False):
            continue
        vals[field] = tuple(columns[field])
        for alias in columns[field]:
            canon = canonical_text(alias)
            if not canon:
                p.add(fpath, "empty_alias", f"el alias {alias!r} queda vacío tras la normalización canónica")
                continue
            owner = canon_owner.setdefault(canon, field)
            if owner != field and (field, canon) not in reported:
                reported.add((field, canon))
                p.add(
                    fpath,
                    "alias_collision",
                    f"el alias {alias!r} (canónico {canon!r}) ya está asignado al campo {owner!r} en este perfil",
                )
    return vals if len(p.items) == before else None


def _check_source(p: _Problems, src: Any, path: str) -> dict[str, Any] | None:
    if not _expect(p, path, src, dict):
        return None
    stype = src.get("type") if _is(src.get("type"), str) and src.get("type") in _SOURCE_TYPES else None
    allowed = _COMMON_KEYS + _COMMON_OPTIONAL
    if stype == "xlsx":
        required, allowed = _COMMON_KEYS + _XLSX_KEYS, allowed + _XLSX_KEYS
        other = _CSV_KEYS
    elif stype == "csv":
        required, allowed = _COMMON_KEYS + _CSV_KEYS, allowed + _CSV_KEYS
        other = _XLSX_KEYS
    else:
        required, allowed, other = _COMMON_KEYS, allowed + _XLSX_KEYS + _CSV_KEYS, ()
    for key in src:
        if key in other:
            p.add(_join(path, key), "key_not_allowed_for_type", f"la clave {key!r} no aplica a perfiles de tipo {stype!r}")
    _check_keys(p, {k: v for k, v in src.items() if k not in other}, path, required, allowed)
    before = len(p.items)
    vals: dict[str, Any] = {
        "branch_aliases": (),
        "currency_symbols": (),
        "sheet": None,
        "accept_native_dates": None,
        "accept_serial_dates": None,
        "encoding": None,
        "delimiter": None,
    }

    def text(key: str) -> None:
        if key in src and _non_empty_str(p, _join(path, key), src[key]):
            vals[key] = src[key]

    for key in ("id", "branch_id", "match"):
        text(key)
    if "type" in src:
        tpath = _join(path, "type")
        if _expect(p, tpath, src["type"], str):
            if src["type"] not in _SOURCE_TYPES:
                p.add(tpath, "unsupported_source_type", "debe ser exactamente 'xlsx' o 'csv'")
            else:
                vals["type"] = src["type"]
    if "header_row" in src:
        hpath = _join(path, "header_row")
        if _expect(p, hpath, src["header_row"], int):
            if src["header_row"] < 1:
                p.add(hpath, "non_positive", "header_row empieza en 1")
            else:
                vals["header_row"] = src["header_row"]
    if "date_formats" in src:
        dpath = _join(path, "date_formats")
        if _str_list(p, dpath, src["date_formats"], allow_empty_items=True, allow_empty_list=False):
            ok = True
            for i, fmt in enumerate(src["date_formats"]):
                problem = _check_date_format(fmt)
                if problem:
                    ok = False
                    p.add(f"{dpath}[{i}]", "invalid_date_format", f"formato de fecha {fmt!r} inválido: {problem}")
            if ok:
                vals["date_formats"] = tuple(src["date_formats"])
    seps_ok = all(
        _expect(p, _join(path, k), src[k], str) for k in ("decimal_separator", "thousands_separator") if k in src
    ) and "decimal_separator" in src and "thousands_separator" in src
    if seps_ok:
        dec, thou = src["decimal_separator"], src["thousands_separator"]
        if dec == thou:
            p.add(_join(path, "decimal_separator"), "separator_conflict", "decimal_separator y thousands_separator no pueden ser iguales")
        elif (dec, thou) not in _SEPARATOR_COMBINATIONS:
            p.add(
                _join(path, "decimal_separator"),
                "unsupported_separators",
                f"combinación ({dec!r}, {thou!r}) no admitida; solo ('.', ''), (',', '') y (',', '.')",
            )
        else:
            vals["decimal_separator"], vals["thousands_separator"] = dec, thou
    for key in ("branch_aliases", "currency_symbols"):
        if key in src:
            kpath = _join(path, key)
            if _str_list(p, kpath, src[key], allow_empty_items=False, allow_empty_list=True):
                vals[key] = tuple(src[key])
    if "columns" in src:
        cols = _check_columns(p, src["columns"], _join(path, "columns"))
        if cols is not None:
            vals["columns"] = MappingProxyType(cols)
    if stype == "xlsx":
        text("sheet")
        for key in ("accept_native_dates", "accept_serial_dates"):
            if key in src and _expect(p, _join(path, key), src[key], bool):
                vals[key] = src[key]
    elif stype == "csv":
        if "encoding" in src and _non_empty_str(p, _join(path, "encoding"), src["encoding"]):
            try:
                codecs.lookup(src["encoding"])
                vals["encoding"] = src["encoding"]
            except (LookupError, ValueError):
                p.add(_join(path, "encoding"), "invalid_encoding", f"encoding desconocido {src['encoding']!r}")
        if "delimiter" in src and _expect(p, _join(path, "delimiter"), src["delimiter"], str):
            d = src["delimiter"]
            if len(d) != 1 or d in '"\r\n':
                p.add(_join(path, "delimiter"), "invalid_delimiter", "debe ser un único carácter distinto de comillas y saltos de línea")
            else:
                vals["delimiter"] = d
    return vals if len(p.items) == before else None


def _validate(data: dict[str, Any]) -> AppConfig:
    p = _Problems()
    _check_keys(p, data, "", _ROOT_KEYS, _ROOT_KEYS + _ROOT_OPTIONAL)
    version_ok = False
    if "config_version" in data and _expect(p, "config_version", data["config_version"], int):
        if data["config_version"] != 1:
            p.add("config_version", "unsupported_version", "config_version debe ser 1")
        else:
            version_ok = True
    run = _check_run(p, data["run"]) if "run" in data else None
    limits = _check_limits(p, data["limits"]) if "limits" in data else None
    sample = _check_sample(p, data["sample"]) if "sample" in data else None
    sources: list[dict[str, Any]] = []
    sources_ok = False
    if "sources" in data:
        raw = data["sources"]
        if _expect(p, "sources", raw, list):
            if not raw:
                p.add("sources", "no_sources", "debe haber al menos un perfil en sources")
            else:
                sources_ok = True
                for i, src in enumerate(raw):
                    result = _check_source(p, src, f"sources[{i}]")
                    if result is None:
                        sources_ok = False
                    else:
                        sources.append(result)
                seen: dict[str, int] = {}
                for i, src in enumerate(raw):
                    sid = src.get("id") if isinstance(src, dict) else None
                    if _is(sid, str) and sid.strip():
                        if sid in seen:
                            p.add(f"sources[{i}].id", "duplicate_id", f"id {sid!r} duplicado (ya usado en sources[{seen[sid]}])")
                            sources_ok = False
                        else:
                            seen[sid] = i
    if p or not (version_ok and run and limits and sources_ok):
        raise ConfigError(tuple(p.items))
    return AppConfig(
        config_version=1,
        run=RunSettings(**run),
        limits=Limits(**limits),
        sources=tuple(SourceProfile(**s) for s in sources),
        sample=SampleSettings(**sample) if sample is not None else None,
    )


# ---------------------------------------------------------------------------------------------------
# API pública
# ---------------------------------------------------------------------------------------------------


def parse_config(text: str) -> AppConfig:
    """Valida el texto TOML y devuelve la configuración, o lanza ``ConfigError`` (C01)."""
    try:
        data = tomllib.loads(text)
    except tomllib.TOMLDecodeError as exc:
        raise ConfigError((ConfigProblem("", "toml_syntax", f"TOML inválido: {exc}"),)) from exc
    return _validate(data)


def load_config(path: str | PathLike[str]) -> AppConfig:
    """Lee ``path`` (UTF-8 estricto) y devuelve la configuración validada, o lanza ``ConfigError`` (C01)."""
    shown = str(path)
    try:
        raw = Path(path).read_bytes()
    except FileNotFoundError as exc:
        raise ConfigError((ConfigProblem(shown, "file_not_found", "el archivo de configuración no existe"),)) from exc
    except OSError as exc:
        raise ConfigError((ConfigProblem(shown, "file_unreadable", f"no se pudo leer el archivo: {exc.strerror or exc}"),)) from exc
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise ConfigError((ConfigProblem(shown, "invalid_utf8", "el archivo no es UTF-8 válido"),)) from exc
    return parse_config(text)
