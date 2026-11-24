"""Constantes, enumeraciones, especificaciones de incidencias e ``Issue`` (F2).

Solo biblioteca estándar. No contiene lógica de reglas de negocio (R01–R14): eso es de F6.
``Record`` (registro canónico) se aplaza a F4 (PO-6).

Fuentes: PHASE0_REPORT Rev. 3.2 §7.1, §7.2, §7.4, §8.1 y §8.2; PO-2 y PO-4 del registro de decisiones.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from enum import StrEnum
from types import MappingProxyType
from typing import TYPE_CHECKING, Mapping

if TYPE_CHECKING:  # solo tipos: config.py importa este módulo
    from .config import SourceProfile


class IssueCode(StrEnum):
    """Los 26 códigos de §8.1 (C/F) y §8.2 (R), en el orden de la especificación."""

    C01_INVALID_CONFIG = "C01_INVALID_CONFIG"
    F01_UNREADABLE_FILE = "F01_UNREADABLE_FILE"
    F02_ENCODING_ERROR = "F02_ENCODING_ERROR"
    F03_UNSUPPORTED_TYPE = "F03_UNSUPPORTED_TYPE"
    F04_LIMIT_EXCEEDED = "F04_LIMIT_EXCEEDED"
    F05_NO_PROFILE = "F05_NO_PROFILE"
    F06_AMBIGUOUS_PROFILE = "F06_AMBIGUOUS_PROFILE"
    F07_SHEET_NOT_FOUND = "F07_SHEET_NOT_FOUND"
    F08_REQUIRED_COLUMN_MISSING = "F08_REQUIRED_COLUMN_MISSING"
    F09_COLUMN_COLLISION = "F09_COLUMN_COLLISION"
    F10_NO_DATA_ROWS = "F10_NO_DATA_ROWS"
    F11_UNMAPPED_COLUMN = "F11_UNMAPPED_COLUMN"
    R01_MISSING_REQUIRED = "R01_MISSING_REQUIRED"
    R02_INVALID_DATE = "R02_INVALID_DATE"
    R03_AMBIGUOUS_DATE = "R03_AMBIGUOUS_DATE"
    R04_INVALID_NUMBER = "R04_INVALID_NUMBER"
    R05_NON_INTEGER_QUANTITY = "R05_NON_INTEGER_QUANTITY"
    R06_ZERO_QUANTITY = "R06_ZERO_QUANTITY"
    R07_NEGATIVE_QUANTITY = "R07_NEGATIVE_QUANTITY"
    R08_ZERO_PRICE = "R08_ZERO_PRICE"
    R09_NEGATIVE_PRICE = "R09_NEGATIVE_PRICE"
    R10_DATE_OUT_OF_PERIOD = "R10_DATE_OUT_OF_PERIOD"
    R11_TOTAL_MISMATCH = "R11_TOTAL_MISMATCH"
    R12_UNKNOWN_SKU = "R12_UNKNOWN_SKU"
    R13_DUPLICATE_SUSPECT = "R13_DUPLICATE_SUSPECT"
    R14_BRANCH_CONFLICT = "R14_BRANCH_CONFLICT"


class Severity(StrEnum):
    """§8: ``error`` → observed, ``scope`` → out_of_scope, ``warning`` e ``info`` no cambian el estado."""

    ERROR = "error"
    SCOPE = "scope"
    WARNING = "warning"
    INFO = "info"


class RecordStatus(StrEnum):
    """§7.1: estados de registro (definición única)."""

    VALID = "valid"
    OBSERVED = "observed"
    OUT_OF_SCOPE = "out_of_scope"


class KpiEffect(StrEnum):
    """§7.4: efecto de un registro en los KPIs."""

    INCLUDED = "included"
    INCLUDED_WARNING = "included_warning"
    EXCLUDED_OBSERVED = "excluded_observed"
    EXCLUDED_OUT_OF_SCOPE = "excluded_out_of_scope"


class IssueLevel(StrEnum):
    """§7.4: nivel de una incidencia."""

    FILE = "file"
    RECORD = "record"


STATUS_LABELS_ES: Mapping[RecordStatus, str] = MappingProxyType(
    {
        RecordStatus.VALID: "válido",
        RecordStatus.OBSERVED: "observado",
        RecordStatus.OUT_OF_SCOPE: "fuera de alcance",
    }
)

# Solo para incidencias de REGISTRO (nivel ``record``). Un error de archivo (F) rechaza el archivo y nunca
# es un estado de registro; C01 aborta la ejecución. ``warning`` e ``info`` no cambian el estado (§7.1),
# por eso no figuran aquí.
SEVERITY_TO_STATUS: Mapping[Severity, RecordStatus] = MappingProxyType(
    {
        Severity.ERROR: RecordStatus.OBSERVED,
        Severity.SCOPE: RecordStatus.OUT_OF_SCOPE,
    }
)


@dataclass(frozen=True, slots=True)
class IssueSpec:
    """Severidad de un código. ``pending_decision`` es metadato provisional (PO-4): NO activa ninguna regla.

    P-01, P-02 y P-03 se aprobaron el 2026-11-20; ningún código conserva hoy una decisión pendiente.
    """

    code: IssueCode
    severity: Severity
    pending_decision: str | None = None


_E, _S, _W, _I = Severity.ERROR, Severity.SCOPE, Severity.WARNING, Severity.INFO

# (código, severidad, decisión pendiente) en el orden de §8.1 y §8.2.
_SPEC_ROWS: tuple[tuple[IssueCode, Severity, str | None], ...] = (
    (IssueCode.C01_INVALID_CONFIG, _E, None),
    (IssueCode.F01_UNREADABLE_FILE, _E, None),
    (IssueCode.F02_ENCODING_ERROR, _E, None),
    (IssueCode.F03_UNSUPPORTED_TYPE, _E, None),
    (IssueCode.F04_LIMIT_EXCEEDED, _E, None),
    (IssueCode.F05_NO_PROFILE, _E, None),
    (IssueCode.F06_AMBIGUOUS_PROFILE, _E, None),
    (IssueCode.F07_SHEET_NOT_FOUND, _E, None),
    (IssueCode.F08_REQUIRED_COLUMN_MISSING, _E, None),
    (IssueCode.F09_COLUMN_COLLISION, _E, None),
    (IssueCode.F10_NO_DATA_ROWS, _W, None),
    (IssueCode.F11_UNMAPPED_COLUMN, _I, None),
    (IssueCode.R01_MISSING_REQUIRED, _E, None),
    (IssueCode.R02_INVALID_DATE, _E, None),
    (IssueCode.R03_AMBIGUOUS_DATE, _E, None),
    (IssueCode.R04_INVALID_NUMBER, _E, None),
    (IssueCode.R05_NON_INTEGER_QUANTITY, _E, None),
    (IssueCode.R06_ZERO_QUANTITY, _E, None),  # P-01 aprobado (2026-11-20)
    (IssueCode.R07_NEGATIVE_QUANTITY, _S, None),
    (IssueCode.R08_ZERO_PRICE, _W, None),
    (IssueCode.R09_NEGATIVE_PRICE, _E, None),
    (IssueCode.R10_DATE_OUT_OF_PERIOD, _S, None),  # P-02 aprobado (2026-11-20)
    (IssueCode.R11_TOTAL_MISMATCH, _E, None),  # P-03 aprobado (2026-11-20)
    (IssueCode.R12_UNKNOWN_SKU, _W, None),
    (IssueCode.R13_DUPLICATE_SUSPECT, _W, None),
    (IssueCode.R14_BRANCH_CONFLICT, _E, None),
)

ISSUE_SPECS: Mapping[IssueCode, IssueSpec] = MappingProxyType(
    {code: IssueSpec(code, severity, pending) for code, severity, pending in _SPEC_ROWS}
)

# §7.2: campos canónicos en el orden de la especificación.
CANONICAL_FIELDS: tuple[str, ...] = (
    "branch_id",
    "transaction_id",
    "sale_date",
    "product_sku",
    "product_name",
    "quantity",
    "unit_price",
    "reported_total",
    "branch_in_file",
    "line_amount",
)

# Obligatorios del esquema (§7.2).
REQUIRED_FIELDS: tuple[str, ...] = (
    "branch_id",
    "transaction_id",
    "sale_date",
    "product_sku",
    "quantity",
    "unit_price",
)

# Los que un perfil puede mapear: excluye ``branch_id`` (viene del perfil, D17) y ``line_amount`` (derivado).
MAPPABLE_FIELDS: tuple[str, ...] = (
    "transaction_id",
    "sale_date",
    "product_sku",
    "product_name",
    "quantity",
    "unit_price",
    "reported_total",
    "branch_in_file",
)

# Obligatorios que deben figurar en ``columns`` (PO-3): ``REQUIRED_FIELDS`` sin ``branch_id``.
REQUIRED_MAPPED_FIELDS: tuple[str, ...] = (
    "transaction_id",
    "sale_date",
    "product_sku",
    "quantity",
    "unit_price",
)


_NON_ALNUM_RUN = re.compile(r"[\W_]+")


def canonical_text(text: str) -> str:
    """Normalización canónica compartida (PO-2).

    1. Unicode NFKD. 2. Eliminar marcas diacríticas combinantes. 3. ``casefold()``.
    4. Sustituir secuencias de caracteres no alfanuméricos por un espacio. 5. Colapsar espacios y recortar.

    En F2 solo sirve para detectar colisiones de alias dentro de un perfil. El mapeo de encabezados (F5)
    debe reutilizar esta misma función; no debe existir un segundo algoritmo.
    """
    decomposed = unicodedata.normalize("NFKD", text)
    without_marks = "".join(ch for ch in decomposed if not unicodedata.combining(ch))
    return _NON_ALNUM_RUN.sub(" ", without_marks.casefold()).strip()


@dataclass(frozen=True, slots=True, kw_only=True)
class Issue:
    """Incidencia de ``data_issues.csv``: los 14 campos de §7.4, en ese orden.

    Los campos ausentes (celdas vacías del CSV) son ``None`` o la tupla vacía.
    C01 no es una ``Issue``: es un error de configuración que aborta la ejecución (``ConfigError``).
    """

    level: IssueLevel
    record_id: str | None = None
    source_id: str | None = None
    source_file: str | None = None
    source_sheet: str | None = None
    source_row: int | None = None
    field: str | None = None
    issue_code: IssueCode
    detail_code: str | None = None
    severity: Severity
    resulting_status: RecordStatus | None = None
    raw_value: str | None = None
    message: str
    related_record_ids: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        if not isinstance(self.level, IssueLevel):
            raise TypeError("level debe ser un IssueLevel")
        if not isinstance(self.issue_code, IssueCode):
            raise TypeError("issue_code debe ser un IssueCode")
        if not isinstance(self.severity, Severity):
            raise TypeError("severity debe ser un Severity")
        if self.resulting_status is not None and not isinstance(self.resulting_status, RecordStatus):
            raise TypeError("resulting_status debe ser un RecordStatus o None")
        if not isinstance(self.related_record_ids, tuple):
            raise TypeError("related_record_ids debe ser una tupla")
        if self.issue_code is IssueCode.C01_INVALID_CONFIG:
            raise ValueError("C01 aborta la ejecución (ConfigError); no es una incidencia")
        expected = ISSUE_SPECS[self.issue_code].severity
        if expected is not self.severity:
            raise ValueError(f"{self.issue_code} tiene severidad {expected}, no {self.severity}")
        if self.issue_code.value.startswith("F") != (self.level is IssueLevel.FILE):
            raise ValueError("los códigos F son de nivel file y los R de nivel record")
        if self.level is IssueLevel.FILE and (self.record_id is not None or self.resulting_status is not None):
            raise ValueError("una incidencia de nivel file no lleva record_id ni resulting_status")
        if self.detail_code is not None and self.issue_code not in (
            IssueCode.F01_UNREADABLE_FILE,
            IssueCode.F04_LIMIT_EXCEEDED,
        ):
            raise ValueError("detail_code solo se admite con F01 o F04")


# ---------------------------------------------------------------------------------------------------
# F4 (DT4-10, DT4-04): modelos de salida de la ingesta y catálogo cerrado de ``detail_code``
# ---------------------------------------------------------------------------------------------------

# 14 de F04 + 13 de F01 (DT4-10 D) + ``not_regular_file`` e ``io_error`` (DT4-04): 29 identificadores.
DETAIL_CODES_F04: frozenset[str] = frozenset(
    {
        "file_size", "zip_entries", "declared_size", "declared_ratio", "entry_uncompressed",
        "total_uncompressed", "shared_strings_count", "shared_strings_chars", "sheets", "columns",
        "cells", "text_chars", "records_file", "records_run",
    }
)
DETAIL_CODES_F01: frozenset[str] = frozenset(
    {
        "zip_unreadable", "unsafe_name", "duplicate_name", "encrypted_entry", "unsupported_method",
        "missing_part", "zip_truncated", "overlapping_entries", "crc_mismatch", "xml_doctype",
        "xml_encoding", "xml_malformed", "csv_malformed", "not_regular_file", "io_error",
    }
)
DETAIL_CODES: frozenset[str] = DETAIL_CODES_F04 | DETAIL_CODES_F01


@dataclass(frozen=True, slots=True)
class EnvironmentReport:
    python_version: str
    openpyxl_version: str
    expat_version: str
    expat_version_info: tuple[int, int, int]
    openpyxl_lxml: bool
    openpyxl_defusedxml: bool
    expat_min_ok: bool
    expat_reparse_deferral: bool


@dataclass(frozen=True, slots=True)
class SourceRecord:
    """Fila de origen con valores nativos y trazabilidad (§7.2). Sin normalizar (eso es de F5)."""

    record_id: str
    source_id: str
    source_file: str
    source_sheet: str  # "" en CSV
    source_row: int
    source_line: int | None  # línea física en CSV; None en XLSX
    values: tuple
    data_types: tuple[str, ...] | None = None  # XLSX: data_type de openpyxl; CSV: None


@dataclass(frozen=True, slots=True)
class IngestedFile:
    source_file: str
    profile: "SourceProfile"
    sha256: str
    sheet: str | None
    header: SourceRecord | None
    records: tuple[SourceRecord, ...]
    physical_rows: int
    preamble_rows: int
    header_rows: int
    blank_rows: int
    sheets_ignored: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class FileOutcome:
    source_file: str
    outcome: str  # "ingested" | "rejected" | "ignored"
    source_id: str | None = None
    issue: Issue | None = None
    ignore_reason: str | None = None
    ingested: IngestedFile | None = None


@dataclass(frozen=True, slots=True)
class ProcessingEvent:
    level: str  # "run" | "file" (DT4-10 C)
    event: str
    source_id: str | None = None
    source_file: str | None = None
    record_id: str | None = None
    field: str | None = None
    metric: str | None = None
    value: str | None = None
    message: str | None = None


@dataclass(frozen=True, slots=True)
class IngestResult:
    environment: EnvironmentReport
    files: tuple[FileOutcome, ...]
    events: tuple[ProcessingEvent, ...]
    issues: tuple[Issue, ...]
