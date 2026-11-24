"""Generador determinista de datos sintéticos y ground truth (F3, Plan §4).

Independiente del pipeline (§10.3, control 1): solo biblioteca estándar y ``openpyxl``. No importa nada del
propio paquete: la CLI resuelve la semilla y llama a :func:`generate` con tipos primitivos.

El generador PRODUCE entradas con anomalías intencionales; NO evalúa reglas. Estado esperado, códigos y
``kpi_effect`` salen de la tabla literal :data:`SCENARIOS` (transcripción de §10.2) y los importes esperados
se calculan con ``Decimal`` sobre los valores del plan antes de darles formato.
"""

import csv
import hashlib
import io
import json
import os
import random
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import openpyxl

GENERATOR_VERSION = "1"
TARGET_CONFIG_VERSION = 1
PERIOD = (date(2026, 10, 1), date(2026, 10, 31))
TOLERANCE = Decimal("0.01")
SEED_MAX = 2**63 - 1

STAGING_NAME = ".generate-sample-staging"
SUBDIRS = ("raw", "reference", "ground_truth")
MANIFEST_REL = "ground_truth/manifest.json"

# Los 26 códigos de §8.1/§8.2 como literales propios (T3-03 los contrasta con ``models.IssueCode``).
ISSUE_CODES = (
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
)

# --------------------------------------------------------------------------------------------------
# Diseño de archivos (§10.1, Plan §5)
# --------------------------------------------------------------------------------------------------

FILE_KEYS = ("N", "C", "SX", "SC")

FILES = {
    "N": {
        "name": "norte_ventas_2026-10.xlsx", "source_id": "norte-xlsx", "branch": "NORTE", "kind": "xlsx",
        "sheet": "Ventas", "header_row": 3, "preamble": 2,
        "headers": ("Fecha", "Nº Ticket", "Código", "Producto", "Cantidad", "Precio Unitario", "Total"),
    },
    "C": {
        "name": "centro_ventas_2026-10.csv", "source_id": "centro-csv", "branch": "CENTRO", "kind": "csv",
        "sheet": "", "header_row": 1, "preamble": 0, "encoding": "cp1252", "delimiter": ";", "bom": False,
        "headers": ("ticket", "fecha_venta", "sku", "descripcion", "unidades", "precio", "importe", "Observaciones"),
    },
    "SX": {
        "name": "sur_ventas_2026-10_tienda.xlsx", "source_id": "sur-xlsx", "branch": "SUR", "kind": "xlsx",
        "sheet": "Datos", "header_row": 1, "preamble": 0,
        "headers": ("Date", "Receipt", "SKU", "Item", "Qty", "Unit Price", "Sucursal", "Cajero"),
    },
    "SC": {
        "name": "sur_ventas_2026-10_web.csv", "source_id": "sur-csv", "branch": "SUR", "kind": "csv",
        "sheet": "", "header_row": 1, "preamble": 0, "encoding": "utf-8-sig", "delimiter": ",", "bom": True,
        "headers": ("fecha", "pedido", "sku", "nombre", "cant", "pu"),
    },
}

# Orden canónico del manifiesto (Plan §7); también el orden de publicación de la fase 4(b).
MANIFEST_PATHS = (
    "raw/centro_ventas_2026-10.csv",
    "raw/norte_ventas_2026-10.xlsx",
    "raw/sur_ventas_2026-10_tienda.xlsx",
    "raw/sur_ventas_2026-10_web.csv",
    "reference/products.csv",
    "ground_truth/expected_records.csv",
    "ground_truth/expected_issues.csv",
    "ground_truth/expected_files.csv",
    "ground_truth/expected_summary.json",
)
ALL_PATHS = MANIFEST_PATHS + (MANIFEST_REL,)
# Terminador físico exigido por archivo CSV (Plan §5.1).
EOL = {
    "raw/centro_ventas_2026-10.csv": b"\r\n",
    "raw/sur_ventas_2026-10_web.csv": b"\r\n",
    "reference/products.csv": b"\n",
    "ground_truth/expected_records.csv": b"\n",
    "ground_truth/expected_issues.csv": b"\n",
    "ground_truth/expected_files.csv": b"\n",
}

# --------------------------------------------------------------------------------------------------
# Tabla literal de escenarios (transcripción de §10.2; Plan §3.1)
# (id, nombre, códigos, estado, kpi_effect, {archivo: cantidad})
# --------------------------------------------------------------------------------------------------

SCENARIOS = (
    ("S01", "Venta ordinaria", (), "valid", "included", {"N": 40, "C": 40, "SX": 25, "SC": 25}),
    ("S02", "Devolución", ("R07_NEGATIVE_QUANTITY",), "out_of_scope", "excluded_out_of_scope", {"N": 1, "C": 1, "SC": 1}),
    ("S03", "Cantidad cero", ("R06_ZERO_QUANTITY",), "observed", "excluded_observed", {"N": 1, "SX": 1}),
    ("S04", "Precio cero", ("R08_ZERO_PRICE",), "valid", "included_warning", {"N": 1, "C": 1}),
    ("S05", "Precio negativo", ("R09_NEGATIVE_PRICE",), "observed", "excluded_observed", {"C": 1, "SX": 1}),
    ("S06", "Campo obligatorio ausente", ("R01_MISSING_REQUIRED",), "observed", "excluded_observed",
     {"N": 1, "C": 1, "SX": 1, "SC": 1}),
    ("S07", "Fila Total general", ("R01_MISSING_REQUIRED",), "observed", "excluded_observed", {"N": 1}),
    ("S08", "Número no parseable", ("R04_INVALID_NUMBER",), "observed", "excluded_observed", {"N": 1, "C": 1, "SC": 1}),
    ("S09", "Cantidad no entera", ("R05_NON_INTEGER_QUANTITY",), "observed", "excluded_observed", {"C": 1, "SX": 1}),
    ("S10", "Fecha inválida", ("R02_INVALID_DATE",), "observed", "excluded_observed", {"N": 1, "C": 1, "SC": 1}),
    ("S11", "Fecha ambigua", ("R03_AMBIGUOUS_DATE",), "observed", "excluded_observed", {"SC": 2}),
    ("S12", "Fuera de período", ("R10_DATE_OUT_OF_PERIOD",), "out_of_scope", "excluded_out_of_scope",
     {"N": 1, "SX": 1, "SC": 1}),
    ("S13", "Total discrepante", ("R11_TOTAL_MISMATCH",), "observed", "excluded_observed", {"N": 1, "C": 1}),
    ("S14", "SKU desconocido", ("R12_UNKNOWN_SKU",), "valid", "included_warning", {"N": 1, "C": 1, "SX": 1}),
    ("S15", "Conflicto de sucursal", ("R14_BRANCH_CONFLICT",), "observed", "excluded_observed", {"SX": 1}),
    ("S16", "Duplicado sospechoso, mismo archivo", ("R13_DUPLICATE_SUSPECT",), "valid", "included_warning", {"N": 2}),
    ("S17", "Duplicado sospechoso, entre archivos", ("R13_DUPLICATE_SUSPECT",), "valid", "included_warning",
     {"SX": 1, "SC": 1}),
    ("S18", "Líneas idénticas legítimas", ("R13_DUPLICATE_SUSPECT",), "valid", "included_warning", {"C": 2}),
    ("S19", "Multifallo", ("R01_MISSING_REQUIRED", "R04_INVALID_NUMBER"), "observed", "excluded_observed", {"N": 1}),
    ("S20", "Fórmula peligrosa en product_name", (), "valid", "included", {"N": 2, "C": 2, "SX": 1}),
    ("S21", "Texto con tabulador y control", (), "valid", "included", {"SC": 2}),
    # No son registros: S22 = incidencia de archivo (F11); S23 = fila en blanco; S24 = atributo de S01.
    ("S22", "Encabezado desconocido", ("F11_UNMAPPED_COLUMN",), None, None, {"C": 1, "SX": 1}),
    ("S23", "Fila en blanco", (), None, None, {"N": 1, "SX": 1, "SC": 1}),
    ("S24", "Ruido benigno", (), None, None, {"N": 1, "C": 1, "SX": 1, "SC": 1}),
)
RECORD_SCENARIOS = tuple(s for s in SCENARIOS if s[3] is not None)
FILE_ISSUES = {"C": "F11_UNMAPPED_COLUMN", "SX": "F11_UNMAPPED_COLUMN"}  # columnas no mapeadas: C Observaciones, SX Cajero
BLANK_ROWS = {"N": 1, "C": 0, "SX": 1, "SC": 1}

# Catálogo ficticio: 20 productos genéricos. El precio base es interno y NO se exporta (DF3-6). Los precios
# no llevan ceros finales ni notación exponencial (Plan §6.6, condiciones de XLSX nativo).
PRODUCTS = (
    ("AB-101", "Café molido 250 g", Decimal("12.5")),
    ("AB-102", "Té verde caja", Decimal("7.25")),
    ("AB-103", "Azúcar 1 kg", Decimal("3")),
    ("AB-104", "Aceite de oliva 500 ml", Decimal("19.99")),
    ("AB-105", "Arroz 5 kg", Decimal("45")),
    ("CD-201", "Jabón neutro", Decimal("8.9")),
    ("CD-202", "Detergente 3 l", Decimal("120.75")),
    ("CD-203", "Papel higiénico x12", Decimal("5.5")),
    ("CD-204", "Cuaderno A5", Decimal("250")),
    ("CD-205", "Bolígrafo azul", Decimal("15.2")),
    ("EF-301", "Galletas surtidas", Decimal("9.99")),
    ("EF-302", "Jugo de naranja 1 l", Decimal("32.5")),
    ("EF-303", "Agua mineral 2 l", Decimal("1.75")),
    ("EF-304", "Pan de molde", Decimal("64")),
    ("EF-305", "Mermelada de fresa", Decimal("27.8")),
    ("GH-401", "Pilas AA x4", Decimal("11.1")),
    ("GH-402", "Bombilla LED", Decimal("99.5")),
    ("GH-403", "Cinta adhesiva", Decimal("4.25")),
    ("GH-404", "Bolsa reutilizable", Decimal("18")),
    ("GH-405", "Linterna de mano", Decimal("75.5")),
)
UNKNOWN_SKUS = ("ZZ-900", "ZZ-901", "ZZ-902")
CATALOG_SKUS = frozenset(p[0] for p in PRODUCTS)
D2 = Decimal("0.01")
SKU_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-"


class GenerateError(Exception):
    """Fallo de generación. ``kind``: ``argument``, ``prepare``, ``interrupted`` o ``cleanup`` (Plan §8.3)."""

    def __init__(self, kind: str, message: str) -> None:
        super().__init__(message)
        self.kind = kind
        self.message = message


# --------------------------------------------------------------------------------------------------
# Modelo del plan (Plan §4.3)
# --------------------------------------------------------------------------------------------------


@dataclass
class PlannedRow:
    file_key: str
    scenario_id: str
    ticket: str | None = None
    sale_date: date | None = None
    date_text: str | None = None       # texto literal de la celda de fecha (sustituye al valor nativo)
    sku: str = ""                      # SKU canónico
    sku_written: str = ""              # SKU escrito en la fuente
    name: str = ""
    qty: int | Decimal | None = None
    qty_text: str | None = None
    price: Decimal | None = None
    price_text: str | None = None
    total: Decimal | None = None       # reported_total (solo N y C)
    branch_in_file: str | None = None
    general_total: bool = False        # fila «Total general» (S07)
    noise: list = field(default_factory=list)
    source_row: int = 0
    source_line: int | None = None

    @property
    def spec(self):
        return next(s for s in RECORD_SCENARIOS if s[0] == self.scenario_id)

    @property
    def expected_status(self) -> str:
        return self.spec[3]

    @property
    def expected_codes(self) -> tuple:
        return tuple(sorted(self.spec[2]))

    @property
    def kpi_effect(self) -> str:
        return self.spec[4]

    @property
    def line_amount(self) -> Decimal | None:
        """Cantidad × precio con los ``Decimal`` del plan; vacío sin cantidad o precio canónicos (DF3-4)."""
        if type(self.qty) is int and self.price is not None and self.qty_text is None and self.price_text is None:
            return self.qty * self.price
        return None


@dataclass
class Plan:
    seed: int
    layout: dict                       # file_key -> lista de PlannedRow o None (fila en blanco)


def _file_rng(seed: int, source_id: str) -> random.Random:
    digest = hashlib.sha256(f"{seed}:{source_id}".encode()).digest()
    return random.Random(int.from_bytes(digest[:8], "big"))


def _price_for(fk: str, price: Decimal) -> Decimal:
    return price.quantize(D2) if fk == "C" else price


_COPY_FIELDS = ("file_key", "scenario_id", "ticket", "sale_date", "sku", "sku_written", "name", "qty", "price", "total")


def _build_file(fk: str, rng: random.Random) -> list:
    """Filas de un archivo en su orden de escritura (``None`` = fila en blanco). Único consumidor del RNG."""
    counter = {"n": 0, "p": 0}
    prefix = {"N": "NT-", "C": "C", "SX": "SX-", "SC": "WEB-"}[fk]
    width = {"N": 5, "C": 6, "SX": 5, "SC": 5}[fk]
    start = {"N": 10000, "C": 200000, "SX": 30000, "SC": 40000}[fk]
    pool = list(PRODUCTS)
    rng.shuffle(pool)
    day_min = 13 if fk == "SC" else 1

    def next_ticket() -> str:
        counter["n"] += 1
        return f"{prefix}{start + counter['n']:0{width}d}"

    def next_product():
        product = pool[counter["p"] % len(pool)]
        counter["p"] += 1
        return product

    def std(sid: str, **over) -> PlannedRow:
        sku, name, base = next_product()
        row = PlannedRow(
            file_key=fk, scenario_id=sid, ticket=next_ticket(), sale_date=date(2026, 10, rng.randint(day_min, 31)),
            sku=sku, sku_written=sku, name=name, qty=rng.randint(1, 5), price=_price_for(fk, base),
        )
        for key, value in over.items():
            setattr(row, key, value)
        return row

    def exact_total(row: PlannedRow) -> None:
        if fk in ("N", "C"):
            row.total = row.qty * row.price

    rows: list = []
    first = None
    for sid, _name, _codes, _status, _kpi, counts in RECORD_SCENARIOS:
        for i in range(counts.get(fk, 0)):
            if sid == "S01":
                r = std(sid)
                if fk in ("N", "C") and i < 2:  # extremos inclusivos del período (P-02)
                    r.sale_date = PERIOD[i]
                exact_total(r)
            elif sid == "S02":
                r = std(sid, qty=-1)
                exact_total(r)
            elif sid == "S03":
                r = std(sid, qty=0)
                exact_total(r)
            elif sid == "S04":
                r = std(sid, price=Decimal("0") if fk == "N" else Decimal("0.00"))
                exact_total(r)
            elif sid == "S05":
                r = std(sid, price=Decimal("-2.50") if fk == "C" else Decimal("-2.5"), qty=rng.randint(1, 4))
                exact_total(r)
            elif sid == "S06":
                r = std(sid)
                exact_total(r)
                if fk in ("N", "SX"):
                    r.ticket = None
                elif fk == "C":
                    r.date_text, r.sale_date = "N/A", None
                else:
                    r.qty = None
            elif sid == "S07":
                r = PlannedRow(file_key=fk, scenario_id=sid, name="Total general", general_total=True)
            elif sid == "S08":
                r = std(sid)
                if fk == "N":
                    r.price_text, r.price = "12,5,0", None
                    r.total = r.qty * Decimal("12.5")
                elif fk == "C":
                    r.qty_text, r.qty = "12 cajas", None
                    r.total = 2 * r.price
                else:
                    r.price_text, r.price = "dos", None
            elif sid == "S09":
                r = std(sid, qty=Decimal("1.5"))
                exact_total(r)
                if fk == "C":
                    r.qty_text = "1,5"
            elif sid == "S10":
                r = std(sid, sale_date=None)
                r.date_text = {"N": "31/02/2026", "C": "31-02-2026", "SC": "31/02/2026"}[fk]
                exact_total(r)
            elif sid == "S11":
                r = std(sid, sale_date=None)
                r.date_text = ("03/10/2026", "09/10/2026")[i]
            elif sid == "S12":
                r = std(sid)
                r.sale_date = date(2026, 11, 1) if fk == "N" else date(2026, 9, 30)
                r.date_text = {"N": "01/11/2026", "SX": "2026-09-30", "SC": "30/09/2026"}[fk]
                exact_total(r)
            elif sid == "S13":
                r = std(sid)
                r.total = r.qty * r.price + Decimal("0.02")
            elif sid == "S14":
                r = std(sid)
                r.sku = r.sku_written = UNKNOWN_SKUS[("N", "C", "SX").index(fk)]
                exact_total(r)
            elif sid == "S15":
                r = std(sid, branch_in_file="CENTRO")
            elif sid == "S16":
                if i == 0:
                    r = first = std(sid)
                    exact_total(r)
                else:
                    r = PlannedRow(**{k: getattr(first, k) for k in _COPY_FIELDS})
                    r.sku_written = first.sku.lower()
            elif sid == "S17":
                sku, name, price = PRODUCTS[0]
                r = std(sid, ticket="SUR-70001", sale_date=date(2026, 10, 20), qty=3, sku=sku, name=name, price=price,
                        sku_written=sku + (" " if fk == "SC" else ""))
            elif sid == "S18":
                if i == 0:
                    r = first = std(sid)
                    exact_total(r)
                else:
                    r = PlannedRow(**{k: getattr(first, k) for k in _COPY_FIELDS})
            elif sid == "S19":
                r = std(sid, sale_date=None, qty=None, qty_text="dos")
                r.total = 2 * r.price
            elif sid == "S20":
                text = {"N": ("=1+1", "@SUM(1,1)"), "C": ("+1", " -1+2"), "SX": ("#N/A",)}[fk][i]
                r = std(sid, name=text)
                exact_total(r)
            elif sid == "S21":
                r = std(sid, name=("\t=1+1", "\x07Texto con control")[i])
            else:  # pragma: no cover - tabla cerrada
                raise AssertionError(sid)
            if fk == "SX" and sid != "S15":
                r.branch_in_file = ("SUR", "sur", "Sucursal Sur")[len(rows) % 3]
            rows.append(r)

    # S24 (ruido benigno): atributo de filas S01; al menos una fila por categoría aplicable.
    s01 = [r for r in rows if r.scenario_id == "S01"]
    categories = {"N": ("date_text", "sku_lower", "sku_padded"), "C": ("sku_padded", "thousands"),
                  "SX": ("sku_lower",), "SC": ("sku_lower", "sku_padded")}[fk]
    free = list(range(len(s01)))
    rng.shuffle(free)
    for cat in categories:
        for _ in range(3):
            r = s01[free.pop()]
            if cat == "date_text":
                r.date_text = r.sale_date.strftime("%d/%m/%Y")
            elif cat == "sku_lower":
                r.sku_written = r.sku.lower()
            elif cat == "sku_padded":
                r.sku_written = (" " + r.sku) if rng.randint(0, 1) else (r.sku + " ")
            else:  # thousands: 5 × 250 = 1250 -> «1.250,00»
                sku, name, base = PRODUCTS[8]
                r.sku = r.sku_written = sku
                r.name, r.qty, r.price = name, 5, _price_for(fk, base)
                exact_total(r)
            r.noise.append(cat)

    # Orden de escritura: se mezclan todas las filas salvo «Total general» (al final) y se intercala la fila
    # en blanco en una posición intermedia (Plan §4.4).
    general = [r for r in rows if r.general_total]
    body = [r for r in rows if not r.general_total]
    rng.shuffle(body)
    if general:
        general[0].total = sum((r.total for r in body if r.total is not None), Decimal("0"))
    items: list = list(body)
    if BLANK_ROWS[fk]:
        items.insert(rng.randint(len(items) // 4, 3 * len(items) // 4), None)
    items.extend(general)
    return items


def build_plan(seed: int) -> Plan:
    """Plan completo del dataset. Pura; único consumidor del RNG; ejecuta las autocomprobaciones del plan."""
    if type(seed) is not int or not 0 <= seed <= SEED_MAX:
        raise GenerateError("argument", "la semilla debe ser un entero en [0, 2^63 - 1]")
    layout = {}
    for fk in FILE_KEYS:
        items = _build_file(fk, _file_rng(seed, FILES[fk]["source_id"]))
        number = FILES[fk]["header_row"] + 1
        for item in items:
            if item is not None:
                item.source_row = number
                item.source_line = number if FILES[fk]["kind"] == "csv" else None
            number += 1
        layout[fk] = items
    plan = Plan(seed=seed, layout=layout)
    self_check(plan)
    return plan


# --------------------------------------------------------------------------------------------------
# Autocomprobaciones del plan (fase 1; Plan §8.3, §6.6, §3.6)
# --------------------------------------------------------------------------------------------------


def native_number(value: Decimal) -> int | float:
    """Valor tal como se escribe en una celda numérica XLSX (Plan §6.6 a)."""
    if value == value.to_integral_value():
        return int(value)
    return float(value)


def _check_native(value: Decimal) -> None:
    """Condiciones de Plan §6.6 (a) para un valor no entero escrito como ``float``; los enteros, escala 0."""
    if value == value.to_integral_value():
        if value.as_tuple().exponent != 0:
            raise GenerateError("prepare", f"autocomprobación del plan: entero con exponente distinto de 0: {value}")
        return
    text = str(value)
    if "E" in text.upper() or "%.16g" % float(value) != text or repr(float(text)) != text:
        raise GenerateError("prepare", f"autocomprobación del plan: {value} no cumple Plan §6.6 (a)")


def self_check(plan: Plan) -> None:
    """Conteos = SCENARIOS, unicidad de tickets, R13/R12 accidentales, Decimal y variantes de SKU."""

    def fail(msg: str) -> None:
        raise GenerateError("prepare", f"autocomprobación del plan: {msg}")

    for fk in FILE_KEYS:
        items = plan.layout[fk]
        rows = [r for r in items if r is not None]
        counts: dict = {}
        for r in rows:
            counts[r.scenario_id] = counts.get(r.scenario_id, 0) + 1
        for sid, _n, _c, status, _k, per_file in SCENARIOS:
            if status is not None and counts.get(sid, 0) != per_file.get(fk, 0):
                fail(f"escenario {sid} en {fk}: {counts.get(sid, 0)} filas, se esperaban {per_file.get(fk, 0)}")
        if sum(1 for i in items if i is None) != BLANK_ROWS[fk]:
            fail(f"filas en blanco de {fk}")
        for r in rows:
            if r.sku_written not in {r.sku, r.sku.lower(), " " + r.sku, r.sku + " "}:
                fail(f"variante de SKU no permitida: {r.sku_written!r}")
            if r.sku:
                if any(ch not in SKU_CHARS for ch in r.sku):
                    fail(f"SKU canónico fuera de [A-Z0-9-]: {r.sku!r}")
                if (r.scenario_id == "S14") == (r.sku in CATALOG_SKUS):
                    fail(f"SKU {r.sku!r}: pertenencia al catálogo incorrecta en {r.scenario_id}")
            if fk in ("N", "SX"):
                for value in (r.price, r.qty if isinstance(r.qty, Decimal) else None):
                    if value is not None:
                        _check_native(value)
            if fk == "C" and r.price is not None and r.price.as_tuple().exponent != -2:
                fail("precio de C sin escala 2")
            if fk == "SC" and r.price is not None and (r.price.as_tuple().exponent > 0 or "E" in str(r.price)):
                fail("precio de SC con exponente")
            if fk == "N" and r.total is not None and Decimal(repr(native_number(r.total))) != r.total:
                fail("total no representable como número nativo")
    groups: dict = {}
    for fk in FILE_KEYS:
        for r in plan.layout[fk]:
            if r is not None and r.ticket is not None:
                groups.setdefault((FILES[fk]["branch"], r.ticket, r.sku), []).append(r.scenario_id)
    dup = sorted(v[0] for v in groups.values() if len(v) > 1)
    if dup != ["S16", "S17", "S18"] or any(len(v) > 2 or len(set(v)) != 1 for v in groups.values()):
        fail(f"grupos R13 inesperados: {dup}")


# --------------------------------------------------------------------------------------------------
# Celdas (Plan §4.2: plan_to_cells)
# --------------------------------------------------------------------------------------------------


def _fmt_centro(value: Decimal) -> str:
    """Texto CSV de CENTRO: decimal ``,`` y miles ``.``; sin ``float`` ni ``locale`` (Plan §6.6 b)."""
    if value == value.quantize(D2):
        value = value.quantize(D2)
    text = format(value, "f")
    sign = "-" if text.startswith("-") else ""
    whole, _, frac = text.lstrip("-").partition(".")
    groups = []
    while len(whole) > 3:
        groups.insert(0, whole[-3:])
        whole = whole[:-3]
    groups.insert(0, whole)
    return sign + ".".join(groups) + ("," + frac if frac else "")


def _qty_cell(r: PlannedRow, fk: str):
    if r.qty_text is not None:
        return r.qty_text
    if r.qty is None:
        return None
    if fk in ("N", "SX"):
        return native_number(r.qty) if isinstance(r.qty, Decimal) else int(r.qty)
    return str(r.qty)


def _date_cell(r: PlannedRow, pattern: str):
    if r.date_text is not None:
        return r.date_text
    return r.sale_date.strftime(pattern) if r.sale_date is not None else None


def _row_cells(r: PlannedRow, fk: str) -> list:
    if fk == "N":
        if r.general_total:
            return [None, None, None, r.name, None, None, native_number(r.total)]
        price = r.price_text if r.price_text is not None else (native_number(r.price) if r.price is not None else None)
        fecha = r.date_text if r.date_text is not None else r.sale_date  # nativa si no hay texto
        return [fecha, r.ticket, r.sku_written, r.name, _qty_cell(r, fk), price,
                native_number(r.total) if r.total is not None else None]
    if fk == "C":
        price = r.price_text if r.price_text is not None else (_fmt_centro(r.price) if r.price is not None else None)
        obs = ("Sin observaciones", "Entrega en mostrador", "Pago en efectivo")[r.source_row % 3]
        return [r.ticket, _date_cell(r, "%d-%m-%Y"), r.sku_written, r.name, _qty_cell(r, fk), price,
                _fmt_centro(r.total) if r.total is not None else None, obs]
    if fk == "SX":
        price = r.price_text if r.price_text is not None else (native_number(r.price) if r.price is not None else None)
        return [_date_cell(r, "%Y-%m-%d"), r.ticket, r.sku_written, r.name, _qty_cell(r, fk), price,
                r.branch_in_file, f"Caja {r.source_row % 3 + 1}"]
    price = r.price_text if r.price_text is not None else (format(r.price, "f") if r.price is not None else None)
    return [_date_cell(r, "%d/%m/%Y"), r.ticket, r.sku_written, r.name, _qty_cell(r, fk), price]


def plan_to_cells(plan: Plan) -> dict:
    """Matriz de celdas por archivo: ``{"cells": {(fila, col): valor}}`` (XLSX) o ``{"records": [...]}`` (CSV)."""
    out = {}
    for fk in FILE_KEYS:
        spec = FILES[fk]
        headers = list(spec["headers"])
        if spec["kind"] == "xlsx":
            cells = {}
            if fk == "N":
                cells[(1, 1)] = "Comercio Norte - Ventas"
                cells[(2, 1)] = "Octubre de 2026"
            for c, h in enumerate(headers, 1):
                cells[(spec["header_row"], c)] = h
            for item in plan.layout[fk]:
                if item is not None:
                    for c, v in enumerate(_row_cells(item, fk), 1):
                        if v is not None:
                            cells[(item.source_row, c)] = v
            out[fk] = {"cells": cells}
        else:
            records = [headers]
            for item in plan.layout[fk]:
                records.append([] if item is None else ["" if v is None else v for v in _row_cells(item, fk)])
            out[fk] = {"records": records}
    return out


# --------------------------------------------------------------------------------------------------
# Hash de contenido cells-v1 (Plan §7) y serialización
# --------------------------------------------------------------------------------------------------


def _cell_line(sheet: str, row: int, col: int, value) -> str:
    if isinstance(value, datetime):
        if value.time() != datetime.min.time():
            raise GenerateError("prepare", "celda de fecha con hora")
        value = value.date()
    if isinstance(value, date):
        kind, text = "d", value.isoformat()
    elif isinstance(value, str):
        kind, text = "s", value
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        kind, text = "n", repr(value)
    else:
        raise GenerateError("prepare", f"tipo de celda no admitido: {type(value).__name__}")
    return json.dumps([sheet, row, col, kind, text], ensure_ascii=False, separators=(",", ":"))


def digest_cells(sheet: str, cells: dict) -> str:
    """sha256 de las líneas JSON de las celdas no vacías en orden (fila, columna)."""
    lines = [_cell_line(sheet, r, c, v) for (r, c), v in sorted(cells.items()) if v is not None and v != ""]
    return hashlib.sha256("\n".join(lines).encode("utf-8")).hexdigest()


def digest_records(records) -> str:
    """``cells-v1`` de un CSV: registros de ``csv.reader`` (hoja vacía; solo campos no vacíos)."""
    cells = {(r, c): v for r, rec in enumerate(records, 1) for c, v in enumerate(rec, 1) if v != ""}
    return digest_cells("", cells)


def canonical_json(obj) -> bytes:
    return (json.dumps(obj, sort_keys=True, indent=2, ensure_ascii=False) + "\n").encode("utf-8")


def render_xlsx(sheet: str, cells: dict) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for (r, c), v in sorted(cells.items()):
        cell = ws.cell(row=r, column=c)
        cell.value = v
        if isinstance(v, str):
            cell.data_type = "s"  # sin esto =1+1 sería fórmula y #N/A un error (§11.2)
    wb.properties.creator = "synthetic"
    wb.properties.created = datetime(2026, 10, 31)
    wb.properties.modified = datetime(2026, 10, 31)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def render_csv(records, delimiter: str, eol: str, encoding: str) -> bytes:
    buf = io.StringIO(newline="")
    writer = csv.writer(buf, delimiter=delimiter, quoting=csv.QUOTE_MINIMAL, lineterminator=eol)
    for rec in records:
        writer.writerow(rec)
    return buf.getvalue().encode(encoding)  # estricto: sin errors="replace"


# --------------------------------------------------------------------------------------------------
# Ground truth (Plan §6) a partir del plan; nunca relee archivos
# --------------------------------------------------------------------------------------------------

RECORD_COLUMNS = ("source_id", "source_file", "source_sheet", "source_row", "source_line", "branch_id",
                  "expected_status", "expected_issue_codes", "expected_line_amount", "kpi_effect", "scenario")
ISSUE_COLUMNS = ("level", "source_file", "branch_id", "source_row", "issue_code")
FILE_COLUMNS = ("source_file", "source_id", "expected_outcome", "rejection_code", "preamble_rows", "header_rows",
                "blank_rows", "expected_records")


def plan_to_ground_truth(plan: Plan) -> dict:
    records, issues, files = [], [], []
    by_file: dict = {}
    by_branch: dict = {}
    by_status: dict = {}
    issue_counts = {code: 0 for code in ISSUE_CODES}
    oos: dict = {}
    dup_members: dict = {}
    dup_amount = Decimal("0")
    units = 0
    gross = Decimal("0")
    gross_branch: dict = {}
    zero_price = 0
    total = 0
    for fk in FILE_KEYS:
        spec = FILES[fk]
        rows = [r for r in plan.layout[fk] if r is not None]
        files.append({
            "source_file": spec["name"], "source_id": spec["source_id"], "expected_outcome": "processed",
            "rejection_code": "", "preamble_rows": spec["preamble"], "header_rows": 1,
            "blank_rows": BLANK_ROWS[fk], "expected_records": len(rows),
        })
        if fk in FILE_ISSUES:
            issues.append((spec["source_id"], spec["name"], 0, 0, ("file", spec["name"], spec["branch"], "", FILE_ISSUES[fk])))
            issue_counts[FILE_ISSUES[fk]] += 1
        for r in rows:
            total += 1
            line = r.line_amount
            records.append({
                "source_id": spec["source_id"], "source_file": spec["name"], "source_sheet": spec["sheet"],
                "source_row": str(r.source_row),
                "source_line": "" if r.source_line is None else str(r.source_line),
                "branch_id": spec["branch"], "expected_status": r.expected_status,
                "expected_issue_codes": ";".join(r.expected_codes),
                "expected_line_amount": "" if line is None else str(line),
                "kpi_effect": r.kpi_effect, "scenario": r.scenario_id,
            })
            for code in r.expected_codes:
                issues.append((spec["source_id"], spec["name"], 1, r.source_row,
                               ("record", spec["name"], spec["branch"], str(r.source_row), code)))
                issue_counts[code] += 1
                if r.expected_status == "out_of_scope":
                    oos[code] = oos.get(code, 0) + 1
            by_file[spec["name"]] = by_file.get(spec["name"], 0) + 1
            by_branch[spec["branch"]] = by_branch.get(spec["branch"], 0) + 1
            by_status[r.expected_status] = by_status.get(r.expected_status, 0) + 1
            if r.expected_status == "valid":
                units += r.qty
                gross += line
                gross_branch[spec["branch"]] = gross_branch.get(spec["branch"], Decimal("0")) + line
                if "R08_ZERO_PRICE" in r.expected_codes:
                    zero_price += 1
            if r.scenario_id in ("S16", "S17", "S18"):
                if r.scenario_id not in dup_members:
                    dup_amount += line  # Σ(n−1)×line_amount: los miembros comparten cantidad y precio
                    dup_members[r.scenario_id] = 1
                else:
                    dup_members[r.scenario_id] += 1
    records.sort(key=lambda d: (d["source_id"], d["source_file"], d["source_sheet"], int(d["source_row"])))
    issues.sort(key=lambda t: t[:4] + (t[4][4],))
    files.sort(key=lambda d: d["source_id"])
    summary = {
        "records": {"total": total, "by_file": dict(sorted(by_file.items())),
                    "by_branch": dict(sorted(by_branch.items())), "by_status": dict(sorted(by_status.items()))},
        "issues_by_code": dict(sorted(issue_counts.items())),
        "out_of_scope_by_code": dict(sorted(oos.items())),
        "duplicates": {"groups": len(dup_members), "records": sum(dup_members.values()),
                       "excess": sum(n - 1 for n in dup_members.values()),
                       "potentially_duplicated_amount": str(dup_amount)},
        "sales": {"gross_ordinary_total": str(gross),
                  "gross_ordinary_by_branch": {k: str(v) for k, v in sorted(gross_branch.items())},
                  "units": units, "zero_price_lines": zero_price},
    }
    return {"records": records, "issues": [t[4] for t in issues], "files": files, "summary": summary}


def _csv_bytes(columns, rows) -> bytes:
    buf = io.StringIO(newline="")
    writer = csv.writer(buf, delimiter=",", quoting=csv.QUOTE_MINIMAL, lineterminator="\n")
    writer.writerow(columns)
    for row in rows:
        writer.writerow(row)
    return buf.getvalue().encode("utf-8")


# --------------------------------------------------------------------------------------------------
# Manifiesto (Plan §7): validación exacta
# --------------------------------------------------------------------------------------------------


def validate_manifest(obj) -> list:
    """Problemas de estructura del manifiesto; lista vacía si es válido. Sin hash propio (no circular)."""
    problems = []
    if not isinstance(obj, dict):
        return ["el manifiesto no es un objeto"]
    if set(obj) != {"seed", "generator_version", "config_version", "files"}:
        problems.append(f"claves de primer nivel incorrectas: {sorted(obj)}")
    seed = obj.get("seed")
    if type(seed) is not int or not 0 <= seed <= SEED_MAX:
        problems.append("seed no es un entero en rango")
    if not isinstance(obj.get("generator_version"), str):
        problems.append("generator_version no es una cadena")
    if type(obj.get("config_version")) is not int:
        problems.append("config_version no es un entero")
    entries = obj.get("files")
    if not isinstance(entries, list) or len(entries) != 9:
        problems.append("files debe ser una lista de exactamente 9 elementos")
        return problems
    paths = []
    for entry in entries:
        if not isinstance(entry, dict) or set(entry) != {"path", "content_sha256"}:
            problems.append("entrada con campos distintos de path y content_sha256")
            continue
        path, digest = entry["path"], entry["content_sha256"]
        if not isinstance(path, str) or not path:
            problems.append("path no es una cadena no vacía")
        else:
            if path.startswith("/") or "\\" in path or ":" in path or any(
                    seg in ("", ".", "..") for seg in path.split("/")):
                problems.append(f"ruta no segura: {path!r}")
            paths.append(path)
        if not isinstance(digest, str) or len(digest) != 64 or any(ch not in "0123456789abcdef" for ch in digest):
            problems.append(f"content_sha256 inválido en {path!r}")
    if len(set(paths)) != len(paths):
        problems.append("rutas duplicadas")
    if set(paths) != set(MANIFEST_PATHS):
        problems.append("el conjunto de rutas no es exactamente el esperado")
    elif paths != list(MANIFEST_PATHS):
        problems.append("orden de rutas no canónico")
    return problems


# --------------------------------------------------------------------------------------------------
# Preparación en memoria (fase 1)
# --------------------------------------------------------------------------------------------------


@dataclass
class Prepared:
    seed: int
    plan: Plan
    ground_truth: dict
    blobs: dict          # ruta relativa -> bytes (10 rutas)
    digests: dict        # ruta relativa -> content_sha256 (nueve rutas)
    manifest: dict


def prepare(seed: int) -> Prepared:
    plan = build_plan(seed)
    gt = plan_to_ground_truth(plan)
    cells = plan_to_cells(plan)
    blobs: dict = {}
    digests: dict = {}
    for fk in FILE_KEYS:
        spec = FILES[fk]
        rel = "raw/" + spec["name"]
        if spec["kind"] == "xlsx":
            blobs[rel] = render_xlsx(spec["sheet"], cells[fk]["cells"])
            digests[rel] = digest_cells(spec["sheet"], cells[fk]["cells"])
        else:
            data = render_csv(cells[fk]["records"], spec["delimiter"], "\r\n", "cp1252" if not spec["bom"] else "utf-8")
            blobs[rel] = (b"\xef\xbb\xbf" if spec["bom"] else b"") + data
            digests[rel] = digest_records(cells[fk]["records"])
    catalog = [[sku, name] for sku, name, _price in PRODUCTS]
    blobs["reference/products.csv"] = _csv_bytes(("product_sku", "product_name"), catalog)
    digests["reference/products.csv"] = digest_records([["product_sku", "product_name"]] + catalog)
    for rel, cols, rows in (
        ("ground_truth/expected_records.csv", RECORD_COLUMNS, [[d[c] for c in RECORD_COLUMNS] for d in gt["records"]]),
        ("ground_truth/expected_issues.csv", ISSUE_COLUMNS, [list(t) for t in gt["issues"]]),
        ("ground_truth/expected_files.csv", FILE_COLUMNS, [[str(d[c]) for c in FILE_COLUMNS] for d in gt["files"]]),
    ):
        blobs[rel] = _csv_bytes(cols, rows)
        digests[rel] = digest_records([list(cols)] + rows)
    summary_rel = "ground_truth/expected_summary.json"
    blobs[summary_rel] = canonical_json(gt["summary"])
    digests[summary_rel] = hashlib.sha256(blobs[summary_rel]).hexdigest()
    manifest = {
        "seed": seed, "generator_version": GENERATOR_VERSION, "config_version": TARGET_CONFIG_VERSION,
        "files": [{"path": p, "content_sha256": digests[p]} for p in MANIFEST_PATHS],
    }
    problems = validate_manifest(manifest)
    if problems:
        raise GenerateError("prepare", "manifiesto en memoria inválido: " + "; ".join(problems))
    blobs[MANIFEST_REL] = canonical_json(manifest)
    return Prepared(seed, plan, gt, blobs, digests, manifest)


# --------------------------------------------------------------------------------------------------
# Verificación por relectura (fase 3; no alimenta el ground truth)
# --------------------------------------------------------------------------------------------------


def check_terminators(rel: str, data: bytes) -> str | None:
    """Control físico de Plan §5.1 sobre los bytes (≠ ``cells-v1``). Devuelve un problema o ``None``."""
    if EOL[rel] == b"\r\n":
        if not (data.count(b"\r") == data.count(b"\n") == data.count(b"\r\n")) or not data.endswith(b"\r\n"):
            return "terminadores distintos de \\r\\n (mezclados o aislados)"
    elif b"\r" in data or not data.endswith(b"\n"):
        return "el archivo exige terminador \\n y contiene \\r o no termina en \\n"
    return None


def _read_xlsx(path: Path):
    wb = openpyxl.load_workbook(path)
    try:
        names = list(wb.sheetnames)
        cells, kinds = {}, set()
        for row in wb[names[0]].iter_rows():
            for cell in row:
                if cell.value is not None:
                    cells[(cell.row, cell.column)] = cell.value
                    kinds.add(cell.data_type)
        return names, cells, kinds
    finally:
        wb.close()


def _verify_staged(staging: Path, prepared: Prepared) -> None:
    def fail(msg: str) -> None:
        raise GenerateError("prepare", f"validación del staging: {msg}")

    for fk in FILE_KEYS:
        spec = FILES[fk]
        rel = "raw/" + spec["name"]
        path = staging / rel
        if spec["kind"] == "xlsx":
            names, cells, kinds = _read_xlsx(path)
            if names != [spec["sheet"]]:
                fail(f"{rel}: hojas {names}")
            if any(cells.get((spec["header_row"], c)) != h for c, h in enumerate(spec["headers"], 1)):
                fail(f"{rel}: encabezados")
            if kinds & {"f", "e"}:
                fail(f"{rel}: celdas de fórmula o error")
            if digest_cells(spec["sheet"], cells) != prepared.digests[rel]:
                fail(f"{rel}: content_sha256 no coincide con el manifiesto en memoria")
        else:
            data = path.read_bytes()
            if data.startswith(b"\xef\xbb\xbf") != spec["bom"]:
                fail(f"{rel}: BOM")
            problem = check_terminators(rel, data)
            if problem:
                fail(f"{rel}: {problem}")
            records = list(csv.reader(io.StringIO(data.decode(spec["encoding"]), newline=""),
                                      delimiter=spec["delimiter"]))
            if records[0] != list(spec["headers"]):
                fail(f"{rel}: encabezados")
            if digest_records(records) != prepared.digests[rel]:
                fail(f"{rel}: content_sha256 no coincide con el manifiesto en memoria")
    for rel in MANIFEST_PATHS[4:]:
        data = (staging / rel).read_bytes()
        if rel.endswith(".json"):
            if hashlib.sha256(data).hexdigest() != prepared.digests[rel]:
                fail(f"{rel}: content_sha256 no coincide con el manifiesto en memoria")
            continue
        problem = check_terminators(rel, data)
        if problem:
            fail(f"{rel}: {problem}")
        if data.startswith(b"\xef\xbb\xbf"):
            fail(f"{rel}: no debe llevar BOM")
        records = list(csv.reader(io.StringIO(data.decode("utf-8"), newline="")))
        if digest_records(records) != prepared.digests[rel]:
            fail(f"{rel}: content_sha256 no coincide con el manifiesto en memoria")
    # manifest.json por separado: estructura exacta y serialización canónica; no tiene hash propio.
    raw = (staging / MANIFEST_REL).read_bytes()
    try:
        loaded = json.loads(raw.decode("utf-8"))
    except ValueError:
        fail("manifest.json no es JSON válido")
    problems = validate_manifest(loaded)
    if problems:
        fail("manifest.json: " + "; ".join(problems))
    if raw != canonical_json(prepared.manifest):
        fail("manifest.json no es la serialización canónica del manifiesto en memoria")


# --------------------------------------------------------------------------------------------------
# Publicación (Plan §8.3). Las operaciones de sistema pasan por estas funciones para poder inyectar fallos.
# --------------------------------------------------------------------------------------------------


def _fs_mkdir(path) -> None:
    os.mkdir(path)


def _fs_rmdir(path) -> None:
    os.rmdir(path)


def _fs_unlink(path) -> None:
    os.unlink(path)


def _fs_replace(src, dst) -> None:
    os.replace(src, dst)


def _fs_write(path, data: bytes) -> None:
    with open(path, "xb") as handle:
        handle.write(data)


def _fs_dev(path) -> int:
    return os.stat(path).st_dev


def _is_real_dir(path) -> bool:
    return os.path.isdir(path) and not os.path.islink(path)


def _is_regular(path) -> bool:
    return os.path.isfile(path) and not os.path.islink(path)


def _check_destination(root: Path, overwrite: bool, phase: str) -> None:
    for sub in SUBDIRS:
        d = root / sub
        if os.path.lexists(d) and not _is_real_dir(d):
            raise GenerateError("prepare", f"{phase}: {d} existe y no es un directorio (archivo, enlace u otro tipo)")
    for rel in ALL_PATHS:
        target = root / rel
        if os.path.lexists(target):
            if not overwrite:
                raise GenerateError("prepare", f"{phase}: {target} ya existe; use --overwrite para reemplazarlo")
            if not _is_regular(target):
                raise GenerateError("prepare", f"{phase}: {target} existe y no es un archivo regular")


def _cleanup_staging(staging: Path) -> bool:
    """Elimina SOLO el staging creado por esta ejecución. Devuelve ``True`` si no queda rastro."""
    for rel in ALL_PATHS:
        try:
            if os.path.lexists(staging / rel):
                _fs_unlink(staging / rel)
        except OSError:
            pass
    for sub in SUBDIRS:
        try:
            if os.path.lexists(staging / sub):
                _fs_rmdir(staging / sub)
        except OSError:
            pass
    try:
        if os.path.lexists(staging):
            _fs_rmdir(staging)
    except OSError:
        pass
    return not os.path.lexists(staging)


def _remove_created(created: list) -> None:
    for d in reversed(created):  # solo directorios vacíos que creó esta ejecución
        try:
            _fs_rmdir(d)
        except OSError:
            pass


INCOMPLETE = "conjunto incompleto: manifest.json ausente; regenere con --overwrite"


def generate(seed: int, output_root, *, overwrite: bool = False) -> dict:
    """Genera y publica el conjunto (Plan §8.3). Lanza :class:`GenerateError`; devuelve un resumen si va bien.

    Precondición de exclusividad operacional: ningún proceso externo modifica los destinos ni el staging
    durante la ejecución. No hay atomicidad del conjunto, ``fsync`` ni exclusión mutua completa.
    """
    # Fases 0 y 1: argumentos y preparación en memoria (nada tocado).
    prepared = prepare(seed)
    root = Path(output_root)
    staging = root / STAGING_NAME

    # Fase 2: comprobación previa del destino (solo lectura).
    if os.path.lexists(root):
        if not _is_real_dir(root):
            raise GenerateError("prepare", f"fase 2: {root} existe y no es un directorio")
        _check_destination(root, overwrite, "fase 2")
        if os.path.lexists(staging):
            raise GenerateError(
                "prepare",
                f"fase 2: existe {staging} (residuo de una ejecución anterior). No se borra automáticamente: "
                "compruebe que no hay una generación activa, elimínelo a mano y vuelva a ejecutar con la misma semilla",
            )
        root_exists = True
    else:
        parent = Path(os.path.abspath(root)).parent
        if not _is_real_dir(parent):
            raise GenerateError("prepare", f"fase 2: {root} no existe y su directorio padre {parent} tampoco")
        root_exists = False

    created: list = []     # directorios creados por esta ejecución (raíz y subdirectorios de destino)
    state = {"staging": False}

    def abort(kind: str, message: str):
        if state["staging"]:
            _cleanup_staging(staging)
        _remove_created(created)
        raise GenerateError(kind, message)

    # Fase 3: staging exclusivo, escritura y validación.
    try:
        if not root_exists:
            _fs_mkdir(root)
            created.append(root)
        _fs_mkdir(staging)
        state["staging"] = True
        for sub in SUBDIRS:
            _fs_mkdir(staging / sub)
        for rel in ALL_PATHS:
            _fs_write(staging / rel, prepared.blobs[rel])
        _verify_staged(staging, prepared)
    except GenerateError as exc:
        abort(exc.kind, exc.message)
    except OSError as exc:
        abort("prepare", f"fase 3: {exc}")

    # Fase 3b: directorios de destino. Todo error detectable ocurre aquí, antes de tocar el conjunto anterior.
    try:
        for sub in SUBDIRS:
            if not os.path.lexists(root / sub):
                _fs_mkdir(root / sub)
                created.append(root / sub)
        _check_destination(root, overwrite, "fase 3b")
        staging_dev = _fs_dev(staging)
        for sub in SUBDIRS:  # comprobación PRELIMINAR: no garantiza el éxito de os.replace (EXDEV posible)
            if _fs_dev(root / sub) != staging_dev:
                raise GenerateError("prepare", f"fase 3b: {root / sub} está en otro sistema de archivos")
    except GenerateError as exc:
        abort(exc.kind, exc.message)
    except OSError as exc:
        abort("prepare", f"fase 3b: {exc}")

    # Fase 4: publicación (a) manifiesto anterior, (b) nueve archivos en orden canónico, (c) manifiesto nuevo.
    manifest_target = root / MANIFEST_REL
    if overwrite and os.path.lexists(manifest_target):
        try:
            _fs_unlink(manifest_target)
        except OSError as exc:
            abort("prepare", f"fase 4(a): no se pudo eliminar el manifest.json anterior ({exc}); "
                             "el conjunto anterior sigue intacto")
    try:
        for rel in MANIFEST_PATHS:
            _fs_replace(staging / rel, root / rel)
        _fs_replace(staging / MANIFEST_REL, manifest_target)
    except OSError as exc:  # caso 1: incluye EXDEV; no se copia como alternativa
        _cleanup_staging(staging)
        raise GenerateError("interrupted", f"{INCOMPLETE} (fallo en la publicación: {exc})") from exc

    # 4(d): limpieza. Un fallo aquí NO invalida el conjunto ya publicado y validado (caso 3).
    if not _cleanup_staging(staging):
        raise GenerateError(
            "cleanup",
            "conjunto publicado y validado (manifest.json presente); limpieza incompleta: no se pudo eliminar "
            f"{STAGING_NAME}; compruebe que no hay una generación activa y elimínelo manualmente antes de "
            "volver a ejecutar",
        )
    return {
        "seed": seed,
        "output_root": str(root),
        "records": prepared.ground_truth["summary"]["records"]["total"],
        "files": len(ALL_PATHS),
        "manifest_sha256": hashlib.sha256(prepared.blobs[MANIFEST_REL]).hexdigest(),
    }
