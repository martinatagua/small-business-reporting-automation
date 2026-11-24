"""Utilidades INDEPENDIENTES de los tests de F3: no importan ``synthetic`` (salvo donde un test lo declara).

Contienen la transcripción propia de §10.2, una implementación propia de ``cells-v1``, del control físico de
terminadores y de la validación del manifiesto (T3-04, T3-06, T3-08, T3-19).
"""

import csv
import datetime as dt
import io
import json
import re
from pathlib import Path

import openpyxl

SEED = 20261001

# Transcripción propia de §10.2 (Plan §3.1): id -> (códigos, estado, kpi_effect, {archivo: registros}).
TABLE = {
    "S01": ((), "valid", "included", {"N": 40, "C": 40, "SX": 25, "SC": 25}),
    "S02": (("R07",), "out_of_scope", "excluded_out_of_scope", {"N": 1, "C": 1, "SC": 1}),
    "S03": (("R06",), "observed", "excluded_observed", {"N": 1, "SX": 1}),
    "S04": (("R08",), "valid", "included_warning", {"N": 1, "C": 1}),
    "S05": (("R09",), "observed", "excluded_observed", {"C": 1, "SX": 1}),
    "S06": (("R01",), "observed", "excluded_observed", {"N": 1, "C": 1, "SX": 1, "SC": 1}),
    "S07": (("R01",), "observed", "excluded_observed", {"N": 1}),
    "S08": (("R04",), "observed", "excluded_observed", {"N": 1, "C": 1, "SC": 1}),
    "S09": (("R05",), "observed", "excluded_observed", {"C": 1, "SX": 1}),
    "S10": (("R02",), "observed", "excluded_observed", {"N": 1, "C": 1, "SC": 1}),
    "S11": (("R03",), "observed", "excluded_observed", {"SC": 2}),
    "S12": (("R10",), "out_of_scope", "excluded_out_of_scope", {"N": 1, "SX": 1, "SC": 1}),
    "S13": (("R11",), "observed", "excluded_observed", {"N": 1, "C": 1}),
    "S14": (("R12",), "valid", "included_warning", {"N": 1, "C": 1, "SX": 1}),
    "S15": (("R14",), "observed", "excluded_observed", {"SX": 1}),
    "S16": (("R13",), "valid", "included_warning", {"N": 2}),
    "S17": (("R13",), "valid", "included_warning", {"SX": 1, "SC": 1}),
    "S18": (("R13",), "valid", "included_warning", {"C": 2}),
    "S19": (("R01", "R04"), "observed", "excluded_observed", {"N": 1}),
    "S20": ((), "valid", "included", {"N": 2, "C": 2, "SX": 1}),
    "S21": ((), "valid", "included", {"SC": 2}),
}

# Archivos: clave -> (ruta relativa a la raíz, source_id, hoja, fila de cabecera, encoding, delimitador)
SOURCES = {
    "N": ("raw/norte_ventas_2026-10.xlsx", "norte-xlsx", "Ventas", 3, None, None),
    "C": ("raw/centro_ventas_2026-10.csv", "centro-csv", "", 1, "cp1252", ";"),
    "SX": ("raw/sur_ventas_2026-10_tienda.xlsx", "sur-xlsx", "Datos", 1, None, None),
    "SC": ("raw/sur_ventas_2026-10_web.csv", "sur-csv", "", 1, "utf-8-sig", ","),
}

# Las nueve rutas del manifiesto y su orden canónico (copia literal propia del test; Plan §7).
EXPECTED_MANIFEST_PATHS = [
    "raw/centro_ventas_2026-10.csv",
    "raw/norte_ventas_2026-10.xlsx",
    "raw/sur_ventas_2026-10_tienda.xlsx",
    "raw/sur_ventas_2026-10_web.csv",
    "reference/products.csv",
    "ground_truth/expected_records.csv",
    "ground_truth/expected_issues.csv",
    "ground_truth/expected_files.csv",
    "ground_truth/expected_summary.json",
]
EOL_CRLF = {"raw/centro_ventas_2026-10.csv", "raw/sur_ventas_2026-10_web.csv"}
EOL_LF = {"reference/products.csv", "ground_truth/expected_records.csv", "ground_truth/expected_issues.csv",
          "ground_truth/expected_files.csv"}


def _line(sheet, row, col, value):
    if isinstance(value, dt.datetime):
        value = value.date()
    if isinstance(value, dt.date):
        kind, text = "d", value.isoformat()
    elif isinstance(value, str):
        kind, text = "s", value
    else:
        kind, text = "n", repr(value)
    return json.dumps([sheet, row, col, kind, text], ensure_ascii=False, separators=(",", ":"))


def _sha(lines):
    import hashlib
    return hashlib.sha256("\n".join(lines).encode("utf-8")).hexdigest()


def xlsx_rows(path):
    wb = openpyxl.load_workbook(path)
    ws = wb.worksheets[0]
    return wb.sheetnames, ws, [[c for c in row] for row in ws.iter_rows()]


def xlsx_digest(path) -> str:
    wb = openpyxl.load_workbook(path)
    ws = wb.worksheets[0]
    lines = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                lines.append(_line(ws.title, cell.row, cell.column, cell.value))
    return _sha(lines)


def csv_records(data: bytes, encoding: str, delimiter: str):
    return list(csv.reader(io.StringIO(data.decode(encoding), newline=""), delimiter=delimiter))


def csv_digest(data: bytes, encoding: str, delimiter: str) -> str:
    lines = [_line("", r, c, v) for r, rec in enumerate(csv_records(data, encoding, delimiter), 1)
             for c, v in enumerate(rec, 1) if v != ""]
    return _sha(lines)


def file_digest(root: Path, rel: str) -> str:
    path = root / rel
    if rel.endswith(".xlsx"):
        return xlsx_digest(path)
    if rel.endswith(".json"):
        import hashlib
        return hashlib.sha256(path.read_bytes()).hexdigest()
    enc, delim = {"raw/centro_ventas_2026-10.csv": ("cp1252", ";"),
                  "raw/sur_ventas_2026-10_web.csv": ("utf-8-sig", ",")}.get(rel, ("utf-8", ","))
    return csv_digest(path.read_bytes(), enc, delim)


def physical_problems(data: bytes, eol: bytes) -> list:
    """Control físico de terminadores (Plan §5.1) con regex; independiente del generador."""
    problems = []
    if eol == b"\r\n":
        if re.search(rb"(?<!\r)\n", data):
            problems.append("LF aislado")
        if re.search(rb"\r(?!\n)", data):
            problems.append("CR aislado")
        if not data.endswith(b"\r\n"):
            problems.append("sin CRLF final")
    else:
        if b"\r" in data:
            problems.append("contiene CR")
        if not data.endswith(b"\n"):
            problems.append("sin LF final")
    return problems


def manifest_problems(obj) -> list:
    """Validación exacta y propia del manifiesto (Plan §7)."""
    if not isinstance(obj, dict):
        return ["no es objeto"]
    out = []
    if sorted(obj) != ["config_version", "files", "generator_version", "seed"]:
        out.append("claves")
    files = obj.get("files")
    if not isinstance(files, list) or len(files) != 9:
        return out + ["files != 9 elementos"]
    paths = []
    for item in files:
        if not isinstance(item, dict) or sorted(item) != ["content_sha256", "path"]:
            out.append("campos de entrada")
            continue
        p, h = item["path"], item["content_sha256"]
        if not isinstance(p, str) or not p or p.startswith("/") or ".." in p.split("/") or "\\" in p:
            out.append(f"ruta no segura {p!r}")
        else:
            paths.append(p)
        if not isinstance(h, str) or not re.fullmatch(r"[0-9a-f]{64}", h):
            out.append(f"hash inválido {p!r}")
    if len(paths) != len(set(paths)):
        out.append("duplicadas")
    if set(paths) != set(EXPECTED_MANIFEST_PATHS):
        out.append("conjunto de rutas")
    elif paths != EXPECTED_MANIFEST_PATHS:
        out.append("orden")
    return out


def read_gt_csv(root: Path, name: str) -> list:
    data = (root / "ground_truth" / name).read_bytes()
    return list(csv.DictReader(io.StringIO(data.decode("utf-8"), newline="")))


def source_records(root: Path, key: str):
    """Filas físicas de una fuente como listas de valores (``None`` = celda vacía): índice 0 = fila 1."""
    rel, _sid, _sheet, _hdr, enc, delim = SOURCES[key]
    if enc is None:
        _names, ws, rows = xlsx_rows(root / rel)
        return [[c.value for c in row] for row in rows]
    return [[(v if v != "" else None) for v in rec] for rec in csv_records((root / rel).read_bytes(), enc, delim)]
