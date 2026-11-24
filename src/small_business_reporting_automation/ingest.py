"""F4 — Ingesta: descubrimiento, perfiles, lectura CSV acotada y frontera XLSX X1–X8.

Entrega filas de origen con valores NATIVOS y trazabilidad (§7.2). No normaliza ni evalúa reglas (F5–F7).
Plan: PHASE0_F4_PLAN Rev. 1.3 y las decisiones del propietario aprobadas (P-10 C, DT4-04/05/10/13/18/23/27/28/29/30).

Decisiones NO aprobadas (se aplica el comportamiento mínimo y se documenta en el README):
DT4-03 (extensión sensible a mayúsculas), DT4-07 (fila en blanco = todo vacío o solo espacios, ``str.strip()``),
DT4-11 (los fallos de zlib/openpyxl usan códigos existentes), DT4-17 (cota de filas = 1 048 576 dentro del
pre-escaneo de P-10 C).
"""

from __future__ import annotations

import csv
import errno
import fnmatch
import hashlib
import io
import os
import posixpath
import re
import stat
import struct
import sys
import warnings
import xml.etree.ElementTree as ET
import zipfile
import zlib
from pathlib import Path

import openpyxl
import pyexpat
from openpyxl.worksheet._read_only import ReadOnlyWorksheet

from .config import AppConfig, Limits, SourceProfile
from .models import (
    DETAIL_CODES,
    EnvironmentReport,
    FileOutcome,
    IngestedFile,
    IngestResult,
    Issue,
    IssueCode,
    IssueLevel,
    ProcessingEvent,
    Severity,
    SourceRecord,
)

__all__ = ["ingest", "environment_report", "EnvironmentCheckError"]

CHUNK = 65_536
PROLOGUE_MAX = 65_536
MAX_EXCEL_ROWS = 1_048_576
MAX_EXCEL_COLS = 16_384
_MB = 1 << 20
_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_T_WS, _T_SST, _T_SI = f"{{{_MAIN_NS}}}worksheet", f"{{{_MAIN_NS}}}sst", f"{{{_MAIN_NS}}}si"
_T_SHEETDATA, _T_ROW, _T_C = f"{{{_MAIN_NS}}}sheetData", f"{{{_MAIN_NS}}}row", f"{{{_MAIN_NS}}}c"
_SST_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
_CT_PART, _WORKBOOK_PART = "[Content_Types].xml", "xl/workbook.xml"
_FIXED_XML = frozenset(
    {_CT_PART, "docProps/core.xml", "docProps/custom.xml", "docProps/app.xml", "xl/styles.xml", "xl/theme/theme1.xml"}
)
_BINARY_REL_TYPES = frozenset(
    {"image", "oleObject", "package", "printerSettings", "vbaProject", "audio", "video", "font"}
)


class EnvironmentCheckError(RuntimeError):
    """DT4-18: el entorno no cumple los requisitos (fallo cerrado, antes de leer ninguna fuente)."""

    def __init__(self, message: str, report: EnvironmentReport) -> None:
        super().__init__(message)
        self.report = report


class _Rejection(Exception):
    """Rechazo de un archivo: la captura un único punto por archivo (Plan §7.6)."""

    def __init__(self, code: IssueCode, detail_code: str | None, message: str) -> None:
        if detail_code is not None and detail_code not in DETAIL_CODES:
            raise ValueError(f"detail_code fuera del catálogo cerrado: {detail_code}")
        super().__init__(message)
        self.code, self.detail_code, self.message = code, detail_code, message


def _f01(detail: str, message: str) -> _Rejection:
    return _Rejection(IssueCode.F01_UNREADABLE_FILE, detail, message)


def _f04(detail: str, message: str) -> _Rejection:
    return _Rejection(IssueCode.F04_LIMIT_EXCEEDED, detail, message)


# ---------------------------------------------------------------------------------------------------
# Entorno (I1, DT4-18)
# ---------------------------------------------------------------------------------------------------


def _probe_deferral() -> bool:
    try:
        parser = pyexpat.ParserCreate("UTF-8")
        setter = getattr(parser, "SetReparseDeferralEnabled", None)
        getter = getattr(parser, "GetReparseDeferralEnabled", None)
        if not (callable(setter) and callable(getter)):
            return False
        setter(True)
        return getter() is True
    except Exception:  # noqa: BLE001 — cualquier fallo es fallo cerrado
        return False


def _new_expat_parser():
    """Fábrica única de parsers Expat: exige el aplazamiento de reanálisis en ESA instancia (DT4-18)."""
    parser = pyexpat.ParserCreate("UTF-8")
    ok = False
    try:
        setter = getattr(parser, "SetReparseDeferralEnabled", None)
        getter = getattr(parser, "GetReparseDeferralEnabled", None)
        if callable(setter) and callable(getter):
            setter(True)
            ok = getter() is True
    except Exception:  # noqa: BLE001
        ok = False
    if not ok:
        raise EnvironmentCheckError(
            "no se pudo confirmar el aplazamiento de reanálisis de Expat; ejecución abortada", environment_report()
        )
    return parser


def environment_report() -> EnvironmentReport:
    info = tuple(pyexpat.version_info)
    return EnvironmentReport(
        python_version=sys.version,
        openpyxl_version=openpyxl.__version__,
        expat_version=pyexpat.EXPAT_VERSION,
        expat_version_info=(int(info[0]), int(info[1]), int(info[2])),
        openpyxl_lxml=bool(openpyxl.LXML),
        openpyxl_defusedxml=bool(openpyxl.DEFUSEDXML),
        expat_min_ok=info >= (2, 4, 1),
        expat_reparse_deferral=_probe_deferral(),
    )


# ---------------------------------------------------------------------------------------------------
# Lectura acotada (X1/X2)
# ---------------------------------------------------------------------------------------------------


def _read_bounded(path: Path, max_bytes: int) -> bytes:
    try:
        fd = os.open(path, os.O_RDONLY | os.O_NOFOLLOW | os.O_NONBLOCK | os.O_CLOEXEC)
    except OSError as exc:
        why = "enlace simbólico introducido tras la clasificación" if exc.errno == errno.ELOOP else exc.strerror
        raise _f01("io_error", f"no se pudo abrir el archivo: {why}") from None
    try:
        with os.fdopen(fd, "rb", buffering=0) as handle:
            st = os.fstat(handle.fileno())
            if not stat.S_ISREG(st.st_mode):
                raise _f01("not_regular_file", "la entrada no es un archivo regular")
            if st.st_size > max_bytes:
                raise _f04("file_size", f"el archivo supera el máximo de {max_bytes} bytes")
            parts: list[bytes] = []
            total = 0
            while total < max_bytes + 1:
                block = handle.read(min(_MB, max_bytes + 1 - total))
                if not block:
                    break
                parts.append(block)
                total += len(block)
            if total > max_bytes:
                raise _f04("file_size", f"el archivo supera el máximo de {max_bytes} bytes")
            return b"".join(parts)
    except OSError as exc:
        raise _f01("io_error", f"error de lectura: {exc.strerror}") from None


# ---------------------------------------------------------------------------------------------------
# Clasificación de filas físicas (común a CSV y XLSX, §7.5)
# ---------------------------------------------------------------------------------------------------


def _is_blank(values: tuple) -> bool:
    return all(v is None or (isinstance(v, str) and v.strip() == "") for v in values)


class _Rows:
    """Acumula filas físicas, las clasifica y aplica los límites de registros (Plan §9.1 paso 8, §10.10)."""

    def __init__(self, profile: SourceProfile, rel_name: str, sheet: str, limits: Limits, budget_used: int):
        self.profile, self.rel_name, self.sheet = profile, rel_name, sheet
        self.limits, self.budget_used = limits, budget_used
        self.header: SourceRecord | None = None
        self.records: list[SourceRecord] = []
        self.physical = self.preamble = self.header_rows = self.blank = 0

    def add(self, row: int, line: int | None, values: tuple, types: tuple[str, ...] | None) -> None:
        self.physical = row
        rid = f"{self.profile.id}:{self.rel_name}:{self.sheet}:{row}"
        rec = SourceRecord(rid, self.profile.id, self.rel_name, self.sheet, row, line, values, types)
        if row < self.profile.header_row:
            self.preamble += 1
        elif row == self.profile.header_row:
            self.header, self.header_rows = rec, 1
        elif _is_blank(values):
            self.blank += 1
        else:
            self.records.append(rec)
            if len(self.records) > self.limits.max_records_per_file:
                raise _f04("records_file", "registros por archivo por encima del máximo")
            if self.budget_used + len(self.records) > self.limits.max_records_per_run:
                raise _f04("records_run", "presupuesto de registros de la ejecución agotado")

    def finish(self, sha: str, sheet: str | None, ignored: tuple[str, ...] = ()) -> IngestedFile:
        return IngestedFile(
            self.rel_name, self.profile, sha, sheet, self.header, tuple(self.records),
            self.physical, self.preamble, self.header_rows, self.blank, ignored,
        )


# ---------------------------------------------------------------------------------------------------
# CSV (§11.1.2)
# ---------------------------------------------------------------------------------------------------


def _read_csv(buffer: bytes, profile: SourceProfile, rel_name: str, limits: Limits, used: int, sha: str) -> IngestedFile:
    try:
        text = buffer.decode(profile.encoding or "utf-8", errors="strict")
    except UnicodeDecodeError as exc:
        raise _Rejection(
            IssueCode.F02_ENCODING_ERROR, None, f"no se puede decodificar con {profile.encoding}: {exc.reason}"
        ) from None
    rows = _Rows(profile, rel_name, "", limits, used)
    previous = csv.field_size_limit(limits.max_text_chars)
    try:
        reader = csv.reader(io.StringIO(text, newline=""), delimiter=profile.delimiter, strict=True)
        prev_line = 0
        for row, fields in enumerate(reader, start=1):
            line = prev_line + 1
            prev_line = reader.line_num
            rows.add(row, line, tuple(fields), None)
    except csv.Error as exc:
        if str(exc).startswith("field larger than field limit"):
            raise _f04("text_chars", f"campo CSV por encima de {limits.max_text_chars} caracteres") from None
        raise _f01("csv_malformed", f"CSV mal formado: {exc}") from None
    finally:
        csv.field_size_limit(previous)
    return rows.finish(sha, None)


# ---------------------------------------------------------------------------------------------------
# XLSX: X2b — precomprobación EOCD/ZIP64 y recorrido acotado del directorio central (DT4-23)
# ---------------------------------------------------------------------------------------------------


def _precheck_zip(buf: bytes, max_entries: int) -> None:
    def bad(why: str) -> _Rejection:
        return _f01("zip_unreadable", f"contenedor ZIP ilegible: {why}")

    n = len(buf)
    if n >= 22 and buf[-22:-18] == b"PK\x05\x06" and buf[-2:] == b"\x00\x00":
        pos = n - 22
    else:
        pos = buf.rfind(b"PK\x05\x06", max(0, n - (65_535 + 22)))
        if pos < 0 or pos + 22 > n:
            raise bad("sin registro de fin de directorio")
    _, _, _, _, _, size_cd, off_cd, _ = struct.unpack_from("<4s4H2LH", buf, pos)
    zip64 = False
    if pos >= 20 and buf[pos - 20 : pos - 16] == b"PK\x06\x07":
        _, diskno, _, disks = struct.unpack_from("<4sLQL", buf, pos - 20)
        if diskno != 0 or disks > 1:
            raise bad("archivos ZIP de varios discos")
        rec = pos - 20 - 56
        if rec >= 0 and buf[rec : rec + 4] == b"PK\x06\x06":
            zip64 = True
            _, _, _, _, _, _, _, _, size_cd, off_cd = struct.unpack_from("<4sQ2H2L4Q", buf, rec)
    shift = 76 if zip64 else 0
    start_dir = pos - shift - size_cd
    if start_dir < 0 or pos - shift > n:
        raise bad("directorio central fuera del archivo")
    cursor, limit, count = start_dir, start_dir + size_cd, 0
    while cursor < limit:
        if cursor + 46 > limit or buf[cursor : cursor + 4] != b"PK\x01\x02":
            raise bad("registro del directorio central truncado")
        count += 1
        if count > max_entries:
            raise _f04("zip_entries", f"más de {max_entries} entradas en el ZIP")
        name_len, extra_len, comment_len = struct.unpack_from("<3H", buf, cursor + 28)
        cursor += 46 + name_len + extra_len + comment_len
    if cursor != limit:
        raise bad("el directorio central no termina donde se declara")


# ---------------------------------------------------------------------------------------------------
# XLSX: X4 / X4b / X5
# ---------------------------------------------------------------------------------------------------

_DRIVE = re.compile(r"^[A-Za-z]:")


def _unsafe_name(orig: str) -> bool:
    name = orig[:-1] if orig.endswith("/") else orig
    if not name or name[0] in "/\\" or "\\" in name or _DRIVE.match(name):
        return True
    if any(ord(ch) < 0x20 for ch in name):
        return True
    return any(seg in ("", ".", "..") for seg in name.split("/"))


def _check_directory(infos: list[zipfile.ZipInfo], max_entries: int) -> None:
    if len(infos) > max_entries:
        raise _f04("zip_entries", f"más de {max_entries} entradas en el ZIP")
    for i in infos:
        if _unsafe_name(i.orig_filename):
            raise _f01("unsafe_name", f"nombre de entrada no seguro: {i.orig_filename!r}")
    seen: set[str] = set()
    for i in infos:
        key = i.orig_filename.lower()
        if key in seen:
            raise _f01("duplicate_name", f"nombre de entrada duplicado: {i.orig_filename!r}")
        seen.add(key)
    for i in infos:
        if i.flag_bits & 0x1 or i.flag_bits & 0x40:
            raise _f01("encrypted_entry", f"entrada cifrada: {i.orig_filename!r}")
    for i in infos:
        if i.compress_type not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED) or i.flag_bits & 0x20:
            raise _f01("unsupported_method", f"método de compresión no admitido: {i.orig_filename!r}")
    names = {i.orig_filename for i in infos}
    for required in (_CT_PART, _WORKBOOK_PART):
        if required not in names:
            raise _f01("missing_part", f"falta la parte obligatoria {required}")


def _entry_ranges(buf: bytes, infos: list[zipfile.ZipInfo], cd_start: int) -> list[tuple[int, int]]:
    """X4b: rangos [data_start, data_end) validados sobre el buffer, uno por entrada."""
    out: list[tuple[int, int]] = []

    def trunc(name: str, why: str) -> _Rejection:
        return _f01("zip_truncated", f"entrada {name!r}: {why}")

    for info in infos:
        off = info.header_offset
        if off < 0 or off + 30 > cd_start:
            raise trunc(info.orig_filename, "cabecera local fuera de rango")
        sig, _v, flag, _m, _t, _d, _c, _cs, _us, n, m = struct.unpack_from("<4s5H3I2H", buf, off)
        if sig != b"PK\x03\x04":
            raise trunc(info.orig_filename, "firma de cabecera local incorrecta")
        raw = buf[off + 30 : off + 30 + n]
        try:
            local_name = raw.decode("utf-8" if flag & 0x800 else "cp437")
        except UnicodeDecodeError:
            raise trunc(info.orig_filename, "nombre local no decodificable") from None
        if len(raw) != n or local_name != info.orig_filename:
            raise trunc(info.orig_filename, "el nombre local no coincide con el del directorio central")
        start = off + 30 + n + m
        end = start + info.compress_size
        if end > cd_start:
            raise trunc(info.orig_filename, "los datos exceden el directorio central")
        out.append((start, end))
    order = sorted(range(len(infos)), key=lambda k: infos[k].header_offset)
    for a, b in zip(order, order[1:]):
        if infos[b].header_offset < out[a][1]:
            raise _f01(
                "overlapping_entries", f"entradas solapadas: {infos[a].orig_filename!r} y {infos[b].orig_filename!r}"
            )
    return out


def _check_declared(infos: list[zipfile.ZipInfo], limits: Limits) -> None:
    per_entry, total_max = limits.max_entry_uncompressed_mb * _MB, limits.max_total_uncompressed_mb * _MB
    for i in infos:
        if i.file_size > per_entry:
            raise _f04("declared_size", f"tamaño declarado de {i.orig_filename!r} por encima del máximo")
    if sum(i.file_size for i in infos) > total_max:
        raise _f04("declared_size", "suma de tamaños declarados por encima del máximo")
    for i in infos:
        if i.file_size and (i.compress_size == 0 or i.file_size / i.compress_size > limits.max_declared_compression_ratio):
            raise _f04("declared_ratio", f"relación de compresión declarada excesiva en {i.orig_filename!r}")


# ---------------------------------------------------------------------------------------------------
# XLSX: X6 — controles XML sobre la salida real (prólogo/encoding, DOCTYPE, pre-escaneos)
# ---------------------------------------------------------------------------------------------------

_ACCEPT, _REJECT, _WAIT = "accept", "reject", "wait"
_UTF8_BOM = b"\xef\xbb\xbf"
_ENC_ATTR = re.compile(rb"""\sencoding\s*=\s*(?:"([^"]*)"|'([^']*)')""", re.IGNORECASE)
_XML_SPACE = b" \t\r\n"


def _decide_prologue(pro: bytes, fin: bool) -> str:
    """Decisión de encoding (Plan §10.7.3, R1–R7). Solo se acepta UTF-8 (BOM opcional)."""
    if pro.startswith((b"\x00\x00\xfe\xff", b"\xff\xfe", b"\xfe\xff")):
        return _REJECT
    c = pro[3:] if pro.startswith(_UTF8_BOM) else pro
    if c.startswith((b"<\x00", b"\x00<", b"\x00\x00\x00<")):
        return _REJECT
    if not fin:
        r2 = (b"<\x00", b"\x00<", b"\x00\x00\x00<")
        r1 = (b"\x00\x00\xfe\xff", b"\xff\xfe", b"\xfe\xff", _UTF8_BOM)
        if any(len(pro) < len(p) and p.startswith(pro) for p in r1):
            return _WAIT
        if any(len(c) < len(p) and p.startswith(c) for p in r2) or (len(c) <= 5 and b"<?xml".startswith(c)):
            return _WAIT
    if not (c.startswith(b"<?xml") and c[5:6] and c[5:6] in _XML_SPACE):
        return _ACCEPT
    bom = len(pro) - len(c)
    end = c.find(b"?>")
    if end < 0 or bom + end + 2 > PROLOGUE_MAX:
        return _REJECT if (len(pro) >= PROLOGUE_MAX or fin) else _WAIT
    decl = c[: end + 2]
    count = decl.lower().count(b"encoding")
    if count == 0:
        return _ACCEPT
    m = _ENC_ATTR.search(decl)
    if count == 1 and m:
        value = m.group(1) if m.group(1) is not None else m.group(2)
        if value.lower() == b"utf-8":
            return _ACCEPT
    return _REJECT


class _DoctypeFound(Exception):
    pass


class _XmlPart:
    """Controles de una parte del conjunto XML. Orden por salida: prólogo → centinela DOCTYPE → pre-escaneos."""

    def __init__(self, name: str, kind: str, limits: Limits, sst_declared: bool):
        self.name, self.kind, self.limits = name, kind, limits  # kind: ct | rels | generic
        self.sst_declared = sst_declared
        self.pending: bytearray | None = bytearray()
        self.sentinel = _new_expat_parser()
        self.sentinel.StartDoctypeDeclHandler = self._on_doctype
        self.pull = ET.XMLPullParser(events=("start", "end"))
        self.root: ET.Element | None = None
        self.mode = kind
        self.sheet_data: ET.Element | None = None
        self.si_count = self.si_chars = 0
        self.last_row = self.last_col = 0
        self.defaults: list[tuple[str, str]] = []
        self.overrides: list[tuple[str, str]] = []
        self.rels: list[tuple[str, str, str]] = []  # (target, type, mode)

    @staticmethod
    def _on_doctype(*_args) -> None:
        raise _DoctypeFound

    def _malformed(self, why: str) -> _Rejection:
        return _f01("xml_malformed", f"XML mal formado en {self.name!r}: {why}")

    def feed(self, chunk: bytes) -> None:
        if self.pending is not None:
            self.pending += chunk
            verdict = _decide_prologue(bytes(self.pending), False)
            if verdict == _WAIT:
                return
            self._decide(verdict)
        else:
            self._deliver(chunk)

    def _decide(self, verdict: str) -> None:
        if verdict == _REJECT:
            raise _f01("xml_encoding", f"codificación no admitida (solo UTF-8) en {self.name!r}")
        block, self.pending = bytes(self.pending or b""), None
        self._deliver(block)

    def _deliver(self, block: bytes) -> None:
        try:
            self.sentinel.Parse(block, False)
        except _DoctypeFound:
            raise _f01("xml_doctype", f"declaración DOCTYPE no admitida en {self.name!r}") from None
        except pyexpat.ExpatError as exc:
            raise self._malformed(str(exc)) from None
        try:
            self.pull.feed(block)
        except ET.ParseError as exc:
            raise self._malformed(str(exc)) from None
        self._events()

    def finish(self) -> None:
        if self.pending is not None:
            self._decide(_decide_prologue(bytes(self.pending), True))
        try:
            self.sentinel.Parse(b"", True)
        except _DoctypeFound:
            raise _f01("xml_doctype", f"declaración DOCTYPE no admitida en {self.name!r}") from None
        except pyexpat.ExpatError as exc:
            raise self._malformed(str(exc)) from None
        try:
            self.pull.close()
        except ET.ParseError as exc:
            raise self._malformed(str(exc)) from None
        self._events()

    # -- pre-escaneos ---------------------------------------------------------------------------
    def _events(self) -> None:
        for event, el in self.pull.read_events():
            if event == "start":
                if self.root is None:
                    self.root = el
                    if self.mode == "generic":
                        if el.tag == _T_WS:
                            self.mode = "ws"
                        elif el.tag == _T_SST or self.sst_declared:
                            self.mode = "sst"
                        else:
                            self.mode = "plain"
                if self.mode == "ws":
                    self._ws_start(el)
            else:
                self._end(el)

    def _end(self, el: ET.Element) -> None:
        mode = self.mode
        if mode == "ct":
            if el.tag.endswith("}Default"):
                self.defaults.append((el.get("Extension", ""), el.get("ContentType", "")))
            elif el.tag.endswith("}Override"):
                self.overrides.append((el.get("PartName", ""), el.get("ContentType", "")))
        elif mode == "rels":
            if el.tag.endswith("}Relationship"):
                self.rels.append((el.get("Target", ""), el.get("Type", ""), el.get("TargetMode", "")))
        elif mode == "sst":
            if el.tag == _T_SI:
                self.si_count += 1
                if self.si_count > self.limits.max_shared_strings:
                    raise _f04("shared_strings_count", "demasiadas cadenas compartidas (<si>)")
                self.si_chars += sum(len(t) for t in el.itertext())
                if self.si_chars > self.limits.max_shared_strings_chars:
                    raise _f04("shared_strings_chars", "demasiados caracteres en las cadenas compartidas")
                el.clear()
                if self.root is not None:
                    try:
                        self.root.remove(el)
                    except ValueError:
                        pass
        elif mode == "ws":
            self._ws_end(el)
        elif el is not self.root:
            el.clear()

    def _ws_start(self, el: ET.Element) -> None:
        if el.tag == _T_SHEETDATA:
            self.sheet_data = el
        elif el.tag == _T_ROW:
            ref = el.get("r")
            try:
                row = int(ref) if ref is not None else self.last_row + 1
            except ValueError:
                raise self._malformed("índice de fila no numérico") from None
            if not 1 <= row <= MAX_EXCEL_ROWS:
                raise self._malformed("índice de fila fuera de rango")
            if row <= self.last_row:
                raise self._malformed("filas duplicadas o fuera de orden")
            self.last_row, self.last_col = row, 0

    def _ws_end(self, el: ET.Element) -> None:
        if el.tag == _T_C:
            ref = el.get("r")
            if ref:
                m = re.fullmatch(r"([A-Za-z]{1,3})([0-9]{1,7})", ref)
                if not m:
                    raise self._malformed("coordenada de celda inválida")
                col = 0
                for ch in m.group(1).upper():
                    col = col * 26 + ord(ch) - 64
                if int(m.group(2)) != self.last_row:
                    raise self._malformed("la celda no coincide con su fila")
            else:
                col = self.last_col + 1
            if not 1 <= col <= MAX_EXCEL_COLS or col <= self.last_col:
                raise self._malformed("celdas duplicadas, fuera de orden o fuera de rango")
            self.last_col = col
        elif el.tag == _T_ROW:
            el.clear()
            if self.sheet_data is not None:
                try:
                    self.sheet_data.remove(el)
                except ValueError:
                    pass


def _is_xml_content_type(ct: str) -> bool:
    return ct.endswith("+xml") or ct in ("application/xml", "text/xml")


def _resolve_target(rels_name: str, target: str) -> str:
    folder = posixpath.dirname(posixpath.dirname(rels_name))  # padre de la carpeta _rels
    if target.startswith("/"):
        return posixpath.normpath(target[1:])
    return posixpath.normpath(posixpath.join(folder, target))


class _Budget:
    real_total = 0


def _preflight(buf: bytes, infos: list[zipfile.ZipInfo], ranges: list[tuple[int, int]], limits: Limits) -> None:
    """X6: descompresión real acotada de TODAS las entradas, CRC, integridad y controles XML (DT4-28/27/29)."""
    view = memoryview(buf)
    entry_max, total_max = limits.max_entry_uncompressed_mb * _MB, limits.max_total_uncompressed_mb * _MB
    budget = _Budget()

    def phase(k: int) -> int:
        n = infos[k].orig_filename
        return 0 if n == _CT_PART else 1 if n.endswith(".rels") else 2

    order = sorted(range(len(infos)), key=lambda k: (phase(k), k))
    xml_names: set[str] = set(_FIXED_XML)
    sst_parts: set[str] = set()
    default_xml_ext: set[str] = set()
    rel_targets: dict[str, list[str]] = {}
    s4_done = False
    for idx in order:
        info = infos[idx]
        name = info.orig_filename
        if phase(idx) == 2 and not s4_done:  # S4: destinos de relación, ya conocidas todas las .rels
            s4_done = True
            for target, types in rel_targets.items():
                if target in xml_names or not all(t in _BINARY_REL_TYPES for t in types):
                    xml_names.add(target)
        is_dir = name.endswith("/")
        ext = posixpath.splitext(name)[1].lstrip(".").lower()
        in_set = not is_dir and (
            phase(idx) < 2 or name.endswith(".xml") or name in xml_names or (ext != "" and ext in default_xml_ext)
        )
        part: _XmlPart | None = None
        if in_set:
            kind = "ct" if name == _CT_PART else "rels" if name.endswith(".rels") else "generic"
            part = _XmlPart(name, kind, limits, name in sst_parts)
        start, end = ranges[idx]
        real, crc = _inflate(info, view[start:end], part, entry_max, total_max, budget)
        if is_dir and (info.file_size != 0 or real != 0):
            raise _f01(
                "unsafe_name",
                f"entrada de directorio con contenido: {name!r} (file_size declarado {info.file_size}, real {real})",
            )
        if real != info.file_size:
            raise _f01("crc_mismatch", f"tamaño real ({real}) distinto del declarado ({info.file_size}) en {name!r}")
        if crc != info.CRC:
            raise _f01("crc_mismatch", f"CRC32 incorrecto en {name!r}")
        if part is None:
            continue
        part.finish()
        if part.kind == "ct":
            for pext, ctype in part.defaults:
                if _is_xml_content_type(ctype):
                    default_xml_ext.add(pext.lower())
            for pname, ctype in part.overrides:
                pname = pname.lstrip("/")
                if _is_xml_content_type(ctype):
                    xml_names.add(pname)
                if ctype == _SST_CT:
                    sst_parts.add(pname)
        elif part.kind == "rels":
            for target, rtype, mode in part.rels:
                if mode != "External" and target:
                    rel_targets.setdefault(_resolve_target(name, target), []).append(rtype.rsplit("/", 1)[-1])


def _inflate(
    info: zipfile.ZipInfo, comp: memoryview, part: _XmlPart | None, entry_max: int, total_max: int, budget: _Budget
) -> tuple[int, int]:
    real, crc = 0, 0

    def consume(out: bytes) -> None:
        nonlocal real, crc
        real += len(out)
        budget.real_total += len(out)
        if real > entry_max:
            raise _f04("entry_uncompressed", f"bytes reales de {info.orig_filename!r} por encima del máximo por entrada")
        if budget.real_total > total_max:
            raise _f04("total_uncompressed", "bytes reales acumulados por encima del máximo total")
        crc = zlib.crc32(out, crc)
        if part is not None:
            part.feed(out)

    if info.compress_type == zipfile.ZIP_STORED:
        for off in range(0, len(comp), CHUNK):
            consume(bytes(comp[off : off + CHUNK]))
        return real, crc
    d = zlib.decompressobj(wbits=-15)
    pos = 0
    while True:
        if d.unconsumed_tail:
            inp, source = d.unconsumed_tail, "tail"
        elif pos < len(comp):
            inp, source = comp[pos : pos + CHUNK], "new"
            pos += len(inp)
        else:
            inp, source = b"", "drain"
        room = min(entry_max - real, total_max - budget.real_total)
        try:
            out = d.decompress(inp, min(CHUNK, room + 1))
        except zlib.error:
            raise _f01("crc_mismatch", f"flujo DEFLATE corrupto en {info.orig_filename!r}") from None
        consume(out)
        if d.eof:
            break
        if not out and (source == "drain" or len(inp) - len(d.unconsumed_tail) == 0):
            raise _f01("zip_truncated", f"flujo DEFLATE incompleto en {info.orig_filename!r}")
    if len(d.unused_data) + (len(comp) - pos) > 0:  # DT4-28 (aclaración): sin usar unconsumed_tail tras EOF
        raise _f01("zip_unreadable", f"datos comprimidos sobrantes tras el final del flujo en {info.orig_filename!r}")
    return real, crc


# ---------------------------------------------------------------------------------------------------
# XLSX: orquestación X2b–X8
# ---------------------------------------------------------------------------------------------------


def _read_xlsx(
    buffer: bytes, profile: SourceProfile, rel_name: str, limits: Limits, used: int, sha: str
) -> tuple[IngestedFile, list[str]]:
    _precheck_zip(buffer, limits.max_zip_entries)
    try:
        with zipfile.ZipFile(io.BytesIO(buffer)) as zf:
            infos = list(zf.infolist())
            cd_start = zf.start_dir
    except (zipfile.BadZipFile, zipfile.LargeZipFile, NotImplementedError, EOFError, OSError, ValueError) as exc:
        raise _f01("zip_unreadable", f"contenedor ZIP ilegible: {exc}") from None
    _check_directory(infos, limits.max_zip_entries)
    ranges = _entry_ranges(buffer, infos, cd_start)
    _check_declared(infos, limits)
    _preflight(buffer, infos, ranges, limits)
    return _iterate_workbook(buffer, profile, rel_name, limits, used, sha)


def _iterate_workbook(
    buffer: bytes, profile: SourceProfile, rel_name: str, limits: Limits, used: int, sha: str
) -> tuple[IngestedFile, list[str]]:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            wb = openpyxl.load_workbook(io.BytesIO(buffer), read_only=True, data_only=True)
        except (ET.ParseError, ValueError, KeyError, IndexError, TypeError) as exc:
            raise _f01("xml_malformed", f"openpyxl no pudo cargar el libro: {exc}") from None
        try:
            names = list(wb.sheetnames)
            if len(names) > limits.max_sheets:
                raise _f04("sheets", f"más de {limits.max_sheets} hojas en el libro")
            if profile.sheet not in names:
                raise _Rejection(IssueCode.F07_SHEET_NOT_FOUND, None, f"la hoja {profile.sheet!r} no existe en el libro")
            ws = wb[profile.sheet]
            if not isinstance(ws, ReadOnlyWorksheet):
                raise _Rejection(IssueCode.F07_SHEET_NOT_FOUND, None, f"{profile.sheet!r} no es una hoja de datos")
            ignored = [n for n in names if n != profile.sheet]
            rows = _Rows(profile, rel_name, profile.sheet, limits, used)
            ws.reset_dimensions()
            it = ws.iter_rows()
            cells = 0
            try:
                for physical, row in enumerate(it, start=1):
                    if len(row) > limits.max_columns:
                        raise _f04("columns", f"fila {physical} con más de {limits.max_columns} columnas")
                    cells += len(row)
                    if cells > limits.max_cells_per_file:
                        raise _f04("cells", "celdas iteradas por encima del máximo")
                    for c in row:
                        v = c.value
                        if isinstance(v, str) and len(v) > limits.max_text_chars:
                            raise _f04("text_chars", f"texto de celda por encima de {limits.max_text_chars} caracteres")
                        if v is not None and getattr(c, "row", physical) != physical:
                            raise _f01("xml_malformed", "la fila de la celda no coincide con la posición")
                    rows.add(physical, None, tuple(c.value for c in row), tuple(c.data_type for c in row))
            finally:
                it.close()
        except (ET.ParseError, KeyError, IndexError, TypeError) as exc:
            raise _f01("xml_malformed", f"error al leer la hoja: {exc}") from None
        finally:
            wb.close()
    return rows.finish(sha, profile.sheet, tuple(ignored)), ignored


# ---------------------------------------------------------------------------------------------------
# Descubrimiento y orquestación (I2–I10)
# ---------------------------------------------------------------------------------------------------


def _issue(source_id: str | None, rel_name: str, rej: _Rejection) -> Issue:
    return Issue(
        level=IssueLevel.FILE,
        source_id=source_id,
        source_file=rel_name,
        issue_code=rej.code,
        detail_code=rej.detail_code,
        severity=Severity.ERROR,
        message=rej.message,
    )


def _event(level: str, event: str, **kw) -> ProcessingEvent:
    return ProcessingEvent(level=level, event=event, **kw)


def ingest(input_dir: Path | str, config: AppConfig) -> IngestResult:
    """Ingiere el directorio (no recursivo). Un archivo rechazado no detiene a los demás; el entorno sí (DT4-18)."""
    env = environment_report()
    events: list[ProcessingEvent] = [
        _event("run", "environment", metric=m, value=str(v))
        for m, v in (
            ("python_version", env.python_version),
            ("openpyxl_version", env.openpyxl_version),
            ("expat_version", env.expat_version),
            ("openpyxl_lxml", env.openpyxl_lxml),
            ("openpyxl_defusedxml", env.openpyxl_defusedxml),
            ("expat_min_ok", env.expat_min_ok),
            ("expat_reparse_deferral", env.expat_reparse_deferral),
        )
    ]
    if not env.expat_min_ok or env.openpyxl_lxml or not env.expat_reparse_deferral:
        raise EnvironmentCheckError("el entorno no cumple los requisitos de seguridad XML de F4", env)

    limits = config.limits
    outcomes: list[FileOutcome] = []
    issues: list[Issue] = []
    used = 0
    root = Path(input_dir)
    with os.scandir(root) as it:
        entries = sorted(it, key=lambda e: e.name)

    for entry in entries:
        name = entry.name
        ignored = None
        if entry.is_symlink():
            ignored = "symlink"
        elif name.startswith("."):
            ignored = "hidden"
        elif name.startswith("~$"):
            ignored = "temporary"
        elif entry.is_dir(follow_symlinks=False):
            continue  # DT4-04: los subdirectorios no se descubren
        if ignored:
            events.append(_event("file", "file_ignored", source_file=name, metric="reason", value=ignored))
            outcomes.append(FileOutcome(name, "ignored", ignore_reason=ignored))
            continue

        source_id: str | None = None
        try:
            if not entry.is_file(follow_symlinks=False):
                raise _f01("not_regular_file", "la entrada no es un archivo regular")
            ext = os.path.splitext(name)[1]
            if ext not in (".xlsx", ".csv"):
                raise _Rejection(IssueCode.F03_UNSUPPORTED_TYPE, None, f"extensión no admitida: {ext or '(ninguna)'}")
            matches = [p for p in config.sources if fnmatch.fnmatchcase(name, p.match)]
            if not matches:
                raise _Rejection(IssueCode.F05_NO_PROFILE, None, "ningún perfil coincide con el archivo")
            if len(matches) > 1:
                ids = ", ".join(p.id for p in matches)
                raise _Rejection(IssueCode.F06_AMBIGUOUS_PROFILE, None, f"varios perfiles coinciden: {ids}")
            profile = matches[0]
            source_id = profile.id
            if ext != f".{profile.type}":
                raise _Rejection(
                    IssueCode.F03_UNSUPPORTED_TYPE, None, f"la extensión {ext} no corresponde al tipo {profile.type} del perfil"
                )
            events.append(_event("file", "profile_applied", source_id=source_id, source_file=name, metric="source_id", value=source_id))

            buffer = _read_bounded(Path(entry.path), limits.max_file_mb * _MB)
            sha = hashlib.sha256(buffer).hexdigest()
            events.append(_event("file", "input_sha256", source_id=source_id, source_file=name, metric="sha256", value=sha))
            if profile.type == "csv":
                done = _read_csv(buffer, profile, name, limits, used, sha)
                ignored_sheets: list[str] = []
            else:
                done, ignored_sheets = _read_xlsx(buffer, profile, name, limits, used, sha)
        except _Rejection as rej:
            issue = _issue(source_id, name, rej)
            issues.append(issue)
            outcomes.append(FileOutcome(name, "rejected", source_id=source_id, issue=issue))
            continue

        used += len(done.records)
        for sheet in ignored_sheets:
            events.append(_event("file", "sheet_ignored", source_id=source_id, source_file=name, metric="sheet", value=sheet))
        counts = (
            ("physical_rows", done.physical_rows), ("preamble_rows", done.preamble_rows),
            ("header_rows", done.header_rows), ("blank_rows", done.blank_rows), ("records_read", len(done.records)),
        )
        for metric, value in counts:
            events.append(_event("file", "rows_counted", source_id=source_id, source_file=name, metric=metric, value=str(value)))
        outcomes.append(FileOutcome(name, "ingested", source_id=source_id, ingested=done))

    return IngestResult(env, tuple(outcomes), tuple(events), tuple(issues))
