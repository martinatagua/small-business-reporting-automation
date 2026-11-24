"""Pruebas de F4 (ingesta). Los casos adversariales se construyen en memoria con datos sintéticos."""

from __future__ import annotations

import csv
import dataclasses
import io
import os
import random
import struct
import zipfile
import zlib
from pathlib import Path

import openpyxl
import pytest

from small_business_reporting_automation import ingest as ingest_mod
from small_business_reporting_automation.config import AppConfig, Limits, load_config
from small_business_reporting_automation.ingest import EnvironmentCheckError, environment_report, ingest
from small_business_reporting_automation.models import DETAIL_CODES, IssueCode

_MB = 1 << 20
REPO = Path(__file__).resolve().parents[2]
CFG = load_config(REPO / "config" / "demo.toml")
XLSX_PROFILE = dataclasses.replace(next(p for p in CFG.sources if p.id == "sur-xlsx"), match="*.xlsx")
CSV_PROFILE = dataclasses.replace(next(p for p in CFG.sources if p.id == "sur-csv"), match="*.csv")


def cfg(*profiles, **limits) -> AppConfig:
    return dataclasses.replace(CFG, sources=tuple(profiles), limits=dataclasses.replace(CFG.limits, **limits))


def run_one(tmp_path: Path, name: str, data: bytes, config: AppConfig):
    (tmp_path / name).write_bytes(data)
    res = ingest(tmp_path, config)
    (out,) = [o for o in res.files if o.source_file == name]
    return res, out


def rejected(out, code: IssueCode, detail: str | None):
    assert out.outcome == "rejected", out
    assert out.issue.issue_code is code and out.issue.detail_code == detail, out.issue
    assert detail is None or detail in DETAIL_CODES


# ---- constructores de XLSX/ZIP ------------------------------------------------------------------


def good_xlsx(rows=(("a", "b"), (1, 2), (3, 4)), title="Datos") -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title
    for r in rows:
        ws.append(list(r))
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def rezip(data: bytes, mutate=None, extra=(), compress=zipfile.ZIP_DEFLATED) -> bytes:
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        parts = {i.filename: zf.read(i.filename) for i in zf.infolist()}
    if mutate:
        parts = mutate(parts)
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w", compress) as zf:
        for n, b in list(parts.items()) + list(extra):
            zf.writestr(n, b)
    return bio.getvalue()


def patch_field(data: bytes, name: str, offset_local: int, offset_central: int, fmt: str, value: int) -> bytes:
    """Cambia un campo numérico de la cabecera local y del directorio central de una entrada."""
    buf = bytearray(data)
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        info = zf.getinfo(name)
    struct.pack_into(fmt, buf, info.header_offset + offset_local, value)
    key = b"PK\x01\x02"
    pos = 0
    while (pos := data.find(key, pos)) >= 0:
        nlen = struct.unpack_from("<H", data, pos + 28)[0]
        if data[pos + 46 : pos + 46 + nlen] == name.encode():
            struct.pack_into(fmt, buf, pos + offset_central, value)
            return bytes(buf)
        pos += 4
    raise AssertionError("entrada no encontrada")


# ---- entorno y modelos -----------------------------------------------------------------------------


def test_environment_report_matches_runtime():
    import pyexpat

    env = environment_report()
    assert env.expat_version == pyexpat.EXPAT_VERSION and env.expat_min_ok and env.expat_reparse_deferral
    assert env.openpyxl_version == openpyxl.__version__ and env.openpyxl_lxml is False


def test_environment_failure_aborts_before_reading(tmp_path, monkeypatch):
    (tmp_path / "x.csv").write_text("a\n")
    monkeypatch.setattr(ingest_mod, "_probe_deferral", lambda: False)
    with pytest.raises(EnvironmentCheckError):
        ingest(tmp_path, cfg(CSV_PROFILE))


def test_detail_code_catalog_is_closed_and_has_29():
    assert len(DETAIL_CODES) == 29 and {"not_regular_file", "io_error"} <= DETAIL_CODES
    with pytest.raises(ValueError):
        ingest_mod._f01("inventado", "x")


# ---- demostración extremo a extremo (ground truth de F3) ------------------------------------------------


def test_demo_data_matches_ground_truth():
    res = ingest(REPO / "data" / "raw", CFG)
    gt_files = {r["source_file"]: r for r in csv.DictReader(open(REPO / "data/ground_truth/expected_files.csv"))}
    got = {o.source_file: o for o in res.files if o.outcome == "ingested"}
    assert set(got) == set(gt_files)
    for name, exp in gt_files.items():
        f = got[name].ingested
        assert (f.preamble_rows, f.header_rows, f.blank_rows, len(f.records)) == (
            int(exp["preamble_rows"]), int(exp["header_rows"]), int(exp["blank_rows"]), int(exp["expected_records"]))
        assert got[name].source_id == exp["source_id"] and len(f.sha256) == 64
    gt = {(r["source_file"], int(r["source_row"])): r for r in csv.DictReader(open(REPO / "data/ground_truth/expected_records.csv"))}
    seen = {(o.source_file, r.source_row): r for o in got.values() for r in o.ingested.records}
    assert set(seen) == set(gt)
    for key, rec in seen.items():
        exp = gt[key]
        assert rec.source_sheet == exp["source_sheet"] and (rec.source_line or "") == int(exp["source_line"] or 0) or (
            rec.source_line is None and exp["source_line"] == "")
        assert rec.record_id == f"{rec.source_id}:{rec.source_file}:{rec.source_sheet}:{rec.source_row}"
    assert res.issues == () and [e.event for e in res.events if e.event == "environment"]


def test_ingest_is_deterministic():
    a, b = ingest(REPO / "data" / "raw", CFG), ingest(REPO / "data" / "raw", CFG)
    assert a == b


# ---- descubrimiento y perfiles --------------------------------------------------------------------------


def test_discovery_classification(tmp_path):
    real = tmp_path / "a.csv"
    real.write_text("h\n1\n")
    (tmp_path / "b_link.csv").symlink_to(real)
    (tmp_path / ".oculto.csv").write_text("h\n")
    (tmp_path / "~$temp.csv").write_text("h\n")
    (tmp_path / "subdir.csv").mkdir()
    (tmp_path / "nota.txt").write_text("x")
    (tmp_path / "ninguno.csv").write_text("h\n")
    p = dataclasses.replace(CSV_PROFILE, match="a.csv", header_row=1)
    res = ingest(tmp_path, cfg(p))
    by = {o.source_file: o for o in res.files}
    assert by["a.csv"].outcome == "ingested"
    assert {k: by[k].ignore_reason for k in ("b_link.csv", ".oculto.csv", "~$temp.csv")} == {
        "b_link.csv": "symlink", ".oculto.csv": "hidden", "~$temp.csv": "temporary"}
    assert "subdir.csv" not in by  # los directorios no se descubren (DT4-04)
    rejected(by["nota.txt"], IssueCode.F03_UNSUPPORTED_TYPE, None)
    rejected(by["ninguno.csv"], IssueCode.F05_NO_PROFILE, None)
    assert [o.source_file for o in res.files] == sorted(by)  # orden determinista


def test_ambiguous_profile_and_type_mismatch(tmp_path):
    (tmp_path / "x.csv").write_text("h\n")
    other = dataclasses.replace(CSV_PROFILE, id="otro")
    _, out = run_one(tmp_path, "y.csv", b"h\n", cfg(CSV_PROFILE, other))
    rejected(out, IssueCode.F06_AMBIGUOUS_PROFILE, None)
    assert "sur-csv" in out.issue.message and "otro" in out.issue.message
    wrong = dataclasses.replace(XLSX_PROFILE, match="z.csv")
    _, out = run_one(tmp_path, "z.csv", b"h\n", cfg(wrong))
    rejected(out, IssueCode.F03_UNSUPPORTED_TYPE, None)


def test_fifo_is_rejected_not_opened(tmp_path):
    os.mkfifo(tmp_path / "pipe.csv")
    res = ingest(tmp_path, cfg(CSV_PROFILE))
    rejected(res.files[0], IssueCode.F01_UNREADABLE_FILE, "not_regular_file")


@pytest.mark.skipif(os.geteuid() == 0, reason="root ignora chmod")
def test_unreadable_file_is_io_error(tmp_path):
    f = tmp_path / "a.csv"
    f.write_text("h\n")
    f.chmod(0)
    res = ingest(tmp_path, cfg(CSV_PROFILE))
    f.chmod(0o600)
    rejected(res.files[0], IssueCode.F01_UNREADABLE_FILE, "io_error")


# ---- CSV ---------------------------------------------------------------------------------------------


def test_csv_traceability_multiline_and_blank(tmp_path):
    data = 'h1,h2\r\n"x\r\ny",2\r\n\r\nz,3\r\n'.encode()
    p = dataclasses.replace(CSV_PROFILE, encoding="utf-8", delimiter=",", header_row=1)
    _, out = run_one(tmp_path, "a.csv", data, cfg(p))
    f = out.ingested
    assert [(r.source_row, r.source_line, r.values) for r in f.records] == [(2, 2, ("x\r\ny", "2")), (4, 5, ("z", "3"))]
    assert (f.blank_rows, f.header_rows, f.physical_rows, f.header.values) == (1, 1, 4, ("h1", "h2"))


def test_csv_encoding_errors_and_bom(tmp_path):
    p = dataclasses.replace(CSV_PROFILE, encoding="utf-8", delimiter=",", header_row=1)
    _, out = run_one(tmp_path, "a.csv", b"h\n\xff\xfe\n", cfg(p))
    rejected(out, IssueCode.F02_ENCODING_ERROR, None)
    _, out = run_one(tmp_path, "b.csv", b"\xef\xbb\xbfh\n1\n", cfg(dataclasses.replace(p, match="b.csv")))
    assert out.ingested.header.values == ("﻿h",)  # con utf-8 el BOM se conserva (Plan §7.8)


@pytest.mark.parametrize("body", [b'h\n"abc"def\n', b'h\n"abc\n'])
def test_csv_malformed_strict(tmp_path, body):
    p = dataclasses.replace(CSV_PROFILE, encoding="utf-8", delimiter=",", header_row=1)
    _, out = run_one(tmp_path, "a.csv", body, cfg(p))
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "csv_malformed")
    assert out.ingested is None  # sin registros parciales


def test_csv_limits_and_field_size_limit_restored(tmp_path):
    p = dataclasses.replace(CSV_PROFILE, encoding="utf-8", delimiter=",", header_row=1)
    before = csv.field_size_limit()
    _, out = run_one(tmp_path, "a.csv", b"h\n" + b"x" * 11 + b"\n", cfg(p, max_text_chars=10))
    rejected(out, IssueCode.F04_LIMIT_EXCEEDED, "text_chars")
    _, out = run_one(tmp_path, "b.csv", b"h\n" + b"x" * 10 + b"\n", cfg(dataclasses.replace(p, match="b.csv"), max_text_chars=10))
    assert out.outcome == "ingested"
    assert csv.field_size_limit() == before
    _, out = run_one(tmp_path, "c.csv", b"h\n1\n2\n3\n", cfg(dataclasses.replace(p, match="c.csv"), max_records_per_file=2))
    rejected(out, IssueCode.F04_LIMIT_EXCEEDED, "records_file")


def test_file_size_limit_exact_and_over(tmp_path):
    p = dataclasses.replace(CSV_PROFILE, encoding="utf-8", delimiter=",", header_row=1)
    exact = b"h\n" + b"a\n" * ((_MB - 2) // 2)
    _, out = run_one(tmp_path, "a.csv", exact, cfg(p, max_file_mb=1, max_records_per_file=10**6, max_records_per_run=10**6))
    assert out.outcome == "ingested"
    _, out = run_one(tmp_path, "b.csv", exact + b"x", cfg(dataclasses.replace(p, match="b.csv"), max_file_mb=1, max_records_per_file=10**6, max_records_per_run=10**6))
    rejected(out, IssueCode.F04_LIMIT_EXCEEDED, "file_size")


def test_run_budget_rejects_file_but_run_continues(tmp_path):
    p = dataclasses.replace(CSV_PROFILE, encoding="utf-8", delimiter=",", header_row=1, match="*.csv")
    (tmp_path / "a.csv").write_text("h\n1\n2\n")
    (tmp_path / "b.csv").write_text("h\n1\n2\n3\n")  # no cabe: 2 + 3 > 4
    (tmp_path / "c.csv").write_text("h\n1\n")  # sí cabe: 2 + 1 <= 4
    res = ingest(tmp_path, cfg(p, max_records_per_run=4))
    st = {o.source_file: o for o in res.files}
    assert st["a.csv"].outcome == "ingested" and st["c.csv"].outcome == "ingested"
    rejected(st["b.csv"], IssueCode.F04_LIMIT_EXCEEDED, "records_run")


# ---- XLSX ---------------------------------------------------------------------------------------------


def xcfg(**limits):
    return cfg(XLSX_PROFILE, **limits)


def test_xlsx_good_roundtrip_native_values(tmp_path):
    _, out = run_one(tmp_path, "a.xlsx", good_xlsx(), xcfg())
    f = out.ingested
    assert f.header.values == ("a", "b") and [r.values for r in f.records] == [(1, 2), (3, 4)]
    assert f.records[0].source_line is None and f.records[0].data_types == ("n", "n")
    assert f.sheet == "Datos" and f.sheets_ignored == ()


def test_xlsx_sheet_not_found_and_extra_sheets_ignored(tmp_path):
    _, out = run_one(tmp_path, "a.xlsx", good_xlsx(title="Otra"), xcfg())
    rejected(out, IssueCode.F07_SHEET_NOT_FOUND, None)
    wb = openpyxl.Workbook()
    wb.active.title = "Datos"
    wb.active.append(["h"])
    wb.create_sheet("Notas")
    bio = io.BytesIO()
    wb.save(bio)
    res, out = run_one(tmp_path, "b.xlsx", bio.getvalue(), xcfg())
    assert out.ingested.sheets_ignored == ("Notas",)
    assert any(e.event == "sheet_ignored" and e.value == "Notas" for e in res.events)
    _, out = run_one(tmp_path, "c.xlsx", bio.getvalue(), xcfg(max_sheets=1))
    rejected(out, IssueCode.F04_LIMIT_EXCEEDED, "sheets")


@pytest.mark.parametrize("data", [b"", b"no soy un zip" * 50, good_xlsx()[:-30]])
def test_xlsx_not_a_zip(tmp_path, data):
    _, out = run_one(tmp_path, "a.xlsx", data, xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "zip_unreadable")


def test_xlsx_too_many_entries(tmp_path):
    data = rezip(good_xlsx(), extra=[(f"extra/{i}.bin", b"x") for i in range(10)])
    _, out = run_one(tmp_path, "a.xlsx", data, xcfg(max_zip_entries=12))
    rejected(out, IssueCode.F04_LIMIT_EXCEEDED, "zip_entries")


@pytest.mark.parametrize("name", ["../evil.bin", "/abs.bin", "a\\b.bin", "C:x.bin", "a//b.bin", "a/./b.bin"])
def test_xlsx_unsafe_names(tmp_path, name):
    _, out = run_one(tmp_path, "a.xlsx", rezip(good_xlsx(), extra=[(name, b"x")]), xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "unsafe_name")


def test_xlsx_duplicate_missing_encrypted_method(tmp_path):
    base = good_xlsx()
    _, out = run_one(tmp_path, "a.xlsx", rezip(base, extra=[("xl/WORKBOOK.xml", b"<x/>")]), xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "duplicate_name")
    _, out = run_one(tmp_path, "b.xlsx", rezip(base, mutate=lambda p: {k: v for k, v in p.items() if k != "xl/workbook.xml"}), xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "missing_part")
    enc = patch_field(base, "xl/workbook.xml", 6, 8, "<H", 1)
    _, out = run_one(tmp_path, "c.xlsx", enc, xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "encrypted_entry")
    meth = patch_field(base, "xl/workbook.xml", 8, 10, "<H", 12)  # bzip2
    _, out = run_one(tmp_path, "d.xlsx", meth, xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "unsupported_method")


def test_xlsx_overlap_and_truncated_ranges(tmp_path):
    base = good_xlsx()
    with zipfile.ZipFile(io.BytesIO(base)) as zf:
        a, b = zf.getinfo("xl/workbook.xml"), zf.getinfo("[Content_Types].xml")
    # dos registros centrales apuntando a la misma cabecera local
    over = patch_field(base, "xl/workbook.xml", 0, 42, "<I", b.header_offset)
    _, out = run_one(tmp_path, "a.xlsx", over, xcfg())
    assert out.issue.detail_code in ("overlapping_entries", "zip_truncated")
    bad_sig = bytearray(base)
    bad_sig[a.header_offset : a.header_offset + 4] = b"XXXX"
    _, out = run_one(tmp_path, "b.xlsx", bytes(bad_sig), xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "zip_truncated")


def test_xlsx_declared_size_and_ratio(tmp_path):
    data = rezip(good_xlsx(), extra=[("big.bin", b"\x00" * (3 * _MB))])  # relación real altísima
    _, out = run_one(tmp_path, "a.xlsx", data, xcfg(max_entry_uncompressed_mb=1, max_total_uncompressed_mb=2))
    rejected(out, IssueCode.F04_LIMIT_EXCEEDED, "declared_size")
    _, out = run_one(tmp_path, "b.xlsx", data, xcfg())
    rejected(out, IssueCode.F04_LIMIT_EXCEEDED, "declared_ratio")


def test_xlsx_real_size_lies_about_declared(tmp_path):
    """Bomba con metadatos mentirosos: el preflight mide los bytes REALES (H2/DT4-28)."""
    data = rezip(good_xlsx(), extra=[("big.bin", b"\x00" * (3 * _MB))])
    lie = patch_field(data, "big.bin", 22, 24, "<I", 1000)  # file_size declarado: 1000
    _, out = run_one(tmp_path, "a.xlsx", lie, xcfg(max_entry_uncompressed_mb=1, max_total_uncompressed_mb=2))
    rejected(out, IssueCode.F04_LIMIT_EXCEEDED, "entry_uncompressed")
    _, out = run_one(tmp_path, "b.xlsx", lie, xcfg(max_entry_uncompressed_mb=50, max_total_uncompressed_mb=2))
    rejected(out, IssueCode.F04_LIMIT_EXCEEDED, "total_uncompressed")
    _, out = run_one(tmp_path, "c.xlsx", lie, xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "crc_mismatch")  # tamaño real ≠ declarado (DT4-27 i)


def test_xlsx_crc_mismatch_and_corrupt_stream(tmp_path):
    base = good_xlsx()
    _, out = run_one(tmp_path, "a.xlsx", patch_field(base, "xl/workbook.xml", 14, 16, "<I", 0xDEADBEEF), xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "crc_mismatch")
    with zipfile.ZipFile(io.BytesIO(base)) as zf:
        i = zf.getinfo("xl/workbook.xml")
    corrupt = bytearray(base)
    start = i.header_offset + 30 + len(i.orig_filename) + struct.unpack_from("<H", base, i.header_offset + 28)[0]
    corrupt[start : start + 8] = b"\xff" * 8
    _, out = run_one(tmp_path, "b.xlsx", bytes(corrupt), xcfg())
    assert out.issue.detail_code in ("crc_mismatch", "zip_truncated")


def manual_zip(entries) -> bytes:
    """ZIP escrito a mano: entries = [(nombre, datos_comprimidos, método, crc, tamaño_declarado)]."""
    out, central = bytearray(), bytearray()
    for name, comp, method, crc, usize in entries:
        n = name.encode()
        off = len(out)
        out += struct.pack("<4s5H3I2H", b"PK\x03\x04", 20, 0, method, 0, 0x21, crc, len(comp), usize, len(n), 0) + n + comp
        central += struct.pack("<4s6H3I5H2I", b"PK\x01\x02", 20, 20, 0, method, 0, 0x21, crc, len(comp), usize,
                               len(n), 0, 0, 0, 0, 0, off) + n
    cd = len(out)
    out += central + struct.pack("<4s4H2LH", b"PK\x05\x06", 0, 0, len(entries), len(entries), len(central), cd, 0)
    return bytes(out)


def deflate_raw(data: bytes) -> bytes:
    c = zlib.compressobj(wbits=-15)
    return c.compress(data) + c.flush()


def xlsx_entries(extra=()):
    with zipfile.ZipFile(io.BytesIO(good_xlsx())) as zf:
        parts = [(n, zf.read(n)) for n in zf.namelist()]
    return [(n, deflate_raw(b), 8, zlib.crc32(b), len(b)) for n, b in parts] + list(extra)


def test_manual_zip_helper_produces_valid_xlsx(tmp_path):
    _, out = run_one(tmp_path, "a.xlsx", manual_zip(xlsx_entries()), xcfg())
    assert out.outcome == "ingested"


def test_xlsx_trailing_data_after_deflate_stream(tmp_path):
    raw = b"<root/>"
    extra = [("extra.bin", deflate_raw(raw) + b"TRAILER", 8, zlib.crc32(raw), len(raw))]
    _, out = run_one(tmp_path, "a.xlsx", manual_zip(xlsx_entries(extra)), xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "zip_unreadable")  # DT4-27 (ii)


def test_xlsx_incomplete_deflate_stream_and_zero_compress_size(tmp_path):
    raw = random.Random(7).randbytes(3000)  # poco compresible: no dispara declared_ratio
    comp = deflate_raw(raw)
    cut = [("extra.bin", comp[: len(comp) // 2], 8, zlib.crc32(raw), len(raw))]
    _, out = run_one(tmp_path, "a.xlsx", manual_zip(xlsx_entries(cut)), xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "zip_truncated")
    empty = [("extra.bin", b"", 8, 0, 0)]  # DEFLATE vacío sin flujo: fallo cerrado
    _, out = run_one(tmp_path, "b.xlsx", manual_zip(xlsx_entries(empty)), xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "zip_truncated")


def test_xlsx_stored_size_mismatch_and_declared_smaller(tmp_path):
    raw = b"hello"
    stored_lie = [("extra.bin", raw, 0, zlib.crc32(raw), 3)]
    _, out = run_one(tmp_path, "a.xlsx", manual_zip(xlsx_entries(stored_lie)), xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "crc_mismatch")


def test_xlsx_directory_entry_with_content(tmp_path):
    data = rezip(good_xlsx(), extra=[("carpeta/", b"payload")])
    _, out = run_one(tmp_path, "a.xlsx", data, xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "unsafe_name")
    assert out.issue.message.startswith("entrada de directorio con contenido")


def test_xlsx_empty_directory_entry_is_accepted(tmp_path):
    data = rezip(good_xlsx(), extra=[("carpeta/", b"")])
    _, out = run_one(tmp_path, "a.xlsx", data, xcfg())
    assert out.outcome == "ingested"


def test_xlsx_doctype_rejected_but_comment_with_sequence_accepted(tmp_path):
    doctype = b'<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "b">]><x/>'
    data = rezip(good_xlsx(), extra=[("docProps/custom.xml", doctype)])
    _, out = run_one(tmp_path, "a.xlsx", data, xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "xml_doctype")
    ok = rezip(good_xlsx(), extra=[("docProps/custom.xml", b"<x><!-- <!DOCTYPE x> --></x>")])
    _, out = run_one(tmp_path, "b.xlsx", ok, xcfg())
    assert out.outcome == "ingested"  # DT4-30 B: solo declaraciones reales


def test_xlsx_doctype_in_nonstandard_extension_part(tmp_path):
    """DT4-13: una hoja con extensión .dat declarada en las relaciones también se controla."""
    base = good_xlsx()
    doc = b'<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "b">]><x/>'

    def mutate(parts):
        parts["xl/_rels/workbook.xml.rels"] = parts["xl/_rels/workbook.xml.rels"].replace(b"worksheets/sheet1.xml", b"worksheets/sheet1.dat")
        parts["xl/worksheets/sheet1.dat"] = doc
        del parts["xl/worksheets/sheet1.xml"]
        return parts

    _, out = run_one(tmp_path, "a.xlsx", rezip(base, mutate), xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "xml_doctype")


@pytest.mark.parametrize(
    "payload",
    [
        '<?xml version="1.0" encoding="UTF-16"?><x/>'.encode("utf-16"),
        b'<?xml version="1.0" encoding="ISO-8859-1"?><x/>',
        "<x/>".encode("utf-16-le"),
    ],
)
def test_xlsx_non_utf8_xml_rejected(tmp_path, payload):
    data = rezip(good_xlsx(), extra=[("docProps/custom.xml", payload)])
    _, out = run_one(tmp_path, "a.xlsx", data, xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "xml_encoding")


def test_xlsx_malformed_xml_part(tmp_path):
    data = rezip(good_xlsx(), extra=[("docProps/custom.xml", b"<x><y></x>")])
    _, out = run_one(tmp_path, "a.xlsx", data, xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "xml_malformed")


def test_prologue_decision_table():
    d = ingest_mod._decide_prologue
    assert d(b"", False) == "wait" and d(b"", True) == "accept"
    assert d(b"<x/>", False) == "accept"
    assert d(b"<?xml", False) == "wait"
    assert d(b'<?xml version="1.0"?><x/>', False) == "accept"
    assert d(b"\xef\xbb\xbf<?xml version='1.0' encoding='utf-8'?>", False) == "accept"
    assert d(b'<?xml version="1.0" encoding="cp1252"?>', False) == "reject"
    assert d(b'<?xml version="1.0" ', False) == "wait" and d(b'<?xml version="1.0" ', True) == "reject"
    assert d(b"\xff\xfe<\x00", False) == "reject"
    assert d(b'<?xml-stylesheet href="a"?><x/>', False) == "accept"


def sheet_with_rows(rows_xml: str) -> bytes:
    def mutate(parts):
        s = parts["xl/worksheets/sheet1.xml"].decode()
        start, end = s.index("<sheetData>"), s.index("</sheetData>") + len("</sheetData>")
        parts["xl/worksheets/sheet1.xml"] = (s[:start] + f"<sheetData>{rows_xml}</sheetData>" + s[end:]).encode()
        return parts

    return rezip(good_xlsx(), mutate)


@pytest.mark.parametrize(
    "rows",
    [
        '<row r="2"><c r="A2"><v>1</v></c></row><row r="1"><c r="A1"><v>1</v></c></row>',  # filas fuera de orden
        '<row r="1"><c r="A1"><v>1</v></c></row><row r="1"><c r="A1"><v>2</v></c></row>',  # fila duplicada
        '<row r="1"><c r="B1"><v>1</v></c><c r="A1"><v>2</v></c></row>',  # celdas fuera de orden
        '<row r="1"><c r="A1"><v>1</v></c><c r="A1"><v>2</v></c></row>',  # celda duplicada
        '<row r="1048577"><c r="A1048577"><v>1</v></c></row>',  # fuera de rango
        '<row r="3"><c r="A2"><v>1</v></c></row>',  # celda de otra fila
    ],
)
def test_xlsx_structural_prescan_rejects_silent_loss(tmp_path, rows):
    _, out = run_one(tmp_path, "a.xlsx", sheet_with_rows(rows), xcfg())
    rejected(out, IssueCode.F01_UNREADABLE_FILE, "xml_malformed")


def test_xlsx_sparse_rows_and_wrong_dimension_are_read_fully(tmp_path):
    rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>h</t></is></c></row><row r="5"><c r="C5"><v>7</v></c></row>'
    data = sheet_with_rows(rows).replace(b"</sheetData>", b"</sheetData>")
    _, out = run_one(tmp_path, "a.xlsx", data, cfg(dataclasses.replace(XLSX_PROFILE, header_row=1)))
    f = out.ingested
    assert f.physical_rows == 5 and f.blank_rows == 3 and [r.source_row for r in f.records] == [5]
    assert f.records[0].values[2] == 7


def test_xlsx_row_cell_limits_no_partials(tmp_path):
    many = good_xlsx(rows=[[1] * 5] + [[1] * 5] * 10)
    _, out = run_one(tmp_path, "a.xlsx", many, xcfg(max_columns=4))
    rejected(out, IssueCode.F04_LIMIT_EXCEEDED, "columns")
    _, out = run_one(tmp_path, "b.xlsx", many, xcfg(max_cells_per_file=20))
    rejected(out, IssueCode.F04_LIMIT_EXCEEDED, "cells")
    _, out = run_one(tmp_path, "c.xlsx", good_xlsx(rows=[["x" * 11]]), xcfg(max_text_chars=10))
    rejected(out, IssueCode.F04_LIMIT_EXCEEDED, "text_chars")
    _, out = run_one(tmp_path, "d.xlsx", many, xcfg(max_records_per_file=3))
    rejected(out, IssueCode.F04_LIMIT_EXCEEDED, "records_file")
    assert out.ingested is None


def xlsx_with_shared_strings(n: int, text: str = "texto") -> bytes:
    sst = ('<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
           + "".join(f"<si><t>{text}-{i}</t></si>" for i in range(n)) + "</sst>").encode()

    def mutate(parts):
        ct = parts["[Content_Types].xml"].decode().replace(
            "</Types>",
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.'
            'spreadsheetml.sharedStrings+xml"/></Types>')
        parts["[Content_Types].xml"] = ct.encode()
        parts["xl/sharedStrings.xml"] = sst
        return parts

    return rezip(good_xlsx(), mutate)


def test_xlsx_shared_strings_limits(tmp_path):
    data = xlsx_with_shared_strings(30)
    _, out = run_one(tmp_path, "a.xlsx", data, xcfg())
    assert out.outcome == "ingested"
    _, out = run_one(tmp_path, "b.xlsx", data, xcfg(max_shared_strings=10))
    rejected(out, IssueCode.F04_LIMIT_EXCEEDED, "shared_strings_count")
    _, out = run_one(tmp_path, "c.xlsx", data, xcfg(max_shared_strings_chars=50))
    rejected(out, IssueCode.F04_LIMIT_EXCEEDED, "shared_strings_chars")


def test_xlsx_buffer_is_read_once(tmp_path, monkeypatch):
    f = tmp_path / "a.xlsx"
    f.write_bytes(good_xlsx())
    opens = []
    real_open = os.open
    monkeypatch.setattr(os, "open", lambda p, *a, **k: (opens.append(str(p)), real_open(p, *a, **k))[1])
    res = ingest(tmp_path, xcfg())
    assert res.files[0].outcome == "ingested" and opens.count(str(f)) == 1


def test_rejection_does_not_stop_the_run_and_events_have_no_paths(tmp_path):
    (tmp_path / "bad.xlsx").write_bytes(b"x" * 100)
    (tmp_path / "ok.csv").write_text("h\n1\n")
    p = dataclasses.replace(CSV_PROFILE, encoding="utf-8", delimiter=",", header_row=1, match="ok.csv")
    q = dataclasses.replace(XLSX_PROFILE, match="bad.xlsx")
    res = ingest(tmp_path, cfg(p, q))
    assert {o.source_file: o.outcome for o in res.files} == {"bad.xlsx": "rejected", "ok.csv": "ingested"}
    assert all(str(tmp_path) not in (e.value or "") for e in res.events)
    assert len(res.issues) == 1 and res.issues[0].source_id == q.id
