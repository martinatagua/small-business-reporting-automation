"""T3-04 … T3-12, T3-18, T3-19, T3-20: contenido, determinismo y contratos del dataset sintético."""

import collections
import csv
import io
import json
import os
import re
import subprocess
import sys
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest

from small_business_reporting_automation import synthetic

import sample_oracle as oracle
from sample_oracle import SEED, SOURCES, TABLE

REPO = Path(__file__).resolve().parents[2]
NAME_TO_KEY = {Path(v[0]).name: k for k, v in SOURCES.items()}


def records(root):
    return oracle.read_gt_csv(root, "expected_records.csv")


def by_key(rows):
    out = collections.defaultdict(list)
    for r in rows:
        out[NAME_TO_KEY[r["source_file"]]].append(r)
    return out


# --------------------------------------------------------------------------------------------------
# T3-04
# --------------------------------------------------------------------------------------------------


def test_t3_04_plan_matches_spec_table(sample_root):
    spec = {s[0]: s for s in synthetic.SCENARIOS if s[3] is not None}
    assert set(spec) == set(TABLE)
    for sid, (codes, status, kpi, counts) in TABLE.items():
        _id, _name, gen_codes, gen_status, gen_kpi, gen_counts = spec[sid]
        assert tuple(c[:3] for c in gen_codes) == codes
        assert (gen_status, gen_kpi, gen_counts) == (status, kpi, counts)
    rows = records(sample_root)
    got = collections.Counter((r["scenario"], NAME_TO_KEY[r["source_file"]]) for r in rows)
    want = {(sid, f): n for sid, (_c, _s, _k, counts) in TABLE.items() for f, n in counts.items()}
    assert dict(got) == want
    assert len(rows) == 177
    per_file = {"N": (55, 46, 7, 2, 42, 4), "C": (53, 46, 6, 1, 42, 4), "SX": (34, 28, 5, 1, 26, 2),
                "SC": (35, 28, 5, 2, 27, 1)}
    for key, rs in by_key(rows).items():
        st = collections.Counter(r["expected_status"] for r in rs)
        kp = collections.Counter(r["kpi_effect"] for r in rs)
        assert (len(rs), st["valid"], st["observed"], st["out_of_scope"], kp["included"],
                kp["included_warning"]) == per_file[key]
    assert collections.Counter(r["expected_status"] for r in rows) == {"valid": 148, "observed": 23, "out_of_scope": 6}
    for r in rows:
        codes, status, kpi, _ = TABLE[r["scenario"]]
        assert r["expected_status"] == status and r["kpi_effect"] == kpi
        assert r["expected_issue_codes"].split(";") == ([c for c in
               [x for x in synthetic.ISSUE_CODES if x[:3] in codes]] if codes else [""])
    # S24: al menos una fila por categoría aplicable a cada archivo
    plan = synthetic.build_plan(SEED)
    applicable = {"N": {"date_text", "sku_lower", "sku_padded"}, "C": {"sku_padded", "thousands"},
                  "SX": {"sku_lower"}, "SC": {"sku_lower", "sku_padded"}}
    for key, cats in applicable.items():
        seen = {n for r in plan.layout[key] if r is not None and r.scenario_id == "S01" for n in r.noise}
        assert cats <= seen
    assert all(not r.noise for k in plan.layout for r in plan.layout[k] if r is not None and r.scenario_id != "S01")


# --------------------------------------------------------------------------------------------------
# T3-05 / T3-07
# --------------------------------------------------------------------------------------------------


def same_dataset(a, b):
    for rel in oracle.EXPECTED_MANIFEST_PATHS + ["ground_truth/manifest.json"]:
        if rel.endswith(".xlsx"):
            assert oracle.xlsx_digest(a / rel) == oracle.xlsx_digest(b / rel), rel
        else:
            assert (a / rel).read_bytes() == (b / rel).read_bytes(), rel


def test_t3_05_generation_determinism(sample_root, tmp_path):
    synthetic.generate(SEED, tmp_path / "again")
    same_dataset(sample_root, tmp_path / "again")
    synthetic.generate(7, tmp_path / "other")
    assert (tmp_path / "other/raw/norte_ventas_2026-10.xlsx").read_bytes() != (
        sample_root / "raw/norte_ventas_2026-10.xlsx").read_bytes()
    assert oracle.xlsx_digest(tmp_path / "other/raw/norte_ventas_2026-10.xlsx") != oracle.xlsx_digest(
        sample_root / "raw/norte_ventas_2026-10.xlsx")
    assert (tmp_path / "other/raw/sur_ventas_2026-10_web.csv").read_bytes() != (
        sample_root / "raw/sur_ventas_2026-10_web.csv").read_bytes()
    count = lambda root: collections.Counter((r["scenario"], r["source_file"]) for r in records(root))
    assert count(tmp_path / "other") == count(sample_root)


def test_t3_07_hashseed_independence(sample_root, tmp_path):
    code = "import sys; from small_business_reporting_automation import synthetic as s; s.generate(%d, sys.argv[1])" % SEED
    for i, hashseed in enumerate(("0", "1", "4242")):
        env = dict(os.environ, PYTHONHASHSEED=hashseed)
        subprocess.run([sys.executable, "-B", "-c", code, str(tmp_path / f"h{i}")], check=True, env=env, cwd=tmp_path)
        same_dataset(sample_root, tmp_path / f"h{i}")


# --------------------------------------------------------------------------------------------------
# T3-06
# --------------------------------------------------------------------------------------------------


def test_t3_06_manifest_matches_files(sample_root):
    raw = (sample_root / "ground_truth/manifest.json").read_bytes()
    manifest = json.loads(raw.decode("utf-8"))
    assert oracle.manifest_problems(manifest) == []
    assert raw == (json.dumps(manifest, sort_keys=True, indent=2, ensure_ascii=False) + "\n").encode()
    assert manifest["seed"] == SEED and manifest["generator_version"] == "1" and manifest["config_version"] == 1
    assert [e["path"] for e in manifest["files"]] == oracle.EXPECTED_MANIFEST_PATHS
    for entry in manifest["files"]:
        assert oracle.file_digest(sample_root, entry["path"]) == entry["content_sha256"], entry["path"]
    # sin hash propio ni campos adicionales
    assert not any("manifest" in e["path"] for e in manifest["files"])
    assert b"bytes_sha256" not in raw and b"manifest_version" not in raw and b"plan_sha256" not in raw
    # control físico (≠ cells-v1) del catálogo y de los tres CSV de ground truth: \n
    for rel in oracle.EXPECTED_MANIFEST_PATHS:
        if rel in oracle.EOL_LF:
            assert oracle.physical_problems((sample_root / rel).read_bytes(), b"\n") == [], rel
            assert b"\r" not in (sample_root / rel).read_bytes()
    for rel in oracle.EOL_CRLF:
        assert oracle.physical_problems((sample_root / rel).read_bytes(), b"\r\n") == [], rel


def test_t3_06_manifest_negative_cases(sample_root):
    good = json.loads((sample_root / "ground_truth/manifest.json").read_text(encoding="utf-8"))

    def mutate(fn):
        copy = json.loads(json.dumps(good))
        fn(copy)
        return copy

    cases = {
        "ruta adicional": lambda m: m["files"].append({"path": "raw/x.csv", "content_sha256": "0" * 64}),
        "ruta ausente": lambda m: m["files"].pop(),
        "duplicada": lambda m: m["files"].__setitem__(1, dict(m["files"][0])),
        "absoluta": lambda m: m["files"][0].__setitem__("path", "/etc/passwd"),
        "traversal": lambda m: m["files"][0].__setitem__("path", "raw/../x.csv"),
        "sustituida": lambda m: m["files"][0].__setitem__("path", "raw/otra.csv"),
        "orden": lambda m: m["files"].reverse(),
        "hash en mayúsculas": lambda m: m["files"][0].__setitem__("content_sha256", "A" * 64),
        "hash corto": lambda m: m["files"][0].__setitem__("content_sha256", "ab"),
        "campo extra": lambda m: m["files"][0].__setitem__("bytes_sha256", "0" * 64),
        "clave extra": lambda m: m.__setitem__("manifest_version", 1),
        "manifiesto en files": lambda m: m["files"][0].__setitem__("path", "ground_truth/manifest.json"),
    }
    for label, fn in cases.items():
        assert oracle.manifest_problems(mutate(fn)), label
    for label, bad in {"8 elementos": 8, "10 elementos": 10}.items():
        m = mutate(lambda x: None)
        m["files"] = (m["files"] * 2)[:bad]
        assert oracle.manifest_problems(m), label
    # el control físico detecta lo que cells-v1 no ve
    cat = (sample_root / "reference/products.csv").read_bytes()
    assert oracle.physical_problems(cat.replace(b"\n", b"\r\n"), b"\n")
    assert oracle.physical_problems(cat + b"\r", b"\n")


# --------------------------------------------------------------------------------------------------
# T3-08
# --------------------------------------------------------------------------------------------------


def test_t3_08_source_file_contracts(sample_root):
    # --- N (XLSX)
    names, ws, rows = oracle.xlsx_rows(sample_root / SOURCES["N"][0])
    assert names == ["Ventas"]
    assert ws.cell(1, 1).value and ws.cell(2, 1).value and ws.cell(1, 2).value is None
    assert [c.value for c in rows[2]] == ["Fecha", "Nº Ticket", "Código", "Producto", "Cantidad", "Precio Unitario",
                                          "Total"]
    last = [c.value for c in rows[-1]]
    assert last[3] == "Total general" and last[:3] == [None] * 3 and last[4:6] == [None, None] and last[6] is not None
    blanks = [i for i, r in enumerate(rows[3:-1], 4) if all(c.value is None for c in r)]
    assert len(blanks) == 1 and 4 < blanks[0] < len(rows) - 1
    dates = [rows[i][0].value for i in range(3, len(rows) - 1) if rows[i][0].value is not None]
    assert any(hasattr(d, "year") for d in dates) and any(isinstance(d, str) and re.fullmatch(r"\d\d/\d\d/\d{4}", d)
                                                          for d in dates)
    assert not {c.data_type for r in rows for c in r} & {"f", "e"}
    texts = {c.value: c.data_type for r in rows for c in r if isinstance(c.value, str)}
    assert texts["=1+1"] == "s" and texts["@SUM(1,1)"] == "s"
    assert any(c.value.startswith(" ") or c.value.endswith(" ") for r in rows[3:-1] for c in [r[2]]
               if isinstance(c.value, str))
    # --- SX (XLSX)
    names, ws, rows = oracle.xlsx_rows(sample_root / SOURCES["SX"][0])
    assert names == ["Datos"]
    assert [c.value for c in rows[0]] == ["Date", "Receipt", "SKU", "Item", "Qty", "Unit Price", "Sucursal", "Cajero"]
    assert {"SUR", "sur", "Sucursal Sur", "CENTRO"} <= {r[6].value for r in rows[1:]}
    assert {c.value: c.data_type for r in rows for c in r}["#N/A"] == "s"
    assert not {c.data_type for r in rows for c in r} & {"f", "e"}
    assert sum(1 for r in rows[1:] if all(c.value is None for c in r)) == 1
    assert all(re.fullmatch(r"\d{4}-\d\d-\d\d", r[0].value) for r in rows[1:] if r[0].value is not None)
    # --- C (CSV cp1252, ;)
    raw_c = (sample_root / SOURCES["C"][0]).read_bytes()
    assert not raw_c.startswith(b"\xef\xbb\xbf") and any(b > 0x7F for b in raw_c)
    recs = oracle.csv_records(raw_c, "cp1252", ";")
    assert recs[0] == ["ticket", "fecha_venta", "sku", "descripcion", "unidades", "precio", "importe", "Observaciones"]
    assert all(len(r) == 8 for r in recs)
    assert all(re.fullmatch(r"\d\d-\d\d-\d{4}|N/A", r[1]) for r in recs[1:])
    assert {"+1", " -1+2"} <= {r[3] for r in recs}
    assert any(re.fullmatch(r"\d{1,3}(\.\d{3})+,\d\d", r[5]) or re.fullmatch(r"\d{1,3}(\.\d{3})+,\d\d", r[6])
               for r in recs[1:])
    # --- SC (CSV UTF-8 con BOM, ,)
    raw_s = (sample_root / SOURCES["SC"][0]).read_bytes()
    assert raw_s.startswith(b"\xef\xbb\xbf")
    recs_s = oracle.csv_records(raw_s, "utf-8-sig", ",")
    assert recs_s[0] == ["fecha", "pedido", "sku", "nombre", "cant", "pu"]
    assert any(r and r[3].startswith("\t=1+1") for r in recs_s) and any(r and r[3].startswith("\x07") for r in recs_s)
    assert sum(1 for r in recs_s if r == []) == 1
    gt = {int(r["source_row"]): r["scenario"] for r in by_key(records(sample_root))["SC"]}
    for number, rec in enumerate(recs_s, 1):
        if number == 1 or not rec or gt[number] in {"S10", "S11", "S12"}:
            continue
        assert int(rec[0].split("/")[0]) >= 13, (number, rec[0])
    # controles solo en SC
    ctrl = re.compile(r"[\x00-\x08\x0b-\x1f]")
    for rel in ("N", "SX"):
        _n, _w, rws = oracle.xlsx_rows(sample_root / SOURCES[rel][0])
        assert not any(ctrl.search(c.value) for r in rws for c in r if isinstance(c.value, str))
    assert not any(ctrl.search(v) for r in recs for v in r) and any(ctrl.search(v) for r in recs_s for v in r)
    # --- terminadores físicos de las fuentes CSV (bytes reales) y su diferencia con cells-v1
    for data, enc, delim in ((raw_c, "cp1252", ";"), (raw_s, "utf-8-sig", ",")):
        assert oracle.physical_problems(data, b"\r\n") == []
        variants = {
            "mixto": data.replace(b"\r\n", b"\n", 1),
            "LF aislado": data.replace(b"\r\n", b"\n"),
            "CR aislado": data.replace(b"\r\n", b"\r"),
            "sin terminador final": data[:-2],
        }
        for label, mutated in variants.items():
            assert oracle.physical_problems(mutated, b"\r\n"), label
        # el cambio de terminador NO altera los registros ni cells-v1, pero el control físico lo rechaza
        lf_only = data.replace(b"\r\n", b"\n")
        assert oracle.csv_records(lf_only, enc, delim) == oracle.csv_records(data, enc, delim)
        assert oracle.csv_digest(lf_only, enc, delim) == oracle.csv_digest(data, enc, delim)
        assert oracle.physical_problems(lf_only, b"\r\n")


# --------------------------------------------------------------------------------------------------
# T3-09
# --------------------------------------------------------------------------------------------------


def test_t3_09_ground_truth_consistency(sample_root):
    recs = records(sample_root)
    issues = oracle.read_gt_csv(sample_root, "expected_issues.csv")
    files = oracle.read_gt_csv(sample_root, "expected_files.csv")
    summary = json.loads((sample_root / "ground_truth/expected_summary.json").read_text(encoding="utf-8"))
    assert list(issues[0]) == ["level", "source_file", "branch_id", "source_row", "issue_code"]
    tuples = {tuple(i.values()) for i in issues}
    assert len(tuples) == len(issues) == 43
    # records ↔ issues
    from_records = set()
    for r in recs:
        for code in [c for c in r["expected_issue_codes"].split(";") if c]:
            from_records.add(("record", r["source_file"], r["branch_id"], r["source_row"], code))
    assert tuples - from_records == {("file", "centro_ventas_2026-10.csv", "CENTRO", "", "F11_UNMAPPED_COLUMN"),
                                     ("file", "sur_ventas_2026-10_tienda.xlsx", "SUR", "", "F11_UNMAPPED_COLUMN")}
    assert from_records <= tuples and len(from_records) == 41
    # estado ↔ kpi_effect y códigos
    for r in recs:
        codes = [c for c in r["expected_issue_codes"].split(";") if c]
        if r["expected_status"] == "valid":
            assert r["kpi_effect"] in {"included", "included_warning"}
            assert (r["kpi_effect"] == "included_warning") == bool(codes)
        else:
            assert codes and r["kpi_effect"] == "excluded_" + r["expected_status"]
    assert {r["scenario"] for r in recs} == {f"S{n:02d}" for n in range(1, 22)}
    # records ↔ files: invariante de §7.5 (filas físicas = preámbulo + 1 + blancos + registros)
    for f in files:
        key = NAME_TO_KEY[f["source_file"]]
        n = sum(1 for r in recs if r["source_file"] == f["source_file"])
        assert int(f["expected_records"]) == n and f["expected_outcome"] == "processed" and f["rejection_code"] == ""
        physical = len(oracle.source_records(sample_root, key))
        assert physical == int(f["preamble_rows"]) + int(f["header_rows"]) + int(f["blank_rows"]) + n
    assert [f["source_file"] for f in files] == sorted((f["source_file"] for f in files),
                                                       key=lambda n: SOURCES[NAME_TO_KEY[n]][1])
    # records ↔ summary
    assert summary["records"]["total"] == 177
    assert summary["records"]["by_status"] == {"observed": 23, "out_of_scope": 6, "valid": 148}
    assert summary["records"]["by_branch"] == {"CENTRO": 53, "NORTE": 55, "SUR": 69}
    assert summary["records"]["by_file"] == {f["source_file"]: int(f["expected_records"]) for f in files}
    by_code = collections.Counter(t[4] for t in tuples)
    assert {k: v for k, v in summary["issues_by_code"].items() if v} == dict(by_code)
    assert len(summary["issues_by_code"]) == 26
    assert summary["issues_by_code"]["R01_MISSING_REQUIRED"] == 6 and summary["issues_by_code"]["R13_DUPLICATE_SUSPECT"] == 6
    assert summary["out_of_scope_by_code"] == {"R07_NEGATIVE_QUANTITY": 3, "R10_DATE_OUT_OF_PERIOD": 3}
    # sumas Decimal e importe potencialmente duplicado = Σ(n−1)×line_amount por scenario S16–S18
    valid = [r for r in recs if r["expected_status"] == "valid"]
    assert all(r["expected_line_amount"] for r in valid)
    assert str(sum((Decimal(r["expected_line_amount"]) for r in valid), Decimal("0"))) == summary["sales"][
        "gross_ordinary_total"]
    per_branch = collections.defaultdict(lambda: Decimal("0"))
    for r in valid:
        per_branch[r["branch_id"]] += Decimal(r["expected_line_amount"])
    assert {k: str(v) for k, v in per_branch.items()} == summary["sales"]["gross_ordinary_by_branch"]
    dup = Decimal("0")
    members = collections.defaultdict(list)
    for r in recs:
        if r["scenario"] in {"S16", "S17", "S18"}:
            members[r["scenario"]].append(Decimal(r["expected_line_amount"]))
    assert {k: len(v) for k, v in members.items()} == {"S16": 2, "S17": 2, "S18": 2}
    for amounts in members.values():
        assert len(set(amounts)) == 1
        dup += amounts[0] * (len(amounts) - 1)
    assert str(dup) == summary["duplicates"]["potentially_duplicated_amount"]
    assert summary["duplicates"] == {"groups": 3, "records": 6, "excess": 3,
                                     "potentially_duplicated_amount": str(dup)}
    assert summary["sales"]["zero_price_lines"] == 2 and isinstance(summary["sales"]["units"], int)
    plan = synthetic.build_plan(SEED)
    assert summary["sales"]["units"] == sum(r.qty for k in plan.layout for r in plan.layout[k]
                                            if r is not None and r.expected_status == "valid")
    # 8 importes vacíos exactamente en los escenarios de Plan §6.1
    empty = collections.Counter((r["scenario"], NAME_TO_KEY[r["source_file"]]) for r in recs
                                if r["expected_line_amount"] == "")
    assert empty == collections.Counter({("S07", "N"): 1, ("S08", "N"): 1, ("S08", "C"): 1, ("S08", "SC"): 1,
                                         ("S19", "N"): 1, ("S06", "SC"): 1, ("S09", "C"): 1, ("S09", "SX"): 1})


# --------------------------------------------------------------------------------------------------
# T3-10
# --------------------------------------------------------------------------------------------------


def test_t3_10_traceability_plan(sample_root):
    rows = records(sample_root)
    for key, rs in by_key(rows).items():
        _rel, _sid, _sheet, header, enc, _d = SOURCES[key]
        physical = oracle.source_records(sample_root, key)
        numbers = [int(r["source_row"]) for r in rs]
        non_empty = {i for i, rec in enumerate(physical, 1) if i > header and any(v is not None for v in rec)}
        assert set(numbers) == non_empty and len(numbers) == len(set(numbers))
        for r in rs:
            if enc is None:
                assert r["source_line"] == "" and r["source_sheet"] == SOURCES[key][2]
            else:
                assert r["source_line"] == r["source_row"] and r["source_sheet"] == ""
        if enc is not None:  # línea física = registro lógico
            raw = (sample_root / SOURCES[key][0]).read_bytes()
            lines = raw.split(b"\r\n")[:-1]
            assert len(lines) == len(physical)
            for i, rec in enumerate(physical):
                assert (lines[i].replace(b"\xef\xbb\xbf", b"") == b"") == all(v is None for v in rec)
    pos = {(NAME_TO_KEY[r["source_file"]], r["scenario"]): [] for r in rows}
    for r in rows:
        pos[(NAME_TO_KEY[r["source_file"]], r["scenario"])].append(int(r["source_row"]))
    n_rows = oracle.source_records(sample_root, "N")
    assert pos[("N", "S07")] == [len(n_rows)] and n_rows[-1][3] == "Total general"
    sx = oracle.source_records(sample_root, "SX")
    assert sx[pos[("SX", "S15")][0] - 1][6] == "CENTRO"
    sc = oracle.source_records(sample_root, "SC")
    assert sorted(sc[i - 1][0] for i in pos[("SC", "S11")]) == ["03/10/2026", "09/10/2026"]
    assert sc[pos[("SC", "S12")][0] - 1][0] == "30/09/2026" and sx[pos[("SX", "S12")][0] - 1][0] == "2026-09-30"
    assert n_rows[pos[("N", "S12")][0] - 1][0] == "01/11/2026"
    assert n_rows[pos[("N", "S19")][0] - 1][0] is None and n_rows[pos[("N", "S19")][0] - 1][4] == "dos"


# --------------------------------------------------------------------------------------------------
# T3-11, T3-12, T3-20
# --------------------------------------------------------------------------------------------------


def catalog(root):
    data = (root / "reference/products.csv").read_bytes()
    assert not data.startswith(b"\xef\xbb\xbf")
    return list(csv.reader(io.StringIO(data.decode("utf-8"), newline="")))


def written_skus(root):
    out = []
    for r in records(root):
        key = NAME_TO_KEY[r["source_file"]]
        rec = oracle.source_records(root, key)[int(r["source_row"]) - 1]
        out.append((r["scenario"], rec[2]))
    return out


def test_t3_11_catalog(sample_root):
    rows = catalog(sample_root)
    assert rows[0] == ["product_sku", "product_name"] and len(rows) == 21 and all(len(r) == 2 for r in rows)
    skus = [r[0] for r in rows[1:]]
    assert len(set(skus)) == 20 and all(re.fullmatch(r"[A-Z0-9-]+", s) for s in skus)
    for scenario, sku in written_skus(sample_root):
        if scenario == "S07":
            continue
        canon = sku.strip(" ").upper()
        assert (canon in skus) == (scenario != "S14"), (scenario, sku)
    assert {sku.strip(" ") for sc, sku in written_skus(sample_root) if sc == "S14"} == set(synthetic.UNKNOWN_SKUS)


def test_t3_12_no_personal_data(sample_root):
    email = re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+")
    phone = re.compile(r"(?:\d[\s().-]?){9,}")  # 9 o más dígitos seguidos (las fechas tienen 8)
    names = {r[1] for r in catalog(sample_root)[1:]}
    danger = {"Total general", "=1+1", "@SUM(1,1)", "+1", " -1+2", "#N/A", "\t=1+1", "\x07Texto con control"}
    obs = {"Sin observaciones", "Entrega en mostrador", "Pago en efectivo"}
    for key in SOURCES:
        header = oracle.source_records(sample_root, key)[SOURCES[key][3] - 1]
        for rec in oracle.source_records(sample_root, key)[SOURCES[key][3]:]:
            for h, v in zip(header, rec):
                if isinstance(v, str):
                    assert not email.search(v) and not phone.search(v), v
                if h == "Cajero" and v is not None:
                    assert re.fullmatch(r"Caja \d", v)
                if h == "Observaciones" and v is not None:
                    assert v in obs
                if h in {"Producto", "descripcion", "Item", "nombre"} and v is not None:
                    assert v in names | danger, v


def test_t3_20_sku_variants_invariant(sample_root):
    for scenario, sku in written_skus(sample_root):
        if scenario == "S07":
            continue
        canon = sku.strip(" ").upper()
        assert re.fullmatch(r"[A-Z0-9-]+", canon)
        assert sku in {canon, canon.lower(), " " + canon, canon + " "}
        assert sku.strip(" ") == sku.strip() and not any(ch.isspace() for ch in sku.strip(" "))
        assert not any(ord(ch) > 127 for ch in sku)
    plan = synthetic.build_plan(SEED)
    for rows in plan.layout.values():
        for r in rows:
            if r is not None and r.sku:
                assert r.sku_written in {r.sku, r.sku.lower(), " " + r.sku, r.sku + " "}
    assert synthetic.__dict__.get("canonical_text") is None


# --------------------------------------------------------------------------------------------------
# T3-18
# --------------------------------------------------------------------------------------------------


def xlsx_decimal(value):
    return Decimal(repr(value))  # §7.3 para valores nativos


def csv_decimal(text, decimal, thousands):
    if thousands:
        text = text.replace(thousands, "")
    return Decimal(text.replace(decimal, "."))


def line_from_source(key, rec):
    """Cantidad × precio con las conversiones de §7.3 hechas en el test; ``None`` si no hay valores canónicos."""
    qty_raw, price_raw = rec[4], rec[5]
    try:
        if SOURCES[key][4] is None:
            if not isinstance(qty_raw, (int, float)) or not isinstance(price_raw, (int, float)):
                return None
            if isinstance(qty_raw, float) and qty_raw != int(qty_raw):
                return None
            return int(qty_raw) * xlsx_decimal(price_raw)
        dec, thou = (",", ".") if key == "C" else (".", "")
        if qty_raw is None or price_raw is None or not re.fullmatch(r"-?\d+", qty_raw):
            return None
        return int(qty_raw) * csv_decimal(price_raw, dec, thou)
    except Exception:
        return None


def test_t3_18_decimal_representation_contract(sample_root, tmp_path):
    # hipótesis de openpyxl 3.1.5 (T-7): 12.0 se relee como int 12
    wb = openpyxl.Workbook()
    wb.active["A1"], wb.active["A2"], wb.active["A3"] = 12.0, 12.5, 3
    wb.save(tmp_path / "t.xlsx")
    vals = [c.value for c in openpyxl.load_workbook(tmp_path / "t.xlsx").active["A"]]
    assert vals == [12, 12.5, 3] and type(vals[0]) is int and Decimal(repr(vals[0])) == Decimal("12")
    assert Decimal(repr(float(12))) == Decimal("12.0")
    assert str(Decimal(repr(float(12)))) != str(Decimal(repr(vals[0])))  # por qué no sirve repr(float(v))
    checked = 0
    for r in records(sample_root):
        key = NAME_TO_KEY[r["source_file"]]
        rec = oracle.source_records(sample_root, key)[int(r["source_row"]) - 1]
        got = line_from_source(key, rec)
        expected = r["expected_line_amount"]
        if expected == "":
            assert got is None, (r["scenario"], key)
            continue
        assert got is not None, (r["scenario"], key, rec)
        assert Decimal(expected) == got, "P1: igualdad de valor"
        assert str(got) == expected, "P2: igualdad de representación"
        assert got.as_tuple() == Decimal(expected).as_tuple()
        checked += 1
    assert checked == 169
    # S17 cruza XLSX y CSV con la misma representación
    s17 = [r["expected_line_amount"] for r in records(sample_root) if r["scenario"] == "S17"]
    assert len(s17) == 2 and s17[0] == s17[1]


# --------------------------------------------------------------------------------------------------
# T3-19
# --------------------------------------------------------------------------------------------------


def test_t3_19_repository_sample(tmp_path):
    data = REPO / "data"
    raw = (data / "ground_truth/manifest.json").read_bytes()
    manifest = json.loads(raw.decode("utf-8"))
    assert oracle.manifest_problems(manifest) == []
    assert [e["path"] for e in manifest["files"]] == oracle.EXPECTED_MANIFEST_PATHS
    for entry in manifest["files"]:
        assert oracle.file_digest(data, entry["path"]) == entry["content_sha256"], entry["path"]
    for rel in oracle.EOL_LF:
        assert oracle.physical_problems((data / rel).read_bytes(), b"\n") == [], rel
    for rel in oracle.EOL_CRLF:
        assert oracle.physical_problems((data / rel).read_bytes(), b"\r\n") == [], rel
    synthetic.generate(manifest["seed"], tmp_path / "regen")
    for rel in oracle.EXPECTED_MANIFEST_PATHS + ["ground_truth/manifest.json"]:
        if rel.endswith(".xlsx"):
            assert oracle.xlsx_digest(data / rel) == oracle.xlsx_digest(tmp_path / "regen" / rel), rel
        else:  # CSV y JSON: byte a byte
            assert (data / rel).read_bytes() == (tmp_path / "regen" / rel).read_bytes(), rel
