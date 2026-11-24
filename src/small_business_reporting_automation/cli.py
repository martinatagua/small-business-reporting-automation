"""CLI ``sb-report``. F3 solo ofrece ``generate-sample``; ``run`` y ``recover`` llegan en fases posteriores."""

from __future__ import annotations

import argparse
import sys
from collections.abc import Sequence
from pathlib import Path

from .config import ConfigError, load_config
from .ingest import EnvironmentCheckError, ingest
from .synthetic import SEED_MAX, GenerateError, generate

PENDING_MESSAGE = (
    "Disponible: generate-sample (F3) e ingest (F4). "
    "Normalización, reglas, consolidación e informes están pendientes (F5 en adelante)."
)
DEMO_CONFIG = Path("config") / "demo.toml"  # relativa al directorio actual (Plan §8.1, E-1)


def _seed(text: str) -> int:
    try:
        value = int(text)
    except ValueError:
        raise argparse.ArgumentTypeError(f"semilla no entera: {text!r}") from None
    if not 0 <= value <= SEED_MAX:
        raise argparse.ArgumentTypeError(f"semilla fuera de rango [0, 2^63 - 1]: {value}")
    return value


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="sb-report",
        description="Small Business Reporting Automation. " + PENDING_MESSAGE,
    )
    sub = parser.add_subparsers(dest="command", metavar="COMANDO")
    sample = sub.add_parser(
        "generate-sample",
        help="genera los datos sintéticos de demostración y su ground truth",
        description="Genera el dataset de demostración (§10.1) con una semilla. Sin --seed lee [sample] de "
        "config/demo.toml (relativo al directorio actual).",
    )
    sample.add_argument("--seed", type=_seed, default=None, help="semilla entera en [0, 2^63-1]; prevalece sobre la configuración")
    sample.add_argument("--output-root", default="data", help="raíz de salida (por defecto: data)")
    sample.add_argument("--overwrite", action="store_true", help="permite reemplazar un conjunto existente")
    ing = sub.add_parser(
        "ingest",
        help="ingiere los archivos de un directorio (F4) y muestra un resumen",
        description="Descubre, valida y lee los CSV/XLSX de un directorio según los perfiles de la configuración. "
        "No normaliza ni escribe archivos de salida (eso es de F5 en adelante).",
    )
    ing.add_argument("--config", default=str(DEMO_CONFIG), help="configuración TOML (por defecto: config/demo.toml)")
    ing.add_argument("--input-dir", default="data/raw", help="directorio de entrada (por defecto: data/raw)")
    return parser


def _run_ingest(config_path: str, input_dir: str) -> int:
    try:
        config = load_config(config_path)
        result = ingest(input_dir, config)
    except ConfigError as exc:
        print(str(exc), file=sys.stderr)
        return 2
    except EnvironmentCheckError as exc:
        print(f"error de entorno: {exc}", file=sys.stderr)
        return 2
    except OSError as exc:
        print(f"error: no se puede leer {input_dir}: {exc.strerror}", file=sys.stderr)
        return 2
    print(f"{'archivo':34} {'estado':9} {'perfil':12} {'filas':>5} {'prea.':>5} {'cab.':>4} {'blanc.':>6} {'regs.':>5}  detalle")
    total = 0
    for o in result.files:
        f = o.ingested
        if f is not None:
            total += len(f.records)
            print(f"{o.source_file:34} {o.outcome:9} {o.source_id:12} {f.physical_rows:5} {f.preamble_rows:5} "
                  f"{f.header_rows:4} {f.blank_rows:6} {len(f.records):5}  sha256={f.sha256[:12]}…")
        elif o.issue is not None:
            print(f"{o.source_file:34} {o.outcome:9} {o.source_id or '-':12} {'':>5} {'':>5} {'':>4} {'':>6} {'':>5}  "
                  f"{o.issue.issue_code.value[:3]} {o.issue.detail_code or ''}: {o.issue.message}")
        else:
            print(f"{o.source_file:34} {o.outcome:9} {'-':12} {'':>5} {'':>5} {'':>4} {'':>6} {'':>5}  motivo={o.ignore_reason}")
    accepted = sum(1 for o in result.files if o.outcome == "ingested")
    print(f"\nIngesta completada: {accepted} archivos aceptados, {len(result.issues)} rechazados, {total} registros de origen.")
    return 0


def _resolve_seed(seed: int | None) -> int | None:
    if seed is not None:
        return seed  # --seed prevalece: no se lee ninguna configuración
    try:
        config = load_config(DEMO_CONFIG)
    except ConfigError as exc:
        print(str(exc), file=sys.stderr)
        return None
    if config.sample is None:
        print("error: no hay semilla: use --seed o añada [sample] a config/demo.toml", file=sys.stderr)
        return None
    return config.sample.seed


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.command == "ingest":
        return _run_ingest(args.config, args.input_dir)
    if args.command != "generate-sample":
        print(PENDING_MESSAGE)
        return 0
    seed = _resolve_seed(args.seed)
    if seed is None:
        return 2
    try:
        result = generate(seed, args.output_root, overwrite=args.overwrite)
    except GenerateError as exc:
        print(f"error: {exc.message}", file=sys.stderr)
        return 2
    print(
        f"Conjunto sintético publicado y validado: {result['records']} registros, {result['files']} archivos "
        f"en {result['output_root']} (semilla {result['seed']}; sha256 del manifiesto {result['manifest_sha256']})."
    )
    return 0
