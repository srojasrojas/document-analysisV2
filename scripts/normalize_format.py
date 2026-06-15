"""Etapa cero standalone: normaliza el formato de DOCX convertidos desde PDF.

Aplica (sobre copias, sin tocar el original) la misma normalizacion que el
pipeline ejecuta antes de la limpieza de encabezados/pies y de las reglas:
fusion de runs fragmentados, eliminacion del tracking de espaciado del
conversor, homologacion de fuente, colapso de parrafos vacios, aplanado de
tablas anidadas e imagenes flotantes grandes convertidas a inline.

Uso:
    venv\\Scripts\\python.exe scripts\\normalize_format.py --input tmp\\run_b25 \
        --output-dir tmp\\run_b25_normalizado --report reports\\format_normalization_standalone.xlsx
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from rule_engine.config import load_config  # noqa: E402
from rule_engine.format_normalization import (  # noqa: E402
    resolve_normalization_config,
    run_format_normalization,
)
from rule_engine.reporting import write_format_normalization_report  # noqa: E402


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--input", type=Path, required=True, help="DOCX file or directory")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("tmp/format_normalized"),
        help="Directory for the normalized copies (originals are never modified)",
    )
    parser.add_argument("--config", default="config.yaml", help="Optional config.yaml to read defaults from")
    parser.add_argument(
        "--report",
        type=Path,
        default=Path("reports/format_normalization_standalone.xlsx"),
        help="Excel report path (a .jsonl with the same stem is written next to it)",
    )
    parser.add_argument("--dry-run", action="store_true", help="Scan and report without writing documents")
    return parser.parse_args()


def _find_inputs(input_path: Path) -> list[Path]:
    if input_path.is_file():
        return [input_path]
    if not input_path.exists():
        raise FileNotFoundError(f"Input path not found: {input_path}")
    return sorted(
        path for path in input_path.glob("**/*.docx") if path.is_file() and not path.name.startswith("~$")
    )


def main() -> int:
    args = parse_args()
    config_path = Path(args.config)
    normalization_cfg = {}
    if config_path.exists():
        config = load_config(config_path.resolve())
        normalization_cfg = dict(config.get("pipeline", {}).get("format_normalization", {}))
    normalization_cfg = resolve_normalization_config(normalization_cfg)
    normalization_cfg["enabled"] = True

    inputs = _find_inputs(args.input)
    if not inputs:
        print(f"No .docx files found in {args.input}", file=sys.stderr)
        return 1

    records = []
    for input_path in inputs:
        if args.dry_run:
            working_path = input_path
        else:
            args.output_dir.mkdir(parents=True, exist_ok=True)
            working_path = args.output_dir / input_path.name
            shutil.copy2(input_path, working_path)
        document_records = run_format_normalization(
            working_path,
            document_name=input_path.name,
            config=normalization_cfg,
            dry_run=args.dry_run,
        )
        records.extend(document_records)
        total = sum(record.count for record in document_records)
        print(f"[normalize] {input_path.name}: actions={total}", file=sys.stderr)

    jsonl_path = args.report.with_suffix(".jsonl")
    write_format_normalization_report(records, args.report, jsonl_path)
    print(
        f"Completed: {len(inputs)} document(s), {sum(record.count for record in records)} action(s). "
        f"Report: {args.report}",
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
