from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

from rule_engine.config import load_config, project_root_from_config, resolve_path
from rule_engine.embedded_artifacts import run_embedded_artifact_cleanup
from rule_engine.reporting import write_embedded_artifact_report


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Inspect or clean body-embedded DOCX headers and footers.")
    parser.add_argument("--config", default="config.yaml", help="Path to config.yaml")
    parser.add_argument("--input", type=Path, required=True, help="DOCX file or directory to inspect")
    parser.add_argument("--output", type=Path, help="Excel report path")
    parser.add_argument("--jsonl", type=Path, help="JSONL report path")
    parser.add_argument("--apply", action="store_true", help="Apply cleanup to copies in --output-dir")
    parser.add_argument("--output-dir", type=Path, help="Directory for cleaned DOCX copies when --apply is used")
    parser.add_argument("--remove-tables", action="store_true", help="Allow repeated metadata tables to be removed")
    parser.add_argument("--no-write-footer", action="store_true", help="Do not reconstruct real DOCX footers")
    parser.add_argument(
        "--overwrite-existing-footer",
        action="store_true",
        help="Replace existing true DOCX footers when reconstructing",
    )
    return parser.parse_args()


def _find_inputs(input_path: Path) -> list[Path]:
    if input_path.is_file():
        if input_path.suffix.lower() != ".docx" or input_path.name.startswith("~$"):
            raise ValueError(f"Input file is not a usable .docx: {input_path}")
        return [input_path]
    if not input_path.exists():
        raise FileNotFoundError(f"Input path not found: {input_path}")
    return sorted(
        path for path in input_path.glob("**/*.docx") if path.is_file() and not path.name.startswith("~$")
    )


def main() -> int:
    args = parse_args()
    config_path = Path(args.config).resolve()
    project_root = project_root_from_config(config_path)
    config = load_config(config_path)
    paths_cfg = config.get("paths", {})
    cleanup_cfg = dict(config.get("pipeline", {}).get("embedded_header_footer_cleanup", {}))
    cleanup_cfg["enabled"] = True
    cleanup_cfg["action"] = "remove" if args.apply else "preview"
    if args.remove_tables:
        cleanup_cfg["remove_table_artifacts"] = True
    if args.no_write_footer:
        cleanup_cfg["write_real_footer"] = False
    if args.overwrite_existing_footer:
        cleanup_cfg["overwrite_existing_footer"] = True

    input_path = args.input if args.input.is_absolute() else project_root / args.input
    docx_inputs = _find_inputs(input_path)
    if not docx_inputs:
        print(f"No .docx files found in {input_path}", file=sys.stderr)
        return 1

    excel_path = args.output or resolve_path(
        project_root, paths_cfg.get("embedded_header_footer_report_excel", "reports/embedded_header_footer_cleanup.xlsx")
    )
    jsonl_path = args.jsonl or resolve_path(
        project_root, paths_cfg.get("embedded_header_footer_report_jsonl", "reports/embedded_header_footer_cleanup.jsonl")
    )
    output_dir = args.output_dir or project_root / "tmp" / "embedded_header_footer_cleanup"

    records = []
    for docx_path in docx_inputs:
        target_path = docx_path
        if args.apply:
            relative = docx_path.relative_to(input_path) if input_path.is_dir() else Path(docx_path.name)
            target_path = output_dir / relative
            target_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(docx_path, target_path)
        document_records = run_embedded_artifact_cleanup(
            target_path,
            document_name=docx_path.name,
            config=cleanup_cfg,
            dry_run=not args.apply,
        )
        records.extend(document_records)
        applied = sum(1 for record in document_records if record.applied)
        footer_actions = sum(1 for record in document_records if record.action in {"write_footer", "would_write_footer"})
        protected_footers = sum(1 for record in document_records if record.action == "footer_protected_existing")
        print(
            f"[embedded-cleanup] {docx_path.name}: candidates={len(document_records)} "
            f"applied={applied} footer_actions={footer_actions} protected_footers={protected_footers}"
        )

    write_embedded_artifact_report(records, excel_path, jsonl_path)
    print(f"Report written: {excel_path}")
    print(f"JSONL written: {jsonl_path}")
    if args.apply:
        print(f"Cleaned DOCX copies: {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())