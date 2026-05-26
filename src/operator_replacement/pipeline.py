from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

from .config import load_config, load_env_file, project_root_from_config, resolve_path
from .docx_io import apply_surgical_change, collect_elements, load_docx, save_docx
from .llm_refine import LlmRefiner
from .models import ChangeRecord, PassSummary, SkipRecord
from .reporting import append_changes_jsonl, write_registry
from .rules import ReplacementRule, load_rules


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Expand operator mentions in editable DOCX files.")
    parser.add_argument("--config", default="config.yaml", help="Path to config.yaml")
    parser.add_argument("--input", type=Path, help="DOCX file or directory. Defaults to paths.input_dir")
    parser.add_argument("--output", type=Path, help="Output directory. Defaults to paths.output_dir")
    parser.add_argument("--passes", type=int, help="Override pipeline.max_passes")
    parser.add_argument("--dry-run", action="store_true", help="Scan without writing documents or reports")
    parser.add_argument("--simple-only", action="store_true", help="Disable optional LLM refining")
    parser.add_argument("--force", action="store_true", help="Recreate output files from the input documents")
    return parser.parse_args()


def _find_inputs(input_path: Path) -> list[Path]:
    if input_path.is_file():
        if input_path.suffix.lower() != ".docx" or input_path.name.startswith("~$"):
            raise ValueError(f"Input file is not a usable .docx: {input_path}")
        return [input_path]
    if not input_path.exists():
        raise FileNotFoundError(f"Input path not found: {input_path}")
    return sorted(
        path
        for path in input_path.glob("**/*.docx")
        if path.is_file() and not path.name.startswith("~$")
    )


def _output_path_for(input_path: Path, output_dir: Path) -> Path:
    stem = input_path.stem
    if stem.endswith("_modificado"):
        return output_dir / input_path.name
    return output_dir / f"{stem}_modificado.docx"


def _prepare_working_copy(input_path: Path, output_path: Path, *, force: bool, dry_run: bool) -> Path:
    if dry_run:
        return input_path
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.exists() and not force:
        return output_path
    shutil.copy2(input_path, output_path)
    return output_path


def _primary_target_phrase(rules: list[ReplacementRule]) -> str:
    for rule in rules:
        if rule.enabled:
            return rule.target_phrase
    return "personal designado por minera Spence"


def _run_rule_pass(
    *,
    doc_path: Path,
    document_name: str,
    pass_index: int,
    rules: list[ReplacementRule],
    skip_heading_styles: bool,
    dry_run: bool,
) -> tuple[list[ChangeRecord], list[SkipRecord], PassSummary]:
    doc = load_docx(doc_path)
    elements = collect_elements(doc)
    changes: list[ChangeRecord] = []
    skips: list[SkipRecord] = []
    summary = PassSummary(pass_index=pass_index)

    for element in elements:
        if skip_heading_styles and element.is_heading:
            continue
        current_text = element.text
        for rule in rules:
            if not rule.enabled or not rule.has_candidate(current_text):
                continue
            decision = rule.apply(current_text)
            summary.candidates += decision.candidates
            summary.already_expanded += decision.already_expanded
            summary.skipped += decision.skipped
            if decision.changed:
                changes.append(
                    ChangeRecord(
                        document_name=document_name,
                        pass_index=pass_index,
                        location=element.location,
                        original_text=current_text,
                        modified_text=decision.modified_text,
                        rule_id=rule.id,
                        source="rule",
                        reason=decision.reason,
                    )
                )
                current_text = decision.modified_text
                if not dry_run:
                    apply_surgical_change(element, current_text)
                summary.changed += 1
            elif decision.skipped:
                skips.append(
                    SkipRecord(
                        document_name=document_name,
                        pass_index=pass_index,
                        location=element.location,
                        text=current_text,
                        rule_id=rule.id,
                        reason=decision.reason,
                    )
                )

    if changes and not dry_run:
        save_docx(doc, doc_path)
    return changes, skips, summary


def _run_llm_pass(
    *,
    doc_path: Path,
    document_name: str,
    pass_index: int,
    llm_refiner: LlmRefiner,
    target_phrase: str,
    dry_run: bool,
) -> tuple[list[ChangeRecord], list[SkipRecord], PassSummary]:
    doc = load_docx(doc_path)
    elements = collect_elements(doc)
    changes: list[ChangeRecord] = []
    skips: list[SkipRecord] = []
    summary = PassSummary(pass_index=pass_index)

    if not llm_refiner.can_run():
        return changes, skips, summary

    for element in elements:
        text = element.text
        if target_phrase.lower() in text.lower():
            continue
        if not llm_refiner.has_candidate(text):
            continue
        summary.llm_attempted += 1
        try:
            result = llm_refiner.refine(text, target_phrase)
        except Exception as exc:  # noqa: BLE001
            skips.append(
                SkipRecord(
                    document_name=document_name,
                    pass_index=pass_index,
                    location=element.location,
                    text=text,
                    rule_id="llm_refine",
                    reason=str(exc),
                    source="llm",
                )
            )
            continue
        if result.changed:
            changes.append(
                ChangeRecord(
                    document_name=document_name,
                    pass_index=pass_index,
                    location=element.location,
                    original_text=text,
                    modified_text=result.modified_text,
                    rule_id="llm_refine",
                    source="llm",
                    reason=result.reason,
                )
            )
            summary.llm_changed += 1
            summary.changed += 1
            if not dry_run:
                apply_surgical_change(element, result.modified_text)
        else:
            skips.append(
                SkipRecord(
                    document_name=document_name,
                    pass_index=pass_index,
                    location=element.location,
                    text=text,
                    rule_id="llm_refine",
                    reason=result.reason,
                    source="llm",
                )
            )

    if changes and not dry_run:
        save_docx(doc, doc_path)
    return changes, skips, summary


def run_document(
    *,
    input_path: Path,
    output_dir: Path,
    config: dict,
    project_root: Path,
    max_passes: int,
    dry_run: bool,
    simple_only: bool,
    force: bool,
) -> tuple[Path, list[ChangeRecord], list[SkipRecord], list[PassSummary]]:
    rules = load_rules(config)
    if not rules:
        raise ValueError("No replacement rules configured")
    output_path = _output_path_for(input_path, output_dir)
    working_path = _prepare_working_copy(input_path, output_path, force=force, dry_run=dry_run)
    all_changes: list[ChangeRecord] = []
    all_skips: list[SkipRecord] = []
    summaries: list[PassSummary] = []

    skip_heading_styles = bool(config.get("pipeline", {}).get("skip_heading_styles", True))
    target_phrase = _primary_target_phrase(rules)
    use_llm = bool(config.get("pipeline", {}).get("use_llm_refine", False)) and not simple_only
    llm_refiner = None
    if use_llm:
        llm_refiner = LlmRefiner(config, project_root)

    for pass_index in range(1, max_passes + 1):
        changes, skips, summary = _run_rule_pass(
            doc_path=working_path,
            document_name=input_path.name,
            pass_index=pass_index,
            rules=rules,
            skip_heading_styles=skip_heading_styles,
            dry_run=dry_run,
        )
        all_changes.extend(changes)
        all_skips.extend(skips)

        if llm_refiner and not changes:
            llm_changes, llm_skips, llm_summary = _run_llm_pass(
                doc_path=working_path,
                document_name=input_path.name,
                pass_index=pass_index,
                llm_refiner=llm_refiner,
                target_phrase=target_phrase,
                dry_run=dry_run,
            )
            all_changes.extend(llm_changes)
            all_skips.extend(llm_skips)
            summary.llm_attempted += llm_summary.llm_attempted
            summary.llm_changed += llm_summary.llm_changed
            summary.changed += llm_summary.changed

        summaries.append(summary)
        print(
            f"[pass {pass_index}] {input_path.name}: candidates={summary.candidates} "
            f"changed={summary.changed} skipped={summary.skipped} "
            f"already_expanded={summary.already_expanded} llm={summary.llm_changed}/{summary.llm_attempted}",
            file=sys.stderr,
        )
        if summary.changed == 0:
            break

    return working_path, all_changes, all_skips, summaries


def main() -> int:
    args = parse_args()
    config_path = Path(args.config).resolve()
    project_root = project_root_from_config(config_path)
    load_env_file(project_root)
    config = load_config(config_path)
    paths_cfg = config.get("paths", {})
    pipeline_cfg = config.get("pipeline", {})

    input_path = args.input or resolve_path(project_root, paths_cfg.get("input_dir", "data/input"))
    if not input_path.is_absolute():
        input_path = project_root / input_path
    output_dir = args.output or resolve_path(project_root, paths_cfg.get("output_dir", "data/output"))
    if not output_dir.is_absolute():
        output_dir = project_root / output_dir
    max_passes = args.passes or int(pipeline_cfg.get("max_passes", 3))

    changes_path = resolve_path(project_root, paths_cfg.get("changes_report", "reports/changes.jsonl"))
    registry_path = resolve_path(project_root, paths_cfg.get("registry_report", "reports/registro_cambios.xlsx"))

    docx_inputs = _find_inputs(input_path)
    if not docx_inputs:
        print(f"No .docx files found in {input_path}", file=sys.stderr)
        return 1

    all_changes: list[ChangeRecord] = []
    all_skips: list[SkipRecord] = []
    for docx_path in docx_inputs:
        print(f"[document] {docx_path.name}", file=sys.stderr)
        working_path, changes, skips, _summaries = run_document(
            input_path=docx_path,
            output_dir=output_dir,
            config=config,
            project_root=project_root,
            max_passes=max_passes,
            dry_run=args.dry_run,
            simple_only=args.simple_only,
            force=args.force or bool(pipeline_cfg.get("overwrite_output", False)),
        )
        all_changes.extend(changes)
        all_skips.extend(skips)
        print(f"[document] output={working_path}", file=sys.stderr)

    if not args.dry_run:
        append_changes_jsonl(all_changes, changes_path)
        write_registry([], all_skips, registry_path, existing_jsonl=changes_path)

    print(
        f"Completed: {len(all_changes)} change(s), {len(all_skips)} skipped/review item(s).",
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())