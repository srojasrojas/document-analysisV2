from __future__ import annotations

import argparse
import hashlib
import re
import shutil
import sys
from pathlib import Path

from .config import load_config, load_env_file, project_root_from_config, resolve_path
from .docx_io import apply_surgical_change, collect_elements, load_docx, normalize_text, save_docx
from .embedded_artifacts import excluded_locations, run_embedded_artifact_cleanup
from .format_normalization import run_format_normalization
from .llm_refine import LlmRefiner
from .models import ChangeRecord, EmbeddedArtifactRecord, FormatNormalizationRecord, PassSummary, SkipRecord
from .reporting import (
    append_changes_jsonl,
    write_audit_report,
    write_embedded_artifact_report,
    write_format_normalization_report,
    write_registry,
)
from .rules import ReplacementRule, load_rules


ACTION_VERB_RE = re.compile(
    r"\b(?:debe(?:n|r[aá]n?)?|realiza(?:r[aá]n?|r|n|ndo)?|revisa(?:r[aá]n?|r|n|ndo)?|"
    r"verifica(?:r[aá]n?|r|n|ndo)?|detiene(?:n)?|detendr[aá]n?|opera(?:r[aá]n?|r|n|ndo)?|"
    r"inspecciona(?:r[aá]n?|r|n|ndo)?|coordina(?:r[aá]n?|r|n|ndo)?|informa(?:r[aá]n?|r|n|ndo)?|"
    r"avisa(?:r[aá]n?|r|n|ndo)?|solicita(?:r[aá]n?|r|n|ndo)?|autoriza(?:r[aá]n?|r|n|ndo)?|"
    r"registra(?:r[aá]n?|r|n|ndo)?|bloquea(?:r[aá]n?|r|n|ndo)?|desbloquea(?:r[aá]n?|r|n|ndo)?|"
    r"energiza(?:r[aá]n?|r|n|ndo)?|desenergiza(?:r[aá]n?|r|n|ndo)?|comunica(?:r[aá]n?|r|n|ndo)?|"
    r"asegura(?:r[aá]n?|r|n|ndo)?|conozca(?:n)?|hace(?:n|r)?|har[aá]n?|procede(?:n|r)?|"
    r"proceder[aá]n?|coloca(?:r[aá]n?|r|n|ndo)?|pesa(?:r[aá]n?|r|n|ndo)?|"
    r"entrega(?:r[aá]n?|r|n|ndo)?|instala(?:r[aá]n?|r|n|ndo)?|retira(?:r[aá]n?|r|n|ndo)?|"
    r"posiciona(?:r[aá]n?|r|n|ndo)?|ubica(?:r[aá]n?|r|n|ndo)?|"
    r"conduce(?:n)?|conducir[aá]n?|traslada(?:r[aá]n?|r|n|ndo)?|"
    r"es\s+(?:el\s+)?responsable|son\s+(?:los\s+)?responsables|encargad[oa]s?\s+de|"
    r"da\s+aviso|dar[aá]?\s+aviso|dando\s+aviso)\b",
    re.IGNORECASE,
)
REPAIRABLE_QA_FLAGS = {
    "changed_in_registros",
    "operator_cas_context",
    "target_already_present_near_match",
    "title_like_expansion",
    "duplicate_target",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Apply configurable replacement rules to editable DOCX files.")
    parser.add_argument("--config", default="config.yaml", help="Path to config.yaml")
    parser.add_argument("--input", type=Path, help="DOCX file or directory. Defaults to paths.input_dir")
    parser.add_argument("--output", type=Path, help="Output directory. Defaults to paths.output_dir")
    parser.add_argument("--passes", type=int, help="Override pipeline.max_passes")
    parser.add_argument("--dry-run", action="store_true", help="Scan without writing documents or reports")
    parser.add_argument("--simple-only", action="store_true", help="Disable optional LLM refining")
    parser.add_argument("--force", action="store_true", help="Recreate output files from the input documents")
    parser.add_argument(
        "--skip-embedded-cleanup",
        action="store_true",
        help="Disable configured cleanup of body-embedded headers and footers",
    )
    parser.add_argument(
        "--skip-format-normalization",
        action="store_true",
        help="Disable the stage-zero format normalization of PDF conversion artifacts",
    )
    parser.add_argument(
        "--snapshot-dir",
        type=Path,
        default=None,
        metavar="DIR",
        help="Save a DOCX copy after each pipeline stage into DIR (format_norm, embedded_cleanup, rule_pass_N)",
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


def _candidate_id(document_name: str, location: str, rule_id: str, text: str) -> str:
    payload = f"{document_name}\0{location}\0{rule_id}\0{text}"
    return hashlib.sha1(payload.encode("utf-8")).hexdigest()[:16]


def _section_text(section_path: tuple[str, ...]) -> str:
    return " > ".join(section_path)


def _run_rule_pass(
    *,
    doc_path: Path,
    document_name: str,
    pass_index: int,
    rules: list[ReplacementRule],
    skip_heading_styles: bool,
    dry_run: bool,
    excluded_element_locations: set[str] | None = None,
) -> tuple[list[ChangeRecord], list[SkipRecord], PassSummary]:
    doc = load_docx(doc_path)
    elements = collect_elements(doc)
    changes: list[ChangeRecord] = []
    skips: list[SkipRecord] = []
    summary = PassSummary(pass_index=pass_index)

    excluded_element_locations = excluded_element_locations or set()
    for element in elements:
        if element.location in excluded_element_locations:
            continue
        if skip_heading_styles and element.is_heading:
            continue
        current_text = element.text
        for rule in rules:
            if not rule.enabled or not rule.has_candidate(current_text):
                continue
            decision = rule.apply(
                current_text,
                section_path=element.section_path,
                in_excluded_section=element.in_excluded_section,
            )
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
                        rule_category=rule.category,
                        source="rule",
                        reason=decision.reason,
                        match_text=decision.match_text,
                        selected_target=decision.selected_target,
                        selector_reason=decision.selector_reason,
                        context_excerpt=decision.context_excerpt,
                        block_type=element.block_type,
                        section_path=element.section_path,
                        candidate_id=_candidate_id(document_name, element.location, rule.id, current_text),
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
                        skip_type=decision.skip_type,
                        match_text=decision.match_text,
                        selected_target=decision.selected_target,
                        context_excerpt=decision.context_excerpt,
                        block_type=element.block_type,
                        section_path=element.section_path,
                        candidate_id=_candidate_id(document_name, element.location, rule.id, current_text),
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
    rules: list[ReplacementRule],
    dry_run: bool,
    excluded_element_locations: set[str] | None = None,
) -> tuple[list[ChangeRecord], list[SkipRecord], PassSummary]:
    doc = load_docx(doc_path)
    elements = collect_elements(doc)
    changes: list[ChangeRecord] = []
    skips: list[SkipRecord] = []
    summary = PassSummary(pass_index=pass_index)

    if not llm_refiner.can_run():
        return changes, skips, summary

    excluded_element_locations = excluded_element_locations or set()
    for element in elements:
        if element.location in excluded_element_locations:
            continue
        current_text = element.text
        for rule in rules:
            if not rule.enabled or not rule.llm.enabled or not rule.has_candidate(current_text):
                continue
            preflight = rule.apply(
                current_text,
                section_path=element.section_path,
                in_excluded_section=element.in_excluded_section,
            )
            if preflight.changed or preflight.skipped or preflight.already_expanded:
                continue
            if rule.target_phrase.lower() in current_text.lower():
                continue
            summary.llm_attempted += 1
            try:
                result = llm_refiner.refine_for_rule(current_text, rule)
            except Exception as exc:  # noqa: BLE001
                skips.append(
                    SkipRecord(
                        document_name=document_name,
                        pass_index=pass_index,
                        location=element.location,
                        text=current_text,
                        rule_id=rule.id,
                        reason=str(exc),
                        skip_type="llm_error",
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
                        original_text=current_text,
                        modified_text=result.modified_text,
                        rule_id=rule.id,
                        rule_category=rule.category,
                        source="llm",
                        reason=result.reason,
                        block_type=element.block_type,
                        section_path=element.section_path,
                        candidate_id=_candidate_id(document_name, element.location, rule.id, current_text),
                    )
                )
                summary.llm_changed += 1
                summary.changed += 1
                current_text = result.modified_text
                if not dry_run:
                    apply_surgical_change(element, current_text)
            else:
                skips.append(
                    SkipRecord(
                        document_name=document_name,
                        pass_index=pass_index,
                        location=element.location,
                        text=current_text,
                        rule_id=rule.id,
                        reason=result.reason,
                        skip_type="llm_gate_or_validation",
                        source="llm",
                        block_type=element.block_type,
                        section_path=element.section_path,
                        candidate_id=_candidate_id(document_name, element.location, rule.id, current_text),
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
    snapshot_dir: Path | None = None,
) -> tuple[
    Path,
    list[ChangeRecord],
    list[SkipRecord],
    list[PassSummary],
    list[EmbeddedArtifactRecord],
    list[FormatNormalizationRecord],
]:
    rules = load_rules(config)
    if not rules:
        raise ValueError("No replacement rules configured")
    output_path = _output_path_for(input_path, output_dir)
    working_path = _prepare_working_copy(input_path, output_path, force=force, dry_run=dry_run)
    all_changes: list[ChangeRecord] = []
    all_skips: list[SkipRecord] = []
    summaries: list[PassSummary] = []
    cleanup_records: list[EmbeddedArtifactRecord] = []
    normalization_records: list[FormatNormalizationRecord] = []

    def _snap(stage_name: str) -> None:
        if snapshot_dir is None or dry_run:
            return
        snapshot_dir.mkdir(parents=True, exist_ok=True)
        dest = snapshot_dir / f"{stage_name}.docx"
        shutil.copy2(working_path, dest)
        print(f"[snapshot] {dest.name}", file=sys.stderr)

    _snap("00_original")

    skip_heading_styles = bool(config.get("pipeline", {}).get("skip_heading_styles", True))
    normalization_cfg = config.get("pipeline", {}).get("format_normalization", {})
    if bool(normalization_cfg.get("enabled", False)) and bool(normalization_cfg.get("apply_before_rules", True)):
        normalization_records = run_format_normalization(
            working_path,
            document_name=input_path.name,
            config=normalization_cfg,
            dry_run=dry_run,
        )
        total_actions = sum(record.count for record in normalization_records)
        print(
            f"[format-normalization] {input_path.name}: actions={total_actions} "
            f"({', '.join(f'{record.action}={record.count}' for record in normalization_records if record.count)})",
            file=sys.stderr,
        )
        _snap("01_after_format_normalization")

    cleanup_cfg = config.get("pipeline", {}).get("embedded_header_footer_cleanup", {})
    if bool(cleanup_cfg.get("enabled", False)) and bool(cleanup_cfg.get("apply_before_rules", True)):
        cleanup_records = run_embedded_artifact_cleanup(
            working_path,
            document_name=input_path.name,
            config=cleanup_cfg,
            dry_run=dry_run,
        )
        applied_count = sum(1 for record in cleanup_records if record.applied)
        cleanup_actionable_actions = {
            "remove",
            "clear_text",
            "exclude",
            "remove_table",
            "move_to_header",
            "write_header",
            "write_footer",
            "would_remove",
            "would_clear_text",
            "would_remove_table",
            "would_move_to_header",
            "would_write_header",
            "would_write_footer",
        }
        actionable_count = sum(
            1
            for record in cleanup_records
            if record.action in cleanup_actionable_actions
        )
        print(
            f"[embedded-cleanup] {input_path.name}: candidates={len(cleanup_records)} "
            f"actionable={actionable_count} applied={applied_count}",
            file=sys.stderr,
        )
        _snap("02_after_embedded_cleanup")

    cleanup_excluded_locations = excluded_locations(cleanup_records)
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
            excluded_element_locations=cleanup_excluded_locations,
        )
        all_changes.extend(changes)
        all_skips.extend(skips)

        if llm_refiner and not changes:
            llm_changes, llm_skips, llm_summary = _run_llm_pass(
                doc_path=working_path,
                document_name=input_path.name,
                pass_index=pass_index,
                llm_refiner=llm_refiner,
                rules=rules,
                dry_run=dry_run,
                excluded_element_locations=cleanup_excluded_locations,
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
        _snap(f"0{2 + pass_index}_after_rule_pass_{pass_index}")
        if summary.changed == 0:
            break

    return working_path, all_changes, all_skips, summaries, cleanup_records, normalization_records


def _configured_targets(config: dict) -> list[str]:
    targets: list[str] = []
    for rule in config.get("rules", []):
        replacement = rule.get("replacement", {})
        target = replacement.get("target_phrase")
        if target and target not in targets:
            targets.append(str(target))
        for conditional in replacement.get("conditional_targets", []):
            conditional_target = conditional.get("target_phrase")
            if conditional_target and conditional_target not in targets:
                targets.append(str(conditional_target))
    return targets


def _is_title_like_change(change: ChangeRecord) -> bool:
    normalized = normalize_text(change.original_text)
    if len(normalized) > 160 or ACTION_VERB_RE.search(change.original_text):
        return False
    return bool(
        re.search(
            r"\b(?:operador(?:a|es|as)?|supervisor(?:a|es|as)?|jefe\s+de\s+[aá]rea|due[ñn]o\s+de\s+[aá]rea)\b",
            change.original_text,
            re.IGNORECASE,
        )
    )


def _qa_flags_for_change(change: ChangeRecord, targets: list[str]) -> list[str]:
    flags: list[str] = []
    section = normalize_text(_section_text(change.section_path))
    match = re.search(r"expanded\s+(\d+)\s+match", change.reason)
    repairs_descriptor = "repaired" in change.reason and "descriptor" in change.reason
    updates_existing_target = "updated" in change.reason and "target phrase" in change.reason
    expected_new_targets = int(match.group(1)) if match else 0 if repairs_descriptor else 1

    if change.source == "llm":
        flags.append("llm_used")
    if re.search(r"\bregistros?\b", section) and not (repairs_descriptor or updates_existing_target):
        flags.append("changed_in_registros")
    if change.rule_id.startswith("operador") and re.search(
        r"\b(?:CAS|CIO)\b|sala\s+de\s+control", change.match_text, re.IGNORECASE
    ):
        flags.append("operator_cas_context")
    if not (repairs_descriptor or updates_existing_target) and _is_title_like_change(change):
        flags.append("title_like_expansion")
    if not repairs_descriptor and _target_already_present_near_match(change, targets):
        flags.append("target_already_present_near_match")
    for target in targets:
        if target and change.modified_text.count(target) > change.original_text.count(target) + expected_new_targets:
            flags.append("duplicate_target")
            break
    return list(dict.fromkeys(flags))


def _target_already_present_near_match(change: ChangeRecord, targets: list[str]) -> bool:
    original_norm = normalize_text(change.original_text)
    selected_targets = [target.strip() for target in change.selected_target.split(" | ") if target.strip()]
    targets_to_check = selected_targets or targets
    for match_text in change.match_text.split(" | "):
        match_norm = normalize_text(match_text)
        if not match_norm:
            continue
        index = original_norm.find(match_norm)
        if index == -1:
            continue
        after = original_norm[index + len(match_norm) : index + len(match_norm) + 180]
        after_same_sentence = re.split(r"[\r\n.;:]", after, maxsplit=1)[0]
        for target in targets_to_check:
            target_norm = normalize_text(target)
            if target_norm and (target_norm in match_norm or target_norm in after_same_sentence):
                return True
    return False


def _repair_flagged_changes(
    *,
    output_by_document: dict[str, Path],
    changes: list[ChangeRecord],
) -> dict[str, bool]:
    repaired: dict[str, bool] = {}
    changes_by_document: dict[str, list[tuple[int, ChangeRecord]]] = {}
    for index, change in enumerate(changes):
        if not REPAIRABLE_QA_FLAGS.intersection(change.qa_flags):
            continue
        changes_by_document.setdefault(change.document_name, []).append((index, change))

    for document_name, document_changes in changes_by_document.items():
        output_path = output_by_document.get(document_name)
        if not output_path or not output_path.exists():
            continue
        doc = load_docx(output_path)
        elements = {element.location: element for element in collect_elements(doc)}
        doc_changed = False
        doc_repaired_changes: list[ChangeRecord] = []
        for _, change in sorted(document_changes, key=lambda item: (item[1].pass_index, item[0]), reverse=True):
            element = elements.get(change.location)
            if element is None or element.text != change.modified_text:
                repaired[change.candidate_id] = False
                continue
            apply_surgical_change(element, change.original_text)
            change.qa_flags.append("auto_reverted")
            repaired[change.candidate_id] = True
            doc_repaired_changes.append(change)
            doc_changed = True
        if doc_changed:
            try:
                save_docx(doc, output_path)
            except OSError:
                for change in doc_repaired_changes:
                    change.qa_flags = [flag for flag in change.qa_flags if flag != "auto_reverted"]
                    change.qa_flags.append("auto_repair_failed")
                    repaired[change.candidate_id] = False
    return repaired


def _run_post_audit(
    *,
    output_by_document: dict[str, Path],
    changes: list[ChangeRecord],
    skips: list[SkipRecord],
    config: dict,
    auto_repair: bool,
) -> list[dict]:
    targets = _configured_targets(config)
    for change in changes:
        change.qa_flags = _qa_flags_for_change(change, targets)

    repaired = _repair_flagged_changes(output_by_document=output_by_document, changes=changes) if auto_repair else {}
    rows: list[dict] = []
    for index, change in enumerate(changes, start=1):
        rows.append(
            {
                "index": index,
                "document_name": change.document_name,
                "location": change.location,
                "rule_id": change.rule_id,
                "source": change.source,
                "qa_flags": ", ".join(change.qa_flags),
                "auto_repaired": repaired.get(change.candidate_id, False),
                "section_path": _section_text(change.section_path),
                "match_text": change.match_text,
                "selected_target": change.selected_target,
                "original_text": change.original_text,
                "modified_text": change.modified_text,
            }
        )
    for skip in skips:
        if "skip_section" in skip.skip_type or "skip_review_only" in skip.skip_type:
            skip.qa_flags = ["expected_skip"]
    return rows


def main() -> int:
    args = parse_args()
    config_path = Path(args.config).resolve()
    project_root = project_root_from_config(config_path)
    load_env_file(project_root)
    config = load_config(config_path)
    paths_cfg = config.get("paths", {})
    pipeline_cfg = config.get("pipeline", {})
    if args.skip_embedded_cleanup:
        config.setdefault("pipeline", {}).setdefault("embedded_header_footer_cleanup", {})["enabled"] = False
    if args.skip_format_normalization:
        config.setdefault("pipeline", {}).setdefault("format_normalization", {})["enabled"] = False

    input_path = args.input or resolve_path(project_root, paths_cfg.get("input_dir", "data/input"))
    if not input_path.is_absolute():
        input_path = project_root / input_path
    output_dir = args.output or resolve_path(project_root, paths_cfg.get("output_dir", "data/output"))
    if not output_dir.is_absolute():
        output_dir = project_root / output_dir
    max_passes = args.passes if args.passes is not None else int(pipeline_cfg.get("max_passes", 3))

    changes_path = resolve_path(project_root, paths_cfg.get("changes_report", "reports/changes.jsonl"))
    registry_path = resolve_path(project_root, paths_cfg.get("registry_report", "reports/registro_cambios.xlsx"))
    audit_json_path = resolve_path(project_root, paths_cfg.get("audit_report_json", "reports/auditoria_post_run.json"))
    audit_excel_path = resolve_path(project_root, paths_cfg.get("audit_report_excel", "reports/auditoria_post_run.xlsx"))
    embedded_report_jsonl_path = resolve_path(
        project_root,
        paths_cfg.get("embedded_header_footer_report_jsonl", "reports/embedded_header_footer_cleanup.jsonl"),
    )
    embedded_report_excel_path = resolve_path(
        project_root,
        paths_cfg.get("embedded_header_footer_report_excel", "reports/embedded_header_footer_cleanup.xlsx"),
    )
    normalization_report_jsonl_path = resolve_path(
        project_root,
        paths_cfg.get("format_normalization_report_jsonl", "reports/format_normalization.jsonl"),
    )
    normalization_report_excel_path = resolve_path(
        project_root,
        paths_cfg.get("format_normalization_report_excel", "reports/format_normalization.xlsx"),
    )

    docx_inputs = _find_inputs(input_path)
    if not docx_inputs:
        print(f"No .docx files found in {input_path}", file=sys.stderr)
        return 1

    all_changes: list[ChangeRecord] = []
    all_skips: list[SkipRecord] = []
    all_cleanup_records: list[EmbeddedArtifactRecord] = []
    all_normalization_records: list[FormatNormalizationRecord] = []
    output_by_document: dict[str, Path] = {}
    for docx_path in docx_inputs:
        print(f"[document] {docx_path.name}", file=sys.stderr)
        snapshot_dir = args.snapshot_dir
        if snapshot_dir and len(docx_inputs) > 1:
            snapshot_dir = args.snapshot_dir / docx_path.stem
        working_path, changes, skips, _summaries, cleanup_records, normalization_records = run_document(
            input_path=docx_path,
            output_dir=output_dir,
            config=config,
            project_root=project_root,
            max_passes=max_passes,
            dry_run=args.dry_run,
            simple_only=args.simple_only,
            force=args.force or bool(pipeline_cfg.get("overwrite_output", False)),
            snapshot_dir=snapshot_dir,
        )
        all_changes.extend(changes)
        all_skips.extend(skips)
        all_cleanup_records.extend(cleanup_records)
        all_normalization_records.extend(normalization_records)
        output_by_document[docx_path.name] = working_path
        print(f"[document] output={working_path}", file=sys.stderr)

    if not args.dry_run:
        if (args.force or bool(pipeline_cfg.get("overwrite_output", False))) and bool(
            pipeline_cfg.get("reset_reports_on_force", True)
        ):
            changes_path.unlink(missing_ok=True)
        post_audit_cfg = pipeline_cfg.get("post_audit", {})
        audit_rows: list[dict] = []
        if bool(post_audit_cfg.get("enabled", True)):
            audit_rows = _run_post_audit(
                output_by_document=output_by_document,
                changes=all_changes,
                skips=all_skips,
                config=config,
                auto_repair=bool(post_audit_cfg.get("auto_repair", True)),
            )
            write_audit_report(audit_rows, audit_excel_path, audit_json_path)
        cleanup_cfg = config.get("pipeline", {}).get("embedded_header_footer_cleanup", {})
        if bool(cleanup_cfg.get("enabled", False)):
            write_embedded_artifact_report(all_cleanup_records, embedded_report_excel_path, embedded_report_jsonl_path)
        normalization_cfg = config.get("pipeline", {}).get("format_normalization", {})
        if bool(normalization_cfg.get("enabled", False)):
            write_format_normalization_report(
                all_normalization_records, normalization_report_excel_path, normalization_report_jsonl_path
            )
        append_changes_jsonl(all_changes, changes_path)
        write_registry([], all_skips, registry_path, existing_jsonl=changes_path, audit_rows=audit_rows)

    print(
        f"Completed: {len(all_changes)} change(s), {len(all_skips)} skipped/review item(s), "
        f"{len(all_cleanup_records)} embedded artifact candidate(s), "
        f"{sum(record.count for record in all_normalization_records)} format normalization action(s).",
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())