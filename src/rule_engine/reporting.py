from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

import xlsxwriter

from .models import ChangeRecord, SkipRecord


def append_changes_jsonl(changes: list[ChangeRecord], path: Path) -> None:
    if not changes:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as handle:
        for change in changes:
            handle.write(json.dumps(change.to_dict(), ensure_ascii=False))
            handle.write("\n")


def load_change_dicts(path: Path) -> list[dict]:
    records: list[dict] = []
    if not path.exists():
        return records
    with path.open("r", encoding="utf-8") as handle:
        for line in handle:
            line = line.strip()
            if not line:
                continue
            try:
                records.append(json.loads(line))
            except json.JSONDecodeError:
                continue
    return records


def write_registry(
    changes: list[ChangeRecord],
    skips: list[SkipRecord],
    registry_path: Path,
    existing_jsonl: Path | None = None,
    audit_rows: list[dict] | None = None,
) -> None:
    registry_path.parent.mkdir(parents=True, exist_ok=True)
    rows: list[dict] = []
    if existing_jsonl:
        rows.extend(load_change_dicts(existing_jsonl))
    rows.extend(change.to_dict() for change in changes)

    workbook = xlsxwriter.Workbook(str(registry_path))
    summary_sheet = workbook.add_worksheet("Resumen")
    changes_sheet = workbook.add_worksheet("Cambios")
    skips_sheet = workbook.add_worksheet("Omitidos")
    review_sheet = workbook.add_worksheet("Revision")
    qa_sheet = workbook.add_worksheet("QA")
    header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7"})
    wrap_format = workbook.add_format({"text_wrap": True, "valign": "top"})
    warning_format = workbook.add_format({"text_wrap": True, "valign": "top", "bg_color": "#FFF2CC"})

    by_document = Counter(str(row.get("document_name", "")) for row in rows)
    by_rule = Counter(str(row.get("rule_id", "")) for row in rows)
    by_target = Counter(str(row.get("selected_target", "")) for row in rows if row.get("selected_target"))
    summary_rows = [
        ("total_changes", len(rows)),
        ("total_skips_current_run", len(skips)),
        ("documents_with_changes", len(by_document)),
        ("rules_with_changes", len(by_rule)),
    ]
    summary_row = 0
    summary_sheet.write(summary_row, 0, "Metric", header_format)
    summary_sheet.write(summary_row, 1, "Value", header_format)
    for summary_row, (metric, value) in enumerate(summary_rows, start=1):
        summary_sheet.write(summary_row, 0, metric)
        summary_sheet.write(summary_row, 1, value)
    summary_row += 2
    summary_sheet.write(summary_row, 0, "Document", header_format)
    summary_sheet.write(summary_row, 1, "Changes", header_format)
    for offset, (document_name, count) in enumerate(by_document.most_common(), start=summary_row + 1):
        summary_sheet.write(offset, 0, document_name)
        summary_sheet.write(offset, 1, count)
    target_start = summary_row
    summary_sheet.write(target_start, 3, "Selected target", header_format)
    summary_sheet.write(target_start, 4, "Changes", header_format)
    for offset, (target, count) in enumerate(by_target.most_common(), start=target_start + 1):
        summary_sheet.write(offset, 3, target)
        summary_sheet.write(offset, 4, count)
    summary_sheet.set_column(0, 0, 42)
    summary_sheet.set_column(1, 4, 24)

    change_headers = [
        "document_name",
        "pass_index",
        "location",
        "block_type",
        "section_path",
        "rule_id",
        "rule_category",
        "source",
        "reason",
        "match_text",
        "selected_target",
        "selector_reason",
        "context_excerpt",
        "original_text",
        "modified_text",
        "candidate_id",
        "qa_flags",
        "changed_at",
    ]
    for column, header in enumerate(change_headers):
        changes_sheet.write(0, column, header, header_format)
    for row_index, row in enumerate(rows, start=1):
        for column, header in enumerate(change_headers):
            value = row.get(header, "")
            if isinstance(value, (list, tuple)):
                value = " > ".join(str(item) for item in value)
            cell_format = warning_format if header == "qa_flags" and value else wrap_format
            changes_sheet.write(row_index, column, value, cell_format)
    changes_sheet.set_column(0, 12, 24)
    changes_sheet.set_column(13, 14, 70)
    changes_sheet.set_column(15, 17, 22)

    skip_headers = [
        "document_name",
        "pass_index",
        "location",
        "block_type",
        "section_path",
        "rule_id",
        "source",
        "skip_type",
        "reason",
        "match_text",
        "selected_target",
        "context_excerpt",
        "text",
        "candidate_id",
        "qa_flags",
    ]
    for column, header in enumerate(skip_headers):
        skips_sheet.write(0, column, header, header_format)
        review_sheet.write(0, column, header, header_format)
    for row_index, skip in enumerate(skips, start=1):
        row = skip.to_dict()
        for column, header in enumerate(skip_headers):
            value = row.get(header, "")
            if isinstance(value, (list, tuple)):
                value = " > ".join(str(item) for item in value)
            skips_sheet.write(row_index, column, value, wrap_format)
            if "review" in str(row.get("skip_type", "")) or "no_action" in str(row.get("skip_type", "")):
                review_sheet.write(row_index, column, value, wrap_format)
    skips_sheet.set_column(0, 11, 24)
    skips_sheet.set_column(12, 12, 90)
    review_sheet.set_column(0, 11, 24)
    review_sheet.set_column(12, 12, 90)

    qa_headers = [
        "index",
        "document_name",
        "location",
        "rule_id",
        "source",
        "qa_flags",
        "auto_repaired",
        "section_path",
        "match_text",
        "selected_target",
        "original_text",
        "modified_text",
    ]
    for column, header in enumerate(qa_headers):
        qa_sheet.write(0, column, header, header_format)
    for row_index, row in enumerate(audit_rows or [], start=1):
        for column, header in enumerate(qa_headers):
            qa_sheet.write(row_index, column, row.get(header, ""), wrap_format)
    qa_sheet.set_column(0, 9, 24)
    qa_sheet.set_column(10, 11, 80)
    workbook.close()


def write_audit_report(audit_rows: list[dict], excel_path: Path, json_path: Path) -> None:
    excel_path.parent.mkdir(parents=True, exist_ok=True)
    json_path.parent.mkdir(parents=True, exist_ok=True)
    json_path.write_text(json.dumps(audit_rows, ensure_ascii=False, indent=2), encoding="utf-8")

    workbook = xlsxwriter.Workbook(str(excel_path))
    sheet = workbook.add_worksheet("Auditoria")
    header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7"})
    wrap_format = workbook.add_format({"text_wrap": True, "valign": "top"})
    warning_format = workbook.add_format({"text_wrap": True, "valign": "top", "bg_color": "#FFF2CC"})
    headers = [
        "index",
        "document_name",
        "location",
        "rule_id",
        "source",
        "qa_flags",
        "auto_repaired",
        "section_path",
        "match_text",
        "selected_target",
        "original_text",
        "modified_text",
    ]
    for column, header in enumerate(headers):
        sheet.write(0, column, header, header_format)
    for row_index, row in enumerate(audit_rows, start=1):
        for column, header in enumerate(headers):
            value = row.get(header, "")
            cell_format = warning_format if header == "qa_flags" and value else wrap_format
            sheet.write(row_index, column, value, cell_format)
    sheet.set_column(0, 9, 24)
    sheet.set_column(10, 11, 90)
    workbook.close()