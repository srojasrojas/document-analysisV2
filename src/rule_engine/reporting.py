from __future__ import annotations

import json
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
) -> None:
    registry_path.parent.mkdir(parents=True, exist_ok=True)
    rows: list[dict] = []
    if existing_jsonl:
        rows.extend(load_change_dicts(existing_jsonl))
    rows.extend(change.to_dict() for change in changes)

    workbook = xlsxwriter.Workbook(str(registry_path))
    changes_sheet = workbook.add_worksheet("cambios")
    skips_sheet = workbook.add_worksheet("omitidos")
    header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7"})
    wrap_format = workbook.add_format({"text_wrap": True, "valign": "top"})

    change_headers = [
        "document_name",
        "pass_index",
        "location",
        "rule_id",
        "rule_category",
        "source",
        "reason",
        "original_text",
        "modified_text",
        "changed_at",
    ]
    for column, header in enumerate(change_headers):
        changes_sheet.write(0, column, header, header_format)
    for row_index, row in enumerate(rows, start=1):
        for column, header in enumerate(change_headers):
            changes_sheet.write(row_index, column, row.get(header, ""), wrap_format)
    changes_sheet.set_column(0, 6, 24)
    changes_sheet.set_column(7, 8, 70)
    changes_sheet.set_column(9, 9, 22)

    skip_headers = ["document_name", "pass_index", "location", "rule_id", "source", "skip_type", "reason", "text"]
    for column, header in enumerate(skip_headers):
        skips_sheet.write(0, column, header, header_format)
    for row_index, skip in enumerate(skips, start=1):
        row = skip.to_dict()
        for column, header in enumerate(skip_headers):
            skips_sheet.write(row_index, column, row.get(header, ""), wrap_format)
    skips_sheet.set_column(0, 6, 24)
    skips_sheet.set_column(7, 7, 90)
    workbook.close()