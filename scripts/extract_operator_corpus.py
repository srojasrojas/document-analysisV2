from __future__ import annotations

import argparse
import json
import re
from collections import Counter
from pathlib import Path
from typing import Any

import xlsxwriter
from docx import Document

from rule_engine.docx_io import collect_elements, normalize_text


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_RUN_DIR = PROJECT_ROOT / "tmp" / "run_b25"
DEFAULT_OUTPUT = PROJECT_ROOT / "reports" / "operator_corpus.xlsx"

OPERATOR_RE = re.compile(
    r"(?i)(?<![A-Za-z0-9_])operador(?:\s*/\s*a|\s*\([aA]\)|a|es|as)?"
    r"(?:\s+a\s+cargo)?(?:\s+(?:(?:de|del)\s+)?"
    r"(?:planta(?:\s+concentradora)?|equipos?|terreno|sala(?:\s+de\s+control)?|control|"
    r"zona\s+aut[oó]noma|puentes?\s+gr[uú]as?|gr[uú]a(?:\s+horquilla)?|"
    r"patio(?:\s+(?:de\s+)?(?:embarque|c[aá]todos))?(?:\s+Spence)?|"
    r"cami[oó]n(?:es)?(?:\s+(?:tolva|pluma))?|camioneta|mina|[aá]rea|"
    r"proceso|procesos|chancado|molienda|flotaci[oó]n|relaves|maquinaria|operaciones|"
    r"correas?|apilador(?:a)?|esparcidor|retro\s*-?\s*excavadora(?:s)?|excavador(?:a|es)?|"
    r"mini\s*-?\s*cargador(?:a|es)?|cargador(?:a|es)?(?:\s+frontal)?|rotopalas?|picarocas?|"
    r"nave\s+ew|c[aá]todos|sx|tf|cas|Spence(?:\s+debidamente\s+certificad[oa]s?(?:\s+en\s+operaci[oó]n\s+de\s+puentes?\s+gr[uú]as?)?)?|"
    r"(?:MLDC|MDC|EW)(?:\s+Spence)?|m[aá]quina\s+despegadora(?:\s+de\s+c[aá]todos)?|"
    r"otras\s+[aá]reas|circuitos?(?:\s+de\s+EW)?|"
    r"encargad[oa]\s+(?:de\s+los\s+circuitos\s+de\s+EW|del\s+puente\s+gr[uú]a)))?"
    r"(?:\s+(?:autorizad[oa]s?|calificad[oa]s?|capacitad[oa]s?|certificad[oa]s?|"
    r"competente(?:s)?|habilitad[oa]s?|acreditad[oa]s?|entrenad[oa]s?|designad[oa]s?))?"
    r"(?![A-Za-z0-9_])"
)
ACTION_RE = re.compile(
    r"(?i)\b(?:debe(?:r[aá]n?|r[aá])?|realiza(?:r|n)?|revisa(?:r|n)?|verifica(?:r|n)?|"
    r"detiene(?:r|n)?|opera(?:r|n)?|inspecciona(?:r|n)?|coordina(?:r|n)?|informa(?:r|n)?|"
    r"avisa(?:r|n)?|solicita(?:r|n)?|autoriza(?:r|n)?|registra(?:r|n)?|bloquea(?:r|n)?|"
    r"desbloquea(?:r|n)?|energiza(?:r|n)?|desenergiza(?:r|n)?|comunica(?:r|n)?|dar\s+aviso\s+a)\b"
)
QUALIFIED_RE = re.compile(
    r"(?i)\b(?:autorizad[oa]s?|calificad[oa]s?|capacitad[oa]s?|"
    r"competente(?:s)?|habilitad[oa]s?|acreditad[oa]s?|entrenad[oa]s?)\b"
)
CERTIFIED_EQUIPMENT_RE = re.compile(
    r"(?i)\b(?:retro\s*-?\s*excavadora(?:s)?|mini\s*-?\s*cargador(?:a|es)?|"
    r"cargador(?:a|es)?(?:\s+frontal)?|excavador(?:a|es)?|cami[oó]n(?:es)?\s+(?:tolva|pluma)|"
    r"rotopalas?|apilador(?:a)?|esparcidor|picarocas?|puentes?\s+gr[uú]as?|"
    r"gr[uú]a\s+horquilla|maquinaria|equipos?|certificad[oa]s?)\b"
)
TECHNICAL_RE = re.compile(
    r"(?i)\b(?:bloqueo|bloquear|desbloqueo|desbloquear|loto|energizaci[oó]n|energizar|"
    r"desenergizaci[oó]n|desenergizar|hmi|panel\s+de\s+control|reset|el[eé]ctric[oa]s?|"
    r"instrumentaci[oó]n|mantenci[oó]n|calibraci[oó]n|calibrar|homologaci[oó]n|diagn[oó]stico)\b"
)
CAS_RE = re.compile(r"(?i)\b(?:CAS|CIO)\b|sala\s+de\s+control")
TARGET_RE = re.compile(
    r"(?i)\b(?:personal\s+designado\s+por\s+minera\s+Spence|"
    r"personal\s+certificado\s+designado\s+por\s+minera\s+Spence|personal\s+calificado)\b"
)
ALT_AFTER_RE = re.compile(
    r"(?i)^\s*(?:o|/|y/o)\s+(?!(?:personal\s+designado\s+por\s+minera\s+Spence|"
    r"personal\s+certificado\s+designado\s+por\s+minera\s+Spence|personal\s+calificado))"
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract operator expansion candidates from run_b25 artifacts.")
    parser.add_argument("--run-dir", type=Path, default=DEFAULT_RUN_DIR)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    return parser.parse_args()


def _context(text: str, match: re.Match[str], window: int = 180) -> str:
    start = max(0, match.start() - window)
    end = min(len(text), match.end() + window)
    return re.sub(r"\s+", " ", text[start:end]).strip()


def _classify(text: str, section_path: tuple[str, ...], match: re.Match[str]) -> tuple[str, str, str]:
    context = _context(text, match)
    after = text[match.end() : match.end() + 120]
    section = normalize_text(" > ".join(section_path))

    if re.search(r"\bregistros?\b", section):
        return "skip_section_registros", "", "Section path contains REGISTROS"
    if TARGET_RE.search(after) or TARGET_RE.search(text):
        return "already_expanded", "", "Target already present"
    if ALT_AFTER_RE.search(after):
        return "skip_existing_alternative", "", "Alternative role follows operator"
    if CAS_RE.search(context):
        return "skip_cas_context", "", "CAS/CIO/Sala de Control context"
    if not ACTION_RE.search(context):
        return "review_no_action_context", "", "No responsibility/action verb near match"
    if QUALIFIED_RE.search(context) or TECHNICAL_RE.search(context):
        return "expand_personal_calificado", "personal calificado", "Qualified or technical operator context"
    if CERTIFIED_EQUIPMENT_RE.search(match.group(0)):
        return (
            "expand_personal_certificado_designado",
            "personal certificado designado por Minera Spence",
            "Certified equipment operator context",
        )
    return "expand_spence", "personal designado por Minera Spence", "Operational operator responsibility"


def _docx_candidates(path: Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    rows: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = []
    try:
        doc = Document(str(path))
        elements = collect_elements(doc)
    except Exception as exc:  # noqa: BLE001
        return rows, [{"document_name": path.name, "error": str(exc)}]

    for element in elements:
        for match in OPERATOR_RE.finditer(element.text):
            classification, target, reason = _classify(element.text, element.section_path, match)
            rows.append(
                {
                    "document_name": path.name,
                    "location": element.location,
                    "block_type": element.block_type,
                    "section_path": " > ".join(element.section_path),
                    "match_text": match.group(0),
                    "classification": classification,
                    "target": target,
                    "reason": reason,
                    "context_excerpt": _context(element.text, match),
                    "text": element.text,
                }
            )
    return rows, errors


def _jsonl_operator_refs(path: Path, limit: int = 2000) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    rows: list[dict[str, Any]] = []
    with path.open("r", encoding="utf-8") as handle:
        for line_number, line in enumerate(handle, start=1):
            if len(rows) >= limit:
                break
            try:
                record = json.loads(line)
            except json.JSONDecodeError:
                continue
            joined = " ".join(str(record.get(key, "")) for key in ("original_text", "modified_text", "activity"))
            if not OPERATOR_RE.search(joined):
                continue
            rows.append(
                {
                    "line": line_number,
                    "document_name": record.get("document_name", ""),
                    "stage_id": record.get("stage_id", ""),
                    "rule_applied": record.get("rule_applied", record.get("rule_description", "")),
                    "location": record.get("location", record.get("match_location", "")),
                    "original_text": record.get("original_text", ""),
                    "modified_text": record.get("modified_text", ""),
                    "activity": record.get("activity", ""),
                }
            )
    return rows


def _write_sheet(workbook, name: str, rows: list[dict[str, Any]], headers: list[str]) -> None:
    sheet = workbook.add_worksheet(name[:31])
    header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7"})
    wrap_format = workbook.add_format({"text_wrap": True, "valign": "top"})
    for column, header in enumerate(headers):
        sheet.write(0, column, header, header_format)
    for row_index, row in enumerate(rows, start=1):
        for column, header in enumerate(headers):
            sheet.write(row_index, column, row.get(header, ""), wrap_format)
    sheet.set_column(0, min(len(headers), 8), 24)
    if headers:
        sheet.set_column(max(0, len(headers) - 2), len(headers) - 1, 80)


def write_workbook(path: Path, candidates: list[dict[str, Any]], jsonl_refs: list[dict[str, Any]], errors: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(str(path))
    counts = Counter(row["classification"] for row in candidates)
    variant_counts = Counter(row["match_text"].lower() for row in candidates)
    summary = [{"metric": key, "value": value} for key, value in counts.most_common()]
    summary.insert(0, {"metric": "total_candidates", "value": len(candidates)})
    summary.insert(1, {"metric": "documents_with_candidates", "value": len({row["document_name"] for row in candidates})})

    _write_sheet(workbook, "Resumen", summary, ["metric", "value"])
    _write_sheet(
        workbook,
        "Candidatos",
        candidates,
        [
            "document_name",
            "location",
            "block_type",
            "section_path",
            "match_text",
            "classification",
            "target",
            "reason",
            "context_excerpt",
            "text",
        ],
    )
    _write_sheet(
        workbook,
        "Variantes",
        [{"variant": variant, "count": count} for variant, count in variant_counts.most_common()],
        ["variant", "count"],
    )
    _write_sheet(
        workbook,
        "Personal Calificado",
        [row for row in candidates if row["classification"] == "expand_personal_calificado"],
        ["document_name", "location", "match_text", "target", "reason", "context_excerpt", "text"],
    )
    _write_sheet(
        workbook,
        "Personal Certificado",
        [row for row in candidates if row["classification"] == "expand_personal_certificado_designado"],
        ["document_name", "location", "match_text", "target", "reason", "context_excerpt", "text"],
    )
    _write_sheet(
        workbook,
        "Omitidos",
        [row for row in candidates if row["classification"].startswith("skip") or row["classification"] == "already_expanded"],
        ["document_name", "location", "section_path", "match_text", "classification", "reason", "context_excerpt", "text"],
    )
    _write_sheet(
        workbook,
        "Revision",
        [row for row in candidates if row["classification"].startswith("review")],
        ["document_name", "location", "section_path", "match_text", "classification", "reason", "context_excerpt", "text"],
    )
    _write_sheet(
        workbook,
        "JSONL",
        jsonl_refs,
        ["line", "document_name", "stage_id", "rule_applied", "location", "original_text", "modified_text", "activity"],
    )
    _write_sheet(workbook, "Errores", errors, ["document_name", "error"])
    workbook.close()


def main() -> int:
    args = parse_args()
    run_dir = args.run_dir if args.run_dir.is_absolute() else PROJECT_ROOT / args.run_dir
    output_path = args.output if args.output.is_absolute() else PROJECT_ROOT / args.output
    candidates: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = []
    for docx_path in sorted(run_dir.glob("*.docx")):
        doc_rows, doc_errors = _docx_candidates(docx_path)
        candidates.extend(doc_rows)
        errors.extend(doc_errors)
    jsonl_refs = _jsonl_operator_refs(run_dir / "phase2_changes_merged.jsonl")
    write_workbook(output_path, candidates, jsonl_refs, errors)
    try:
        display_path = output_path.relative_to(PROJECT_ROOT)
    except ValueError:
        display_path = output_path
    print(f"Wrote {display_path} with {len(candidates)} operator candidate(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())