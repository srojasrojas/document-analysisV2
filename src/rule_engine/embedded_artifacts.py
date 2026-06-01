from __future__ import annotations

import hashlib
import re
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from docx.document import Document as DocxDocument
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from .docx_io import collect_elements, load_docx, normalize_text, save_docx
from .models import EmbeddedArtifactRecord


ACTION_VERB_RE = re.compile(
    r"\b(?:debe(?:n)?|deber[aá]n?|realiza(?:r|n)?|revisa(?:r|n)?|verifica(?:r|n)?|"
    r"detiene(?:r|n)?|opera(?:r|n)?|inspecciona(?:r|n)?|coordina(?:r|n)?|informa(?:r|n)?|"
    r"avisa(?:r|n)?|solicita(?:r|n)?|autoriza(?:r|n)?|registra(?:r|n)?|bloquea(?:r|n)?|"
    r"desbloquea(?:r|n)?|energiza(?:r|n)?|desenergiza(?:r|n)?|comunica(?:r|n)?|asegura(?:r|n)?)\b",
    re.IGNORECASE,
)

MONTH_RE = r"(?:enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)"
TEXT_CHILD_TAGS = {
    qn("w:t"),
    qn("w:delText"),
    qn("w:instrText"),
    qn("w:tab"),
    qn("w:noBreakHyphen"),
    qn("w:softHyphen"),
}

DEFAULT_CLEANUP_CONFIG: dict[str, Any] = {
    "enabled": False,
    "action": "preview",
    "apply_before_rules": True,
    "min_confidence_remove": 0.82,
    "min_confidence_review": 0.45,
    "max_removable_chars": 160,
    "remove_table_artifacts": False,
    "protect_front_matter_tables": True,
    "front_matter_table_count": 1,
    "patterns": [],
    "table_patterns": [],
}

DEFAULT_PATTERNS = [
    r"\beste\s+es\s+un\s+documento\s+controlado\b",
    r"\bfecha\s+de\s+autorizacion\b",
    r"\bproxima\s+revision\b",
    r"\bpagina\s+\d+(?:\s+de\s+\d+)?\b",
]

DEFAULT_TABLE_PATTERNS = [
    r"\bcodigo\b",
    r"\bversion\b",
    r"\belaboro\b",
    r"\breviso\b",
    r"\baprobo\b",
    r"\bfecha\s+de\s+autorizacion\b",
]


@dataclass
class _ParagraphInfo:
    paragraph: Paragraph
    index: int
    location: str
    text: str
    normalized: str
    style_name: str
    alignment: str
    section_path: tuple[str, ...]


@dataclass
class _TableInfo:
    table: Table
    index: int
    location: str
    text: str
    normalized: str
    signature: str


@dataclass
class _ArtifactCandidate:
    record: EmbeddedArtifactRecord
    paragraph: Paragraph | None = None
    table: Table | None = None
    body_index: int | None = None
    table_index: int | None = None
    has_structural_break: bool = False


def resolve_cleanup_config(config: dict[str, Any] | None = None) -> dict[str, Any]:
    resolved = dict(DEFAULT_CLEANUP_CONFIG)
    for key, value in (config or {}).items():
        resolved[key] = value
    resolved["patterns"] = [*DEFAULT_PATTERNS, *list(resolved.get("patterns") or [])]
    resolved["table_patterns"] = [*DEFAULT_TABLE_PATTERNS, *list(resolved.get("table_patterns") or [])]
    return resolved


def run_embedded_artifact_cleanup(
    doc_path: str | Path,
    *,
    document_name: str | None = None,
    config: dict[str, Any] | None = None,
    dry_run: bool = True,
) -> list[EmbeddedArtifactRecord]:
    cleanup_config = resolve_cleanup_config(config)
    doc = load_docx(doc_path)
    candidates = _scan_candidates(doc, document_name or Path(doc_path).name, cleanup_config)

    should_apply = cleanup_config.get("action") == "remove" and not dry_run
    changed = False
    if should_apply:
        changed = _apply_candidates(candidates)
        if changed:
            save_docx(doc, doc_path)
    return [candidate.record for candidate in candidates]


def scan_embedded_artifacts(
    doc_path: str | Path,
    *,
    document_name: str | None = None,
    config: dict[str, Any] | None = None,
) -> list[EmbeddedArtifactRecord]:
    return run_embedded_artifact_cleanup(
        doc_path,
        document_name=document_name,
        config={**(config or {}), "action": "preview"},
        dry_run=True,
    )


def excluded_locations(records: list[EmbeddedArtifactRecord]) -> set[str]:
    return {record.location for record in records if record.action == "exclude"}


def _scan_candidates(
    doc: DocxDocument,
    document_name: str,
    config: dict[str, Any],
) -> list[_ArtifactCandidate]:
    paragraphs = _paragraph_infos(doc)
    paragraph_counts = Counter(info.normalized for info in paragraphs if info.normalized)
    tables = _table_infos(doc)
    table_counts = Counter(info.signature for info in tables if info.signature)
    real_header_footer_texts = _real_header_footer_texts(doc)

    candidates: list[_ArtifactCandidate] = []
    for info in paragraphs:
        candidate = _paragraph_candidate(
            info,
            paragraphs=paragraphs,
            document_name=document_name,
            occurrence_count=paragraph_counts[info.normalized],
            real_header_footer_texts=real_header_footer_texts,
            config=config,
        )
        if candidate:
            candidates.append(candidate)

    for info in tables:
        candidate = _table_candidate(
            info,
            document_name=document_name,
            occurrence_count=table_counts[info.signature],
            real_header_footer_texts=real_header_footer_texts,
            config=config,
        )
        if candidate:
            candidates.append(candidate)

    return candidates


def _paragraph_infos(doc: DocxDocument) -> list[_ParagraphInfo]:
    sections_by_location = {element.location: element.section_path for element in collect_elements(doc)}
    infos: list[_ParagraphInfo] = []
    for index, paragraph in enumerate(doc.paragraphs):
        text = _compact(paragraph.text)
        if not text:
            continue
        location = f"body:{index}"
        infos.append(
            _ParagraphInfo(
                paragraph=paragraph,
                index=index,
                location=location,
                text=text,
                normalized=normalize_text(text),
                style_name=(paragraph.style.name if paragraph.style is not None else ""),
                alignment=_alignment_name(paragraph),
                section_path=sections_by_location.get(location, ()),
            )
        )
    return infos


def _table_infos(doc: DocxDocument) -> list[_TableInfo]:
    infos: list[_TableInfo] = []
    for index, table in enumerate(doc.tables):
        text = _table_text(table)
        normalized = normalize_text(text)
        infos.append(
            _TableInfo(
                table=table,
                index=index,
                location=f"table[{index}]",
                text=text,
                normalized=normalized,
                signature=_table_signature(table),
            )
        )
    return infos


def _paragraph_candidate(
    info: _ParagraphInfo,
    *,
    paragraphs: list[_ParagraphInfo],
    document_name: str,
    occurrence_count: int,
    real_header_footer_texts: set[str],
    config: dict[str, Any],
) -> _ArtifactCandidate | None:
    reasons: list[str] = []
    confidence = 0.0
    normalized = info.normalized
    neighbor_norms = _neighbor_norms(paragraphs, info.index, radius=3)
    metadata_neighbor_count = sum(1 for text in neighbor_norms if _is_metadata_text(text, config))

    confidence = max(confidence, _score_metadata_text(normalized, reasons, config))

    if re.fullmatch(r"\d+(?:\.\d+){0,2}", normalized) and any(
        neighbor == "version" or neighbor.startswith("version ") for neighbor in neighbor_norms
    ):
        confidence = max(confidence, 0.86)
        reasons.append("numeric_version_neighbor")

    if len(normalized) <= 90 and metadata_neighbor_count >= 2:
        confidence = max(confidence, 0.74)
        reasons.append("near_footer_metadata_cluster")

    if occurrence_count >= 3 and _is_metadata_text(normalized, config):
        confidence += 0.14
        reasons.append("recurring_metadata_text")

    real_match = normalized in real_header_footer_texts
    if real_match:
        confidence += 0.2
        reasons.append("matches_real_header_or_footer")

    if info.alignment in {"CENTER", "RIGHT"} and len(normalized) <= 90 and _is_metadata_text(normalized, config):
        confidence += 0.04
        reasons.append("short_aligned_metadata")

    if ACTION_VERB_RE.search(info.text) and not _is_metadata_text(normalized, config):
        confidence -= 0.35
        reasons.append("protected_action_text")

    confidence = max(0.0, min(1.0, confidence))
    if confidence < float(config.get("min_confidence_review", 0.45)):
        return None

    has_structural_break = _has_structural_break(info.paragraph)
    action = _action_for_paragraph(info, confidence, has_structural_break, config, reasons)
    record = EmbeddedArtifactRecord(
        document_name=document_name,
        location=info.location,
        block_type="body",
        action=action,
        confidence=round(confidence, 3),
        reasons=_unique(reasons),
        text=info.text,
        normalized_text=normalized,
        style_name=info.style_name,
        alignment=info.alignment,
        section_path=info.section_path,
        occurrence_count=occurrence_count,
        recurring_group_id=_hash_id(normalized) if occurrence_count > 1 else "",
        real_header_footer_match=real_match,
        context_before=_context_before(paragraphs, info.index),
        context_after=_context_after(paragraphs, info.index),
        candidate_id=_candidate_id(document_name, info.location, info.text),
    )
    return _ArtifactCandidate(
        record=record,
        paragraph=info.paragraph,
        body_index=info.index,
        has_structural_break=has_structural_break,
    )


def _table_candidate(
    info: _TableInfo,
    *,
    document_name: str,
    occurrence_count: int,
    real_header_footer_texts: set[str],
    config: dict[str, Any],
) -> _ArtifactCandidate | None:
    if not info.normalized:
        return None
    reasons: list[str] = []
    label_hits = _table_label_hits(info.normalized, config)
    confidence = 0.0
    if label_hits >= 3:
        confidence = max(confidence, 0.55)
        reasons.append("metadata_table_labels")
    if occurrence_count >= 2 and label_hits >= 2:
        confidence += 0.27
        reasons.append("recurring_metadata_table")
    real_match = any(part in real_header_footer_texts for part in _table_cell_norms(info.table))
    if real_match:
        confidence += 0.12
        reasons.append("contains_real_header_or_footer_text")

    confidence = max(0.0, min(1.0, confidence))
    if confidence < float(config.get("min_confidence_review", 0.45)):
        return None

    action = _action_for_table(info, confidence, config, reasons)
    record = EmbeddedArtifactRecord(
        document_name=document_name,
        location=info.location,
        block_type="table",
        action=action,
        confidence=round(confidence, 3),
        reasons=_unique(reasons),
        text=info.text,
        normalized_text=info.normalized,
        table_index=info.index,
        occurrence_count=occurrence_count,
        recurring_group_id=_hash_id(info.signature) if occurrence_count > 1 else "",
        real_header_footer_match=real_match,
        candidate_id=_candidate_id(document_name, info.location, info.text),
    )
    return _ArtifactCandidate(record=record, table=info.table, table_index=info.index)


def _action_for_paragraph(
    info: _ParagraphInfo,
    confidence: float,
    has_structural_break: bool,
    config: dict[str, Any],
    reasons: list[str],
) -> str:
    action = str(config.get("action", "preview")).lower()
    min_remove = float(config.get("min_confidence_remove", 0.82))
    max_chars = int(config.get("max_removable_chars", 160))
    if "protected_action_text" in reasons or len(info.text) > max_chars:
        return "protected"
    if confidence < min_remove:
        return "review"
    if action == "exclude":
        return "exclude"
    cleanup_action = "clear_text" if has_structural_break else "remove"
    if action == "remove":
        return cleanup_action
    return f"would_{cleanup_action}"


def _action_for_table(
    info: _TableInfo,
    confidence: float,
    config: dict[str, Any],
    reasons: list[str],
) -> str:
    action = str(config.get("action", "preview")).lower()
    min_remove = float(config.get("min_confidence_remove", 0.82))
    remove_tables = bool(config.get("remove_table_artifacts", False))
    protect_front_matter = bool(config.get("protect_front_matter_tables", True))
    front_matter_count = int(config.get("front_matter_table_count", 1))

    if protect_front_matter and info.index < front_matter_count:
        reasons.append("protected_front_matter_table")
        return "protected"
    if not remove_tables:
        reasons.append("table_removal_disabled")
        return "review"
    if confidence < min_remove:
        return "review"
    if action == "remove":
        return "remove_table"
    if action == "exclude":
        return "exclude"
    return "would_remove_table"


def _apply_candidates(candidates: list[_ArtifactCandidate]) -> bool:
    changed = False
    for candidate in candidates:
        if candidate.record.action != "clear_text" or candidate.paragraph is None:
            continue
        _clear_paragraph_text_preserving_breaks(candidate.paragraph)
        candidate.record.applied = True
        changed = True

    for candidate in sorted(
        (item for item in candidates if item.record.action == "remove" and item.paragraph is not None),
        key=lambda item: item.body_index if item.body_index is not None else -1,
        reverse=True,
    ):
        _remove_paragraph(candidate.paragraph)
        candidate.record.applied = True
        changed = True

    for candidate in sorted(
        (item for item in candidates if item.record.action == "remove_table" and item.table is not None),
        key=lambda item: item.table_index if item.table_index is not None else -1,
        reverse=True,
    ):
        _remove_table(candidate.table)
        candidate.record.applied = True
        changed = True
    return changed


def _score_metadata_text(normalized: str, reasons: list[str], config: dict[str, Any]) -> float:
    score = 0.0
    if normalized == "version" or re.fullmatch(r"version\s+\d+(?:\.\d+){0,2}", normalized):
        score = max(score, 0.72)
        reasons.append("footer_label_version")
    if "este es un documento controlado" in normalized:
        score = max(score, 0.88)
        reasons.append("controlled_document_footer")
    if "fecha de autorizacion" in normalized:
        score = max(score, 0.86)
        reasons.append("authorization_date_footer")
    if "proxima revision" in normalized:
        score = max(score, 0.86)
        reasons.append("next_revision_footer")
    if re.fullmatch(r"pagina\s+\d+(?:\s+de\s+\d+)?", normalized):
        score = max(score, 0.88)
        reasons.append("page_number_footer")
    if re.fullmatch(rf"{MONTH_RE}\s+\d{{4}}", normalized):
        score = max(score, 0.5)
        reasons.append("date_only_footer")
    if re.fullmatch(r"[-_=]{4,}", normalized):
        score = max(score, 0.54)
        reasons.append("separator_line")
    if any(re.search(pattern, normalized, re.IGNORECASE) for pattern in config.get("patterns") or []):
        score = max(score, 0.74)
        reasons.append("configured_footer_pattern")
    return score


def _is_metadata_text(normalized: str, config: dict[str, Any]) -> bool:
    if not normalized:
        return False
    if _score_metadata_text(normalized, [], config) >= 0.5:
        return True
    return bool(re.fullmatch(rf"{MONTH_RE}\s+\d{{4}}", normalized))


def _table_label_hits(normalized: str, config: dict[str, Any]) -> int:
    patterns = config.get("table_patterns") or []
    return sum(1 for pattern in patterns if re.search(pattern, normalized, re.IGNORECASE))


def _real_header_footer_texts(doc: DocxDocument) -> set[str]:
    texts: set[str] = set()
    for section in doc.sections:
        for story in (section.header, section.footer):
            for paragraph in story.paragraphs:
                normalized = normalize_text(paragraph.text)
                if normalized:
                    texts.add(normalized)
            for table in story.tables:
                texts.update(text for text in _table_cell_norms(table) if text)
    return texts


def _table_text(table: Table) -> str:
    parts: list[str] = []
    for row in table.rows:
        for cell in row.cells:
            text = _compact(cell.text)
            if text:
                parts.append(text)
    return " | ".join(parts)


def _table_signature(table: Table) -> str:
    return "|".join(_table_cell_norms(table))


def _table_cell_norms(table: Table) -> list[str]:
    values: list[str] = []
    for row in table.rows:
        for cell in row.cells:
            normalized = normalize_text(cell.text)
            if normalized:
                values.append(normalized)
    return values


def _neighbor_norms(paragraphs: list[_ParagraphInfo], body_index: int, *, radius: int) -> list[str]:
    values: list[str] = []
    for info in paragraphs:
        if info.index == body_index:
            continue
        if abs(info.index - body_index) <= radius:
            values.append(info.normalized)
    return values


def _context_before(paragraphs: list[_ParagraphInfo], body_index: int) -> str:
    previous = [info.text for info in paragraphs if info.index < body_index]
    return previous[-1] if previous else ""


def _context_after(paragraphs: list[_ParagraphInfo], body_index: int) -> str:
    for info in paragraphs:
        if info.index > body_index:
            return info.text
    return ""


def _has_structural_break(paragraph: Paragraph) -> bool:
    paragraph_properties = paragraph._p.pPr
    if paragraph_properties is not None and paragraph_properties.find(qn("w:sectPr")) is not None:
        return True
    for run in paragraph.runs:
        for child in run._r.iterchildren():
            if child.tag in {qn("w:br"), qn("w:lastRenderedPageBreak")}:
                return True
    return False


def _clear_paragraph_text_preserving_breaks(paragraph: Paragraph) -> None:
    for run in paragraph.runs:
        for child in list(run._r.iterchildren()):
            if child.tag in TEXT_CHILD_TAGS:
                run._r.remove(child)


def _remove_paragraph(paragraph: Paragraph) -> None:
    element = paragraph._element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def _remove_table(table: Table) -> None:
    element = table._element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def _alignment_name(paragraph: Paragraph) -> str:
    if paragraph.alignment is None:
        return ""
    return getattr(paragraph.alignment, "name", str(paragraph.alignment))


def _compact(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def _unique(values: list[str]) -> list[str]:
    return list(dict.fromkeys(values))


def _hash_id(value: str) -> str:
    return hashlib.sha1(value.encode("utf-8")).hexdigest()[:12]


def _candidate_id(document_name: str, location: str, text: str) -> str:
    return _hash_id(f"{document_name}\0{location}\0{text}")