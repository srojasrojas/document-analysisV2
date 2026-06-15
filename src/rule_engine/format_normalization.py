"""Etapa cero: normalizacion de formato para DOCX convertidos desde PDF.

Aplica dos correcciones antes de la limpieza de encabezados/pies y de las
reglas de reemplazo:

- Aplanado de tablas anidadas: sustituye tablas dentro de celdas (mas alla de
  `max_table_depth`) por sus parrafos en orden de lectura.
- Homologacion de fuente: unifica la familia tipografica de todos los runs a la
  fuente dominante del documento (o a una fija si se configura `target_font`),
  conservando las fuentes de simbolos/vinetas.

Puede ejecutarse como etapa integrada del pipeline o de forma standalone con
`scripts/normalize_format.py`.
"""

from __future__ import annotations

import copy
from collections import Counter
from pathlib import Path
from typing import Any

from docx.document import Document as DocxDocument
from docx.oxml.ns import qn

from .docx_io import load_docx, save_docx
from .models import FormatNormalizationRecord


DEFAULT_NORMALIZATION_CONFIG: dict[str, Any] = {
    "enabled": False,
    "apply_before_rules": True,
    "normalize_fonts": True,
    "target_font": None,
    "flatten_nested_tables": True,
    "max_table_depth": 1,
}

# Fuentes de simbolos/vinetas que nunca deben homologarse a la fuente objetivo.
SYMBOL_FONTS = {
    "symbol",
    "wingdings",
    "wingdings 2",
    "wingdings 3",
    "webdings",
    "mt extra",
    "marlett",
    "zapfdingbats",
    "cambria math",
}

_FONT_ATTRS = (qn("w:ascii"), qn("w:hAnsi"), qn("w:cs"), qn("w:eastAsia"))


def resolve_normalization_config(config: dict[str, Any] | None = None) -> dict[str, Any]:
    resolved = dict(DEFAULT_NORMALIZATION_CONFIG)
    for key, value in (config or {}).items():
        resolved[key] = value
    return resolved


def run_format_normalization(
    docx_path: str | Path,
    *,
    document_name: str,
    config: dict[str, Any] | None = None,
    dry_run: bool = False,
) -> list[FormatNormalizationRecord]:
    """Aplica la normalizacion de formato sobre el documento en sitio."""
    cfg = resolve_normalization_config(config)
    doc = load_docx(docx_path)
    counts: Counter[str] = Counter()

    if cfg.get("flatten_nested_tables", True):
        counts["nested_tables_flattened"] += _flatten_nested_tables(
            doc, max_depth=int(cfg.get("max_table_depth", 1))
        )
    if cfg.get("normalize_fonts", True):
        normalized_runs, target_font = _normalize_fonts(doc, target_font=cfg.get("target_font"))
        counts["font_overrides_normalized"] += normalized_runs
        if target_font:
            counts[f"font_target::{target_font}"] = 0

    changed = any(value for key, value in counts.items() if not key.startswith("font_target::"))
    applied = bool(changed and not dry_run)
    if applied:
        save_docx(doc, docx_path)

    records = [
        FormatNormalizationRecord(
            document_name=document_name,
            action=action if applied or not changed else f"would_{action}",
            count=count,
            applied=applied,
        )
        for action, count in sorted(counts.items())
        if count or action.startswith("font_target::")
    ]
    return records


# ---------------------------------------------------------------------------
# Fuentes


def _dominant_font(doc: DocxDocument) -> str | None:
    counter: Counter[str] = Counter()
    for run in doc.element.body.iter(qn("w:r")):
        text = "".join(t.text or "" for t in run.findall(qn("w:t")))
        if not text.strip():
            continue
        rpr = run.find(qn("w:rPr"))
        if rpr is None:
            continue
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            continue
        name = rfonts.get(qn("w:ascii")) or rfonts.get(qn("w:hAnsi"))
        if name and name.lower() not in SYMBOL_FONTS:
            counter[name] += len(text)
    if not counter:
        return None
    return counter.most_common(1)[0][0]


def _normalize_fonts(doc: DocxDocument, *, target_font: str | None) -> tuple[int, str | None]:
    """Homologa la fuente de todos los runs de texto a una sola familia.

    Si no se configura `target_font` se usa la fuente dominante del documento
    (ponderada por cantidad de texto). Las fuentes de simbolos se conservan
    para no romper vinetas.
    """
    target = target_font or _dominant_font(doc)
    if not target:
        return 0, None
    changed = 0
    for rfonts in doc.element.body.iter(qn("w:rFonts")):
        current = rfonts.get(qn("w:ascii")) or rfonts.get(qn("w:hAnsi"))
        if current and current.lower() in SYMBOL_FONTS:
            continue
        updated = False
        for attr in _FONT_ATTRS:
            existing = rfonts.get(attr)
            if existing is not None and existing.lower() in SYMBOL_FONTS:
                continue
            if existing != target:
                rfonts.set(attr, target)
                updated = True
        # Limpia referencias a temas que ganarian sobre los atributos directos.
        for theme_attr in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
            if rfonts.get(qn(theme_attr)) is not None:
                del rfonts.attrib[qn(theme_attr)]
                updated = True
        if updated:
            changed += 1
    changed += _normalize_default_font(doc, target)
    return changed, target


def _normalize_default_font(doc: DocxDocument, target: str) -> int:
    styles = getattr(doc, "styles", None)
    element = getattr(styles, "element", None)
    if element is None:
        return 0
    changed = 0
    for rpr_default in element.iter(qn("w:rPrDefault")):
        for rfonts in rpr_default.iter(qn("w:rFonts")):
            updated = False
            for attr in _FONT_ATTRS:
                if rfonts.get(attr) != target:
                    rfonts.set(attr, target)
                    updated = True
            for theme_attr in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
                if rfonts.get(qn(theme_attr)) is not None:
                    del rfonts.attrib[qn(theme_attr)]
                    updated = True
            if updated:
                changed += 1
    return changed


# ---------------------------------------------------------------------------
# Tablas anidadas


def _table_nesting_level(table) -> int:
    level = 0
    parent = table.getparent()
    while parent is not None:
        if parent.tag == qn("w:tbl"):
            level += 1
        parent = parent.getparent()
    return level


def _paragraph_is_disposable(paragraph) -> bool:
    """Vacio de texto y sin contenido estructural (imagenes, saltos, marcas)."""
    text = "".join(t.text or "" for t in paragraph.iter(qn("w:t")))
    if text.strip():
        return False
    blocking_tags = (
        qn("w:drawing"),
        qn("w:pict"),
        qn("w:object"),
        qn("w:fldSimple"),
        qn("w:fldChar"),
        qn("w:bookmarkStart"),
        qn("w:commentRangeStart"),
        qn("w:sectPr"),
        qn("w:numPr"),
    )
    for tag in blocking_tags:
        if paragraph.find(f".//{tag}") is not None:
            return False
    for br in paragraph.iter(qn("w:br")):
        if br.get(qn("w:type")) in {"page", "column"}:
            return False
    return True


def _flatten_nested_tables(doc: DocxDocument, *, max_depth: int) -> int:
    """Sustituye tablas anidadas mas alla de `max_depth` por sus parrafos.

    El contenido textual se conserva en orden de lectura dentro de la celda
    contenedora; solo desaparece la estructura de tabla interna.
    """
    flattened = 0
    while True:
        candidates = [
            table
            for table in doc.element.body.iter(qn("w:tbl"))
            if _table_nesting_level(table) >= max_depth
        ]
        if not candidates:
            break
        # Procesa la mas profunda primero para no invalidar referencias.
        target = max(candidates, key=_table_nesting_level)
        parent = target.getparent()
        index = list(parent).index(target)
        replacement: list[Any] = []
        for row in target.findall(qn("w:tr")):
            for cell in row.findall(qn("w:tc")):
                for child in cell:
                    if child.tag == qn("w:p"):
                        text = "".join(t.text or "" for t in child.iter(qn("w:t")))
                        if text.strip() or not _paragraph_is_disposable(child):
                            replacement.append(copy.deepcopy(child))
        parent.remove(target)
        for offset, paragraph in enumerate(replacement):
            parent.insert(index + offset, paragraph)
        if parent.tag == qn("w:tc") and parent.find(qn("w:p")) is None:
            parent.append(parent.makeelement(qn("w:p"), {}))
        flattened += 1
    return flattened
