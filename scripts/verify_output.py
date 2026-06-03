from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from docx import Document
import yaml

from rule_engine.docx_io import collect_elements, normalize_text


PROJECT_ROOT = Path(__file__).resolve().parents[1]
OPERATOR_ROLE_RE = re.compile(
    r"\boperador(?:\s*/\s*a|\s*\([aA]\)|a|es|as)?(?:\s+a\s+cargo)?"
    r"(?:\s+(?!o\b)[A-Za-zÁÉÍÓÚÑáéíóúñ&/().-]+){0,8}",
    re.IGNORECASE,
)
SPLIT_COMPOUND_DESCRIPTOR_RE = re.compile(
    r"\boperador(?:\s*/\s*a|\s*\([aA]\)|a|es|as)?"
    r"(?:\s+(?:de|del)\s+(?:(?:el|la|los|las)\s+)?)?cami[oó]n(?:es)?\s+o\s+"
    r"(?P<target>personal\s+(?:certificado\s+designado|designado)\s+por\s+minera\s+Spence|personal\s+calificado)"
    r"\s+(?:tolva|pluma)\b",
    re.IGNORECASE,
)
SUPERVISOR_LEGACY_RE = re.compile(
    r"\b(?:supervisor(?:a|es|as)?|jefe\s+de\s+[aá]rea|due[ñn]o\s+de\s+[aá]rea)"
    r"(?:\s+(?!o\b)[A-Za-zÁÉÍÓÚÑáéíóúñ&/().-]+){0,8}\s+o\s+experto\s+t[eé]cnico\b",
    re.IGNORECASE,
)
OPERATOR_TABLE_RESIDUAL_RE = re.compile(
    r"\boperador(?:\s*/\s*a|\s*\([aA]\)|a|es|as)?(?:\s+a\s+cargo)?"
    r"(?:\s+(?!o\b)[A-Za-zÁÉÍÓÚÑáéíóúñ0-9&/().-]+){0,10}\s+o\s+"
    r"(?P<target>personal\s+(?:certificado\s+designado|designado)\s+por\s+minera\s+Spence|personal\s+calificado)"
    r"\s+(?P<residual>Spence\b|(?:de\s+)?(?:EW|MLDC|MDC|SX|TF)\b(?:\s+Spence)?|"
    r"de\s+c[aá]todos\b|(?:de\s+)?otras\s+[aá]reas\b|(?:de\s+la\s+)?m[aá]quina\s+despegadora\b|"
    r"(?:de\s+)?embarque\b|(?:de\s+(?:el|la|los|las)\s+|de\s+|del\s+)?puentes?\s+gr[uú]as?\b|"
    r"encargad[oa]\b|designad[oa]s?\b)",
    re.IGNORECASE,
)
SPENCE_DUPLICATE_RE = re.compile(r"\bSpence\s+Spence\b", re.IGNORECASE)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def iter_text(path: Path):
    doc = Document(str(path))
    for element in collect_elements(doc):
        yield element.text


def _load_config(config_path: Path) -> dict:
    return yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}


def _target_phrases(config: dict) -> list[str]:
    targets: list[str] = []
    for rule in config.get("rules", []):
        if not rule.get("enabled", True):
            continue
        target = rule.get("replacement", {}).get("target_phrase")
        if target and target not in targets:
            targets.append(str(target))
        for conditional in rule.get("replacement", {}).get("conditional_targets", []):
            conditional_target = conditional.get("target_phrase")
            if conditional_target and conditional_target not in targets:
                targets.append(str(conditional_target))
    return targets


def _operator_rule(config: dict) -> dict:
    for rule in config.get("rules", []):
        if rule.get("id") == "operador_to_spence":
            return rule
    return {}


def _operator_descriptor_patterns(config: dict) -> list[re.Pattern[str]]:
    rule = _operator_rule(config)
    patterns = rule.get("guards", {}).get("descriptor_order_repair_patterns", [])
    return [re.compile(str(pattern), re.IGNORECASE) for pattern in patterns]


def _certified_operator_patterns(config: dict) -> tuple[str, list[re.Pattern[str]]]:
    rule = _operator_rule(config)
    for target in rule.get("replacement", {}).get("conditional_targets", []):
        target_phrase = str(target.get("target_phrase", ""))
        if "certificado designado" not in normalize_text(target_phrase):
            continue
        patterns = [
            re.compile(str(pattern), re.IGNORECASE)
            for pattern in target.get("match_patterns", [])
        ]
        return target_phrase, patterns
    return "", []


def _has_suspicious_duplicate(text: str, target: str) -> bool:
    target_re = re.escape(target)
    adjacent_pattern = re.compile(
        rf"{target_re}\s*(?:(?:,|;|/)?\s*(?:o|y|y/o)\s+)?{target_re}",
        re.IGNORECASE,
    )
    if adjacent_pattern.search(text):
        return True
    return bool(re.search(r"\bo\s+o\s+", text, re.IGNORECASE))


def _has_operator_descriptor_order_issue(
    text: str, targets: list[str], descriptor_patterns: list[re.Pattern[str]]
) -> bool:
    for target in targets:
        if "personal" not in normalize_text(target):
            continue
        if SPLIT_COMPOUND_DESCRIPTOR_RE.search(text):
            return True
        for descriptor_pattern in descriptor_patterns:
            pattern = re.compile(
                r"\boperador(?:\s*/\s*a|\s*\([aA]\)|a|es|as)?\s+o\s+"
                + re.escape(target)
                + r"\s+(?:"
                + descriptor_pattern.pattern
                + r")",
                re.IGNORECASE,
            )
            if pattern.search(text):
                return True
    return False


def _has_operator_certification_target_issue(
    text: str, certified_target: str, certified_patterns: list[re.Pattern[str]]
) -> bool:
    if not certified_target or not certified_patterns:
        return False
    default_target = "personal designado por Minera Spence"
    expanded_pattern = re.compile(
        rf"(?P<role>{OPERATOR_ROLE_RE.pattern})\s+o\s+{re.escape(default_target)}",
        re.IGNORECASE,
    )
    for match in expanded_pattern.finditer(text):
        role_text = match.group("role")
        if any(pattern.search(role_text) for pattern in certified_patterns):
            return True
    return False


def _has_supervisor_legacy_issue(text: str) -> bool:
    return bool(SUPERVISOR_LEGACY_RE.search(text))


def _has_operator_table_residual_issue(text: str) -> bool:
    return bool(OPERATOR_TABLE_RESIDUAL_RE.search(text) or SPENCE_DUPLICATE_RE.search(text))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Verify duplicate configured target phrases in a DOCX.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--config", type=Path, default=PROJECT_ROOT / "config.yaml")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    path = args.docx
    if not path.exists():
        print(f"File not found: {path}", file=sys.stderr)
        return 1
    config = _load_config(args.config)
    targets = _target_phrases(config)
    if not targets:
        print(f"No enabled rule targets found in {args.config}", file=sys.stderr)
        return 1
    descriptor_patterns = _operator_descriptor_patterns(config)
    certified_target, certified_patterns = _certified_operator_patterns(config)

    texts = list(iter_text(path))
    failed = False
    order_issues = [
        text for text in texts if _has_operator_descriptor_order_issue(text, targets, descriptor_patterns)
    ]
    certification_issues = [
        text
        for text in texts
        if _has_operator_certification_target_issue(text, certified_target, certified_patterns)
    ]
    table_residual_issues = [text for text in texts if _has_operator_table_residual_issue(text)]
    supervisor_legacy_issues = [text for text in texts if _has_supervisor_legacy_issue(text)]
    for target in targets:
        target_count = sum(text.count(target) for text in texts)
        duplicated = [text for text in texts if text.count(target) > 1]
        suspicious = [text for text in duplicated if _has_suspicious_duplicate(text, target)]
        print(f"target={target}")
        print(f"target_count={target_count}")
        print(f"paragraphs_with_duplicate_target={len(duplicated)}")
        print(f"suspicious_duplicate_target={len(suspicious)}")
        if suspicious:
            failed = True
            for text in suspicious[:5]:
                print(text)
    print(f"operator_descriptor_order_issues={len(order_issues)}")
    if order_issues:
        failed = True
        for text in order_issues[:5]:
            print(text)
    print(f"operator_certification_target_issues={len(certification_issues)}")
    if certification_issues:
        failed = True
        for text in certification_issues[:5]:
            print(text)
    print(f"operator_table_residue_issues={len(table_residual_issues)}")
    if table_residual_issues:
        failed = True
        for text in table_residual_issues[:5]:
            print(text)
    print(f"supervisor_legacy_target_issues={len(supervisor_legacy_issues)}")
    if supervisor_legacy_issues:
        failed = True
        for text in supervisor_legacy_issues[:5]:
            print(text)
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())