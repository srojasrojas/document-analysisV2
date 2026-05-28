from __future__ import annotations

import sys
import argparse
import re
from pathlib import Path

from docx import Document
import yaml

from rule_engine.docx_io import collect_elements, normalize_text


PROJECT_ROOT = Path(__file__).resolve().parents[1]
ORDER_DESCRIPTORS_RE = re.compile(
    r"(?:de\s+(?:el|la|los|las)\s+|de\s+|del\s+)?"
    r"(?:retro\s*-?\s*excavadoras?|mini\s*-?\s*cargador(?:a|es)?|"
    r"cargador(?:a|es)?(?:\s+frontal)?|camion(?:es)?\s+tolva|"
    r"rotopalas?|equipos?)\b",
    re.IGNORECASE,
)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def iter_text(path: Path):
    doc = Document(str(path))
    for element in collect_elements(doc):
        yield element.text


def _target_phrases(config_path: Path) -> list[str]:
    config = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
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


def _has_suspicious_duplicate(text: str, target: str) -> bool:
    target_re = re.escape(target)
    adjacent_pattern = re.compile(
        rf"{target_re}\s*(?:(?:,|;|/)?\s*(?:o|y|y/o)\s+)?{target_re}",
        re.IGNORECASE,
    )
    if adjacent_pattern.search(text):
        return True
    return bool(re.search(r"\bo\s+o\s+", text, re.IGNORECASE))


def _has_operator_descriptor_order_issue(text: str, targets: list[str]) -> bool:
    normalized = normalize_text(text)
    for target in targets:
        target_norm = normalize_text(target)
        if "personal" not in target_norm:
            continue
        pattern = re.compile(
            r"\boperador(?:a|es|as)?\s+o\s+"
            + re.escape(target_norm)
            + r"\s+"
            + ORDER_DESCRIPTORS_RE.pattern,
            re.IGNORECASE,
        )
        if pattern.search(normalized):
            return True
    return False


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
    targets = _target_phrases(args.config)
    if not targets:
        print(f"No enabled rule targets found in {args.config}", file=sys.stderr)
        return 1

    texts = list(iter_text(path))
    failed = False
    order_issues = [text for text in texts if _has_operator_descriptor_order_issue(text, targets)]
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
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())