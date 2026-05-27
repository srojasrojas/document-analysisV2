from __future__ import annotations

import sys
import argparse
from pathlib import Path

from docx import Document
import yaml


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def iter_text(path: Path):
    doc = Document(str(path))
    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            yield paragraph.text
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if paragraph.text.strip():
                        yield paragraph.text


def _target_phrases(config_path: Path) -> list[str]:
    config = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
    targets: list[str] = []
    for rule in config.get("rules", []):
        if not rule.get("enabled", True):
            continue
        target = rule.get("replacement", {}).get("target_phrase")
        if target and target not in targets:
            targets.append(str(target))
    return targets


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
    for target in targets:
        target_count = sum(text.count(target) for text in texts)
        duplicated = [text for text in texts if text.count(target) > 1]
        print(f"target={target}")
        print(f"target_count={target_count}")
        print(f"paragraphs_with_duplicate_target={len(duplicated)}")
        if duplicated:
            failed = True
            for text in duplicated[:5]:
                print(text)
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())