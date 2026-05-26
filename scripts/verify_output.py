from __future__ import annotations

import sys
from pathlib import Path

from docx import Document


TARGET = "personal designado por minera Spence"


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


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python scripts/verify_output.py path/to/file.docx", file=sys.stderr)
        return 2
    path = Path(sys.argv[1])
    if not path.exists():
        print(f"File not found: {path}", file=sys.stderr)
        return 1
    texts = list(iter_text(path))
    target_count = sum(text.count(TARGET) for text in texts)
    duplicated = [text for text in texts if text.count(TARGET) > 1]
    print(f"target_count={target_count}")
    print(f"paragraphs_with_duplicate_target={len(duplicated)}")
    if duplicated:
        for text in duplicated[:5]:
            print(text)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())