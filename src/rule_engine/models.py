from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import datetime
from typing import Any


@dataclass
class DocxElement:
    text: str
    normalized: str
    location: str
    paragraph_obj: Any
    is_heading: bool = False


@dataclass
class ChangeRecord:
    document_name: str
    pass_index: int
    location: str
    original_text: str
    modified_text: str
    rule_id: str
    rule_category: str
    source: str
    reason: str
    changed_at: str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class SkipRecord:
    document_name: str
    pass_index: int
    location: str
    text: str
    rule_id: str
    reason: str
    skip_type: str = "guard"
    source: str = "rule"

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class PassSummary:
    pass_index: int
    candidates: int = 0
    changed: int = 0
    skipped: int = 0
    already_expanded: int = 0
    llm_attempted: int = 0
    llm_changed: int = 0