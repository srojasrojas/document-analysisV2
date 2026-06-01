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
    block_type: str = "body"
    section_path: tuple[str, ...] = field(default_factory=tuple)
    in_excluded_section: bool = False
    table_index: int | None = None
    row_index: int | None = None
    cell_index: int | None = None


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
    match_text: str = ""
    selected_target: str = ""
    selector_reason: str = ""
    context_excerpt: str = ""
    block_type: str = ""
    section_path: tuple[str, ...] = field(default_factory=tuple)
    candidate_id: str = ""
    qa_flags: list[str] = field(default_factory=list)
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
    match_text: str = ""
    selected_target: str = ""
    context_excerpt: str = ""
    block_type: str = ""
    section_path: tuple[str, ...] = field(default_factory=tuple)
    candidate_id: str = ""
    qa_flags: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class EmbeddedArtifactRecord:
    document_name: str
    location: str
    block_type: str
    action: str
    confidence: float
    reasons: list[str]
    text: str
    normalized_text: str
    style_name: str = ""
    alignment: str = ""
    section_path: tuple[str, ...] = field(default_factory=tuple)
    table_index: int | None = None
    row_index: int | None = None
    cell_index: int | None = None
    occurrence_count: int = 1
    recurring_group_id: str = ""
    real_header_footer_match: bool = False
    context_before: str = ""
    context_after: str = ""
    applied: bool = False
    candidate_id: str = ""
    detected_at: str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))

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