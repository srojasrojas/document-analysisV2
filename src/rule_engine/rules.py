from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from .docx_io import normalize_text


@dataclass
class RuleDecision:
    rule_id: str
    changed: bool
    original_text: str
    modified_text: str
    reason: str
    candidates: int = 0
    already_expanded: int = 0
    skipped: int = 0
    skip_type: str = "guard"


@dataclass(frozen=True)
class DetectionConfig:
    regex: str


@dataclass(frozen=True)
class ReplacementConfig:
    mode: str
    target_phrase: str
    connector: str
    format_template: str


@dataclass(frozen=True)
class GuardConfig:
    skip_if_target_exists_in_paragraph: bool
    skip_if_alternative_exists: bool
    alternative_pattern: str | None
    review_only_patterns: tuple[str, ...]
    exempt_context_patterns: tuple[str, ...]
    exempt_window_chars: int


@dataclass(frozen=True)
class LlmRuleConfig:
    enabled: bool
    gate_prompt: str | None
    constructor_prompt: str | None
    validate_target_in_output: bool
    max_tokens: int
    min_original_overlap_ratio: float
    max_length_ratio: float


class ReplacementRule:
    def __init__(self, config: dict[str, Any]) -> None:
        self.id = str(config["id"])
        self.enabled = bool(config.get("enabled", True))
        self.category = str(config.get("category", "general"))
        self.description = str(config.get("description", self.id))

        detection_cfg = config.get("detection", {})
        replacement_cfg = config.get("replacement", {})
        guards_cfg = config.get("guards", {})
        llm_cfg = config.get("llm_refining", {})

        self.detection = DetectionConfig(regex=str(detection_cfg["regex"]))
        self.replacement = ReplacementConfig(
            mode=str(replacement_cfg.get("mode", "expansion")),
            target_phrase=str(replacement_cfg["target_phrase"]),
            connector=str(replacement_cfg.get("connector", "o")).strip(),
            format_template=str(replacement_cfg.get("format", "{matched} {connector} {target}")),
        )
        self.guards = GuardConfig(
            skip_if_target_exists_in_paragraph=bool(
                guards_cfg.get("skip_if_target_exists_in_paragraph", False)
            ),
            skip_if_alternative_exists=bool(guards_cfg.get("skip_if_alternative_exists", True)),
            alternative_pattern=guards_cfg.get("alternative_pattern"),
            review_only_patterns=tuple(str(pattern) for pattern in guards_cfg.get("review_only_patterns", [])),
            exempt_context_patterns=tuple(
                str(pattern) for pattern in guards_cfg.get("exempt_context_patterns", [])
            ),
            exempt_window_chars=int(guards_cfg.get("exempt_window_chars", 40)),
        )
        self.llm = LlmRuleConfig(
            enabled=bool(llm_cfg.get("enabled", False)),
            gate_prompt=llm_cfg.get("gate_prompt"),
            constructor_prompt=llm_cfg.get("constructor_prompt"),
            validate_target_in_output=bool(llm_cfg.get("validate_target_in_output", True)),
            max_tokens=int(llm_cfg.get("max_tokens", 700)),
            min_original_overlap_ratio=float(llm_cfg.get("min_original_overlap_ratio", 0.65)),
            max_length_ratio=float(llm_cfg.get("max_length_ratio", 1.8)),
        )

        self.pattern = re.compile(self.detection.regex)
        alternative_regex = self.guards.alternative_pattern
        self.existing_alternative_pattern = (
            re.compile(str(alternative_regex)) if alternative_regex else None
        )
        self.review_only_patterns = [
            re.compile(pattern) for pattern in self.guards.review_only_patterns
        ]
        self.exempt_context_patterns = [
            re.compile(pattern) for pattern in self.guards.exempt_context_patterns
        ]
        self._target_after_match_pattern = re.compile(
            r"^\s*(?:" + re.escape(self.replacement.connector) + r"|/|y/o)\s+(?:el\s+)?"
            + re.escape(self.target_phrase),
            re.IGNORECASE,
        )

    @property
    def target_phrase(self) -> str:
        return self.replacement.target_phrase

    def has_candidate(self, text: str) -> bool:
        return bool(self.pattern.search(text))

    def _is_review_only(self, match_text: str, full_text: str) -> bool:
        return any(pattern.search(match_text) or pattern.search(full_text) for pattern in self.review_only_patterns)

    def _is_exempt_context(self, full_text: str, match: re.Match[str]) -> bool:
        if not self.exempt_context_patterns:
            return False
        end = min(len(full_text), match.end() + self.guards.exempt_window_chars)
        phrase = full_text[match.start() : end]
        return any(pattern.search(phrase) for pattern in self.exempt_context_patterns)

    def build_replacement(self, match_text: str) -> str:
        if self.replacement.mode != "expansion":
            raise ValueError(f"Unsupported replacement mode for {self.id}: {self.replacement.mode}")
        return self.replacement.format_template.format(
            matched=match_text,
            connector=self.replacement.connector,
            target=self.replacement.target_phrase,
        )

    def apply(self, text: str) -> RuleDecision:
        if not self.enabled:
            return RuleDecision(self.id, False, text, text, "rule disabled")

        if not self.pattern.search(text):
            return RuleDecision(self.id, False, text, text, "no candidate")

        target_norm = normalize_text(self.target_phrase)
        if self.guards.skip_if_target_exists_in_paragraph and target_norm in normalize_text(text):
            return RuleDecision(
                self.id,
                False,
                text,
                text,
                "target phrase already present in paragraph",
                candidates=len(list(self.pattern.finditer(text))),
                already_expanded=1,
            )

        candidates = 0
        already_expanded = 0
        skipped = 0
        changed = 0

        def replace(match: re.Match[str]) -> str:
            nonlocal candidates, already_expanded, skipped, changed
            candidates += 1
            match_text = match.group(0)
            after = text[match.end() : match.end() + 120]

            if self._target_after_match_pattern.search(after):
                already_expanded += 1
                return match_text

            if self.guards.skip_if_alternative_exists and self.existing_alternative_pattern:
                if self.existing_alternative_pattern.search(after):
                    skipped += 1
                    return match_text

            if self._is_exempt_context(text, match):
                skipped += 1
                return match_text

            if self._is_review_only(match_text, text):
                skipped += 1
                return match_text

            changed += 1
            return self.build_replacement(match_text)

        modified = self.pattern.sub(replace, text)
        if changed == 0:
            if already_expanded:
                reason = "all candidates already expanded"
            elif skipped:
                reason = "candidate skipped by guard"
            else:
                reason = "no replacement produced"
            return RuleDecision(
                self.id,
                False,
                text,
                modified,
                reason,
                candidates=candidates,
                already_expanded=already_expanded,
                skipped=skipped,
            )

        return RuleDecision(
            self.id,
            True,
            text,
            modified,
            f"expanded {changed} match(es)",
            candidates=candidates,
            already_expanded=already_expanded,
            skipped=skipped,
        )


def load_rules(config: dict[str, Any]) -> list[ReplacementRule]:
    return [ReplacementRule(rule_config) for rule_config in config.get("rules", [])]