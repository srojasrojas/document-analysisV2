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


class ReplacementRule:
    def __init__(self, config: dict[str, Any]) -> None:
        self.id = str(config["id"])
        self.enabled = bool(config.get("enabled", True))
        self.description = str(config.get("description", self.id))
        self.target_phrase = str(config["target_phrase"])
        self.connector = str(config.get("connector", "o")).strip()
        self.pattern = re.compile(str(config["detection_regex"]))
        self.skip_when_target_present_in_paragraph = bool(
            config.get("skip_when_target_present_in_paragraph", False)
        )
        self.skip_when_existing_alternative_after_match = bool(
            config.get("skip_when_existing_alternative_after_match", True)
        )
        alternative_regex = config.get("existing_alternative_regex")
        self.existing_alternative_pattern = (
            re.compile(str(alternative_regex)) if alternative_regex else None
        )
        self.review_only_patterns = [
            re.compile(str(pattern)) for pattern in config.get("review_only_regexes", [])
        ]
        self._target_after_match_pattern = re.compile(
            r"^\s*" + re.escape(self.connector) + r"\s+" + re.escape(self.target_phrase),
            re.IGNORECASE,
        )

    @property
    def expansion_text(self) -> str:
        return f"{self.connector} {self.target_phrase}"

    def has_candidate(self, text: str) -> bool:
        return bool(self.pattern.search(text))

    def _is_review_only(self, match_text: str, full_text: str) -> bool:
        return any(pattern.search(match_text) or pattern.search(full_text) for pattern in self.review_only_patterns)

    def apply(self, text: str) -> RuleDecision:
        if not self.enabled:
            return RuleDecision(self.id, False, text, text, "rule disabled")

        if not self.pattern.search(text):
            return RuleDecision(self.id, False, text, text, "no candidate")

        target_norm = normalize_text(self.target_phrase)
        if self.skip_when_target_present_in_paragraph and target_norm in normalize_text(text):
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

            if self.skip_when_existing_alternative_after_match and self.existing_alternative_pattern:
                if self.existing_alternative_pattern.search(after):
                    skipped += 1
                    return match_text

            if self._is_review_only(match_text, text):
                skipped += 1
                return match_text

            changed += 1
            return f"{match_text} {self.expansion_text}"

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
            f"expanded {changed} operator mention(s)",
            candidates=candidates,
            already_expanded=already_expanded,
            skipped=skipped,
        )


def load_rules(config: dict[str, Any]) -> list[ReplacementRule]:
    return [ReplacementRule(rule_config) for rule_config in config.get("rules", [])]