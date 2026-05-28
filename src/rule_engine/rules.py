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
    match_text: str = ""
    selected_target: str = ""
    selector_reason: str = ""
    context_excerpt: str = ""


@dataclass(frozen=True)
class DetectionConfig:
    regex: str


@dataclass(frozen=True)
class ConditionalTarget:
    target_phrase: str
    reason: str
    patterns: tuple[str, ...]
    context_window_chars: int


@dataclass(frozen=True)
class ReplacementConfig:
    mode: str
    target_phrase: str
    connector: str
    format_template: str
    conditional_targets: tuple[ConditionalTarget, ...]


@dataclass(frozen=True)
class GuardConfig:
    skip_if_target_exists_in_paragraph: bool
    skip_if_alternative_exists: bool
    alternative_pattern: str | None
    review_only_patterns: tuple[str, ...]
    review_only_full_text_patterns: tuple[str, ...]
    exempt_context_patterns: tuple[str, ...]
    exempt_window_chars: int
    required_context_patterns: tuple[str, ...]
    required_context_window_chars: int
    target_after_match_window_chars: int
    descriptor_order_repair_patterns: tuple[str, ...]
    excluded_section_patterns: tuple[str, ...]


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
        conditional_targets = tuple(
            ConditionalTarget(
                target_phrase=str(target_cfg["target_phrase"]),
                reason=str(target_cfg.get("reason", target_cfg["target_phrase"])),
                patterns=tuple(str(pattern) for pattern in target_cfg.get("patterns", [])),
                context_window_chars=int(target_cfg.get("context_window_chars", 160)),
            )
            for target_cfg in replacement_cfg.get("conditional_targets", [])
        )
        self.replacement = ReplacementConfig(
            mode=str(replacement_cfg.get("mode", "expansion")),
            target_phrase=str(replacement_cfg["target_phrase"]),
            connector=str(replacement_cfg.get("connector", "o")).strip(),
            format_template=str(replacement_cfg.get("format", "{matched} {connector} {target}")),
            conditional_targets=conditional_targets,
        )
        self.guards = GuardConfig(
            skip_if_target_exists_in_paragraph=bool(
                guards_cfg.get("skip_if_target_exists_in_paragraph", False)
            ),
            skip_if_alternative_exists=bool(guards_cfg.get("skip_if_alternative_exists", True)),
            alternative_pattern=guards_cfg.get("alternative_pattern"),
            review_only_patterns=tuple(str(pattern) for pattern in guards_cfg.get("review_only_patterns", [])),
            review_only_full_text_patterns=tuple(
                str(pattern) for pattern in guards_cfg.get("review_only_full_text_patterns", [])
            ),
            exempt_context_patterns=tuple(
                str(pattern) for pattern in guards_cfg.get("exempt_context_patterns", [])
            ),
            exempt_window_chars=int(guards_cfg.get("exempt_window_chars", 40)),
            required_context_patterns=tuple(
                str(pattern) for pattern in guards_cfg.get("required_context_patterns", [])
            ),
            required_context_window_chars=int(guards_cfg.get("required_context_window_chars", 180)),
            target_after_match_window_chars=int(guards_cfg.get("target_after_match_window_chars", 180)),
            descriptor_order_repair_patterns=tuple(
                str(pattern) for pattern in guards_cfg.get("descriptor_order_repair_patterns", [])
            ),
            excluded_section_patterns=tuple(
                str(pattern) for pattern in guards_cfg.get("excluded_section_patterns", [])
            ),
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
        self.review_only_full_text_patterns = [
            re.compile(pattern) for pattern in self.guards.review_only_full_text_patterns
        ]
        self.exempt_context_patterns = [
            re.compile(pattern) for pattern in self.guards.exempt_context_patterns
        ]
        self.required_context_patterns = [
            re.compile(pattern) for pattern in self.guards.required_context_patterns
        ]
        self.excluded_section_patterns = [
            re.compile(pattern) for pattern in self.guards.excluded_section_patterns
        ]
        self.conditional_target_patterns = [
            (target, [re.compile(pattern) for pattern in target.patterns])
            for target in self.replacement.conditional_targets
        ]
        self.all_target_phrases = tuple(
            dict.fromkeys(
                [self.replacement.target_phrase]
                + [target.target_phrase for target in self.replacement.conditional_targets]
            )
        )
        target_pattern = "|".join(re.escape(target) for target in self.all_target_phrases)
        self._target_after_match_pattern = re.compile(
            r"^\s*(?:" + re.escape(self.replacement.connector) + r"|/|y/o)\s+(?:el\s+)?"
            + r"(?:" + target_pattern + r")",
            re.IGNORECASE,
        )
        descriptor_pattern = "|".join(
            f"(?:{pattern})" for pattern in self.guards.descriptor_order_repair_patterns
        )
        self._target_before_descriptor_pattern = (
            re.compile(
                r"(?P<operator>\boperador(?:\s*/\s*a|\s*\([aA]\)|a|es|as)?)"
                r"\s+"
                + re.escape(self.replacement.connector)
                + r"\s+(?P<target>"
                + target_pattern
                + r")\s+(?P<descriptor>"
                + descriptor_pattern
                + r")",
                re.IGNORECASE,
            )
            if descriptor_pattern
            else None
        )

    @property
    def target_phrase(self) -> str:
        return self.replacement.target_phrase

    def has_candidate(self, text: str) -> bool:
        return bool(self.pattern.search(text))

    def _is_review_only(self, match_text: str, full_text: str) -> bool:
        if any(pattern.search(match_text) or pattern.search(full_text) for pattern in self.review_only_patterns):
            return True
        return any(pattern.search(full_text) for pattern in self.review_only_full_text_patterns)

    def _is_exempt_context(self, full_text: str, match: re.Match[str]) -> bool:
        if not self.exempt_context_patterns:
            return False
        end = min(len(full_text), match.end() + self.guards.exempt_window_chars)
        phrase = full_text[match.start() : end]
        return any(pattern.search(phrase) for pattern in self.exempt_context_patterns)

    def _context_excerpt(self, full_text: str, match: re.Match[str], window_chars: int) -> str:
        start = max(0, match.start() - window_chars)
        end = min(len(full_text), match.end() + window_chars)
        return re.sub(r"\s+", " ", full_text[start:end]).strip()

    def _is_in_excluded_section(self, section_path: tuple[str, ...], in_excluded_section: bool) -> bool:
        if in_excluded_section:
            return True
        if not self.excluded_section_patterns:
            return False
        section_text = " > ".join(section_path)
        return any(pattern.search(section_text) for pattern in self.excluded_section_patterns)

    def _missing_required_context(self, full_text: str, match: re.Match[str]) -> bool:
        if not self.required_context_patterns:
            return False
        context = self._context_excerpt(full_text, match, self.guards.required_context_window_chars)
        return not any(pattern.search(context) for pattern in self.required_context_patterns)

    def _target_exists_near_match(self, full_text: str, match: re.Match[str]) -> bool:
        match_norm = normalize_text(match.group(0))
        if any(normalize_text(target) in match_norm for target in self.all_target_phrases):
            return True

        after = full_text[match.end() : match.end() + self.guards.target_after_match_window_chars]
        after_same_sentence = re.split(r"[\r\n.;:]", after, maxsplit=1)[0]
        after_norm = normalize_text(after_same_sentence)
        return any(normalize_text(target) in after_norm for target in self.all_target_phrases)

    def _select_target(self, full_text: str, match: re.Match[str]) -> tuple[str, str, str]:
        for target, patterns in self.conditional_target_patterns:
            context = self._context_excerpt(full_text, match, target.context_window_chars)
            if any(pattern.search(context) for pattern in patterns):
                return target.target_phrase, target.reason, context
        context = self._context_excerpt(full_text, match, self.guards.required_context_window_chars)
        return self.replacement.target_phrase, "default_target", context

    def _repair_target_before_descriptor(self, text: str) -> tuple[str, int, list[str], list[str]]:
        if self._target_before_descriptor_pattern is None:
            return text, 0, [], []
        repaired_matches: list[str] = []
        repaired_targets: list[str] = []

        def repair(match: re.Match[str]) -> str:
            repaired_matches.append(match.group(0))
            repaired_targets.append(match.group("target"))
            return (
                f"{match.group('operator')} {match.group('descriptor')} "
                f"{self.replacement.connector} {match.group('target')}"
            )

        repaired = self._target_before_descriptor_pattern.sub(repair, text)
        return repaired, len(repaired_matches), repaired_matches, repaired_targets

    def build_replacement(self, match_text: str, target_phrase: str | None = None) -> str:
        if self.replacement.mode != "expansion":
            raise ValueError(f"Unsupported replacement mode for {self.id}: {self.replacement.mode}")
        return self.replacement.format_template.format(
            matched=match_text,
            connector=self.replacement.connector,
            target=target_phrase or self.replacement.target_phrase,
        )

    def apply(
        self,
        text: str,
        *,
        section_path: tuple[str, ...] = (),
        in_excluded_section: bool = False,
    ) -> RuleDecision:
        if not self.enabled:
            return RuleDecision(self.id, False, text, text, "rule disabled")

        if not self.pattern.search(text):
            return RuleDecision(self.id, False, text, text, "no candidate")

        matches = list(self.pattern.finditer(text))
        if self._is_in_excluded_section(section_path, in_excluded_section):
            return RuleDecision(
                self.id,
                False,
                text,
                text,
                "candidate skipped in excluded section",
                candidates=len(matches),
                skipped=len(matches),
                skip_type="skip_section",
                match_text=" | ".join(match.group(0) for match in matches[:5]),
                context_excerpt=" > ".join(section_path),
            )

        original_text = text
        repaired_text, repaired_count, repaired_match_texts, repaired_targets = (
            self._repair_target_before_descriptor(text)
        )
        text = repaired_text

        text_norm = normalize_text(text)
        if self.guards.skip_if_target_exists_in_paragraph and any(
            normalize_text(target) in text_norm for target in self.all_target_phrases
        ):
            return RuleDecision(
                self.id,
                False,
                original_text,
                text,
                "target phrase already present in paragraph",
                candidates=len(matches),
                already_expanded=1,
            )

        candidates = 0
        already_expanded = 0
        skipped = 0
        changed = 0
        skip_types: list[str] = []
        changed_match_texts: list[str] = []
        changed_targets: list[str] = []
        selector_reasons: list[str] = []
        changed_contexts: list[str] = []
        skipped_match_texts: list[str] = []
        skipped_contexts: list[str] = []

        def note_skip(skip_type: str, match: re.Match[str]) -> None:
            skip_types.append(skip_type)
            skipped_match_texts.append(match.group(0))
            skipped_contexts.append(self._context_excerpt(text, match, self.guards.required_context_window_chars))

        def replace(match: re.Match[str]) -> str:
            nonlocal candidates, already_expanded, skipped, changed
            candidates += 1
            match_text = match.group(0)
            after = text[match.end() : match.end() + 120]

            if self._target_exists_near_match(text, match):
                already_expanded += 1
                return match_text

            if self._target_after_match_pattern.search(after):
                already_expanded += 1
                return match_text

            if self.guards.skip_if_alternative_exists and self.existing_alternative_pattern:
                if self.existing_alternative_pattern.search(after):
                    skipped += 1
                    note_skip("skip_existing_alternative", match)
                    return match_text

            if self._is_exempt_context(text, match):
                skipped += 1
                note_skip("skip_exempt_context", match)
                return match_text

            if self._is_review_only(match_text, text):
                skipped += 1
                note_skip("skip_review_only", match)
                return match_text

            if self._missing_required_context(text, match):
                skipped += 1
                note_skip("skip_no_action_context", match)
                return match_text

            changed += 1
            selected_target, selector_reason, context_excerpt = self._select_target(text, match)
            changed_match_texts.append(match_text)
            changed_targets.append(selected_target)
            selector_reasons.append(selector_reason)
            changed_contexts.append(context_excerpt)
            return self.build_replacement(match_text, selected_target)

        modified = self.pattern.sub(replace, text)
        if changed == 0 and repaired_count == 0:
            if already_expanded:
                reason = "all candidates already expanded"
            elif skipped:
                reason = "candidate skipped by guard"
            else:
                reason = "no replacement produced"
            return RuleDecision(
                self.id,
                False,
                original_text,
                modified,
                reason,
                candidates=candidates,
                already_expanded=already_expanded,
                skipped=skipped,
                skip_type="|".join(sorted(set(skip_types))) if skip_types else "guard",
                match_text=" | ".join(skipped_match_texts[:5]),
                context_excerpt=" | ".join(skipped_contexts[:3]),
            )

        reason_parts: list[str] = []
        if repaired_count:
            reason_parts.append(f"repaired {repaired_count} descriptor order issue(s)")
        if changed:
            reason_parts.append(f"expanded {changed} match(es)")
        unique_targets = list(dict.fromkeys(changed_targets))
        for target in repaired_targets:
            if target not in unique_targets:
                unique_targets.append(target)
        unique_reasons = list(dict.fromkeys(selector_reasons))
        return RuleDecision(
            self.id,
            True,
            original_text,
            modified,
            "; ".join(reason_parts),
            candidates=candidates,
            already_expanded=already_expanded,
            skipped=skipped,
            match_text=" | ".join((repaired_match_texts + changed_match_texts)[:5]),
            selected_target=" | ".join(unique_targets),
            selector_reason=" | ".join(unique_reasons),
            context_excerpt=" | ".join(changed_contexts[:3]),
        )


def load_rules(config: dict[str, Any]) -> list[ReplacementRule]:
    return [ReplacementRule(rule_config) for rule_config in config.get("rules", [])]