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
    match_patterns: tuple[str, ...]
    context_window_chars: int


@dataclass(frozen=True)
class ReplacementConfig:
    mode: str
    target_phrase: str
    connector: str
    format_template: str
    conditional_targets: tuple[ConditionalTarget, ...]
    upgradeable_target_phrases: tuple[str, ...]
    legacy_target_patterns: tuple[str, ...]


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
                match_patterns=tuple(str(pattern) for pattern in target_cfg.get("match_patterns", [])),
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
            upgradeable_target_phrases=tuple(
                str(target) for target in replacement_cfg.get("upgradeable_target_phrases", [])
            ),
            legacy_target_patterns=tuple(
                str(pattern) for pattern in replacement_cfg.get("legacy_target_patterns", [])
            ),
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
            (
                target,
                [re.compile(pattern) for pattern in target.patterns],
                [re.compile(pattern) for pattern in target.match_patterns],
            )
            for target in self.replacement.conditional_targets
        ]
        self.all_target_phrases = tuple(
            dict.fromkeys(
                [self.replacement.target_phrase]
                + [target.target_phrase for target in self.replacement.conditional_targets]
                + list(self.replacement.upgradeable_target_phrases)
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
        operator_descriptor_stopwords = (
            r"(?:o|u|y|a|para|que|cuando|si|debe(?:n|r[aá]n?|r[aá])?|realiza(?:r|n)?|"
            r"revisa(?:r|n)?|verifica(?:r|n)?|detiene(?:r|n)?|opera(?:r|n)?|"
            r"inspecciona(?:r|n)?|coordina(?:r|n)?|informa(?:r|n)?|avisa(?:r|n)?|"
            r"solicita(?:r|n)?|autoriza(?:r|n)?|registra(?:r|n)?|bloquea(?:r|n)?|"
            r"desbloquea(?:r|n)?|energiza(?:r|n)?|desenergiza(?:r|n)?|comunica(?:r|n)?|"
            r"operador(?:a|es|as)?|supervisor(?:a|es|as)?|CAS|CIO)\b"
        )
        operator_pattern = (
            r"\boperador(?:\s*/\s*a|\s*\([aA]\)|a|es|as)?(?:\s+a\s+cargo)?"
            r"(?:\s+(?!"
            + operator_descriptor_stopwords
            + r")[A-Za-zÁÉÍÓÚÑáéíóúñ0-9&/().-]+){0,10}?"
        )
        self._target_before_descriptor_pattern = (
            re.compile(
                r"(?P<operator>" + operator_pattern + r")"
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
        target_repair_patterns = [
            re.escape(target) for target in self.replacement.upgradeable_target_phrases
        ] + list(self.replacement.legacy_target_patterns)
        self._target_repair_after_match_patterns = [
            re.compile(
                r"^\s*(?:"
                + re.escape(self.replacement.connector)
                + r"|/|y/o)\s+(?:el\s+)?(?P<target>"
                + pattern
                + r")",
                re.IGNORECASE,
            )
            for pattern in target_repair_patterns
        ]

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
        return self._context_excerpt_for_span(full_text, match.start(), match.end(), window_chars)

    def _context_excerpt_for_span(
        self, full_text: str, start_index: int, end_index: int, window_chars: int
    ) -> str:
        start = max(0, start_index - window_chars)
        end = min(len(full_text), end_index + window_chars)
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
        return self._select_target_for_text(full_text, match.group(0), match.start(), match.end())

    def _select_target_for_text(
        self, full_text: str, match_text: str, start_index: int, end_index: int
    ) -> tuple[str, str, str]:
        for target, context_patterns, match_patterns in self.conditional_target_patterns:
            if any(pattern.search(match_text) for pattern in match_patterns):
                context = self._context_excerpt_for_span(
                    full_text, start_index, end_index, target.context_window_chars
                )
                return target.target_phrase, target.reason, context
            context = self._context_excerpt_for_span(
                full_text, start_index, end_index, target.context_window_chars
            )
            if any(pattern.search(context) for pattern in context_patterns):
                return target.target_phrase, target.reason, context
        context = self._context_excerpt_for_span(
            full_text, start_index, end_index, self.guards.required_context_window_chars
        )
        return self.replacement.target_phrase, "default_target", context

    def _repair_targets_after_matches(
        self, text: str
    ) -> tuple[str, int, list[str], list[str], list[str], list[str]]:
        if not self._target_repair_after_match_patterns:
            return text, 0, [], [], [], []

        repairs: list[tuple[int, int, str, str, str, str, str]] = []
        for match in self.pattern.finditer(text):
            selected_target, selector_reason, context_excerpt = self._select_target(text, match)
            after = text[match.end() : match.end() + self.guards.target_after_match_window_chars]
            for target_pattern in self._target_repair_after_match_patterns:
                target_match = target_pattern.search(after)
                if target_match is None:
                    continue
                current_target = target_match.group("target")
                if normalize_text(current_target) == normalize_text(selected_target):
                    break
                start = match.end() + target_match.start("target")
                end = match.end() + target_match.end("target")
                repairs.append(
                    (
                        start,
                        end,
                        selected_target,
                        f"{match.group(0)} {self.replacement.connector} {current_target}",
                        selected_target,
                        selector_reason,
                        context_excerpt,
                    )
                )
                break

        if not repairs:
            return text, 0, [], [], [], []

        repaired = text
        applied: list[tuple[int, int, str, str, str, str, str]] = []
        last_start = len(text) + 1
        for repair in sorted(repairs, key=lambda item: item[0], reverse=True):
            start, end, selected_target, *_ = repair
            if end > last_start:
                continue
            repaired = repaired[:start] + selected_target + repaired[end:]
            applied.append(repair)
            last_start = start

        applied.reverse()
        return (
            repaired,
            len(applied),
            [item[3] for item in applied],
            [item[4] for item in applied],
            [item[5] for item in applied],
            [item[6] for item in applied],
        )

    def _repair_target_before_descriptor(
        self, text: str
    ) -> tuple[str, int, list[str], list[str], list[str], list[str]]:
        if self._target_before_descriptor_pattern is None:
            return text, 0, [], [], [], []
        repaired_matches: list[str] = []
        repaired_targets: list[str] = []
        repaired_reasons: list[str] = []
        repaired_contexts: list[str] = []

        def repair(match: re.Match[str]) -> str:
            role_text = f"{match.group('operator')} {match.group('descriptor')}"
            selected_target, selector_reason, context_excerpt = self._select_target_for_text(
                text, role_text, match.start(), match.end()
            )
            repaired_matches.append(match.group(0))
            repaired_targets.append(selected_target)
            repaired_reasons.append(selector_reason)
            repaired_contexts.append(context_excerpt)
            return (
                f"{match.group('operator')} {match.group('descriptor')} "
                f"{self.replacement.connector} {selected_target}"
            )

        repaired = self._target_before_descriptor_pattern.sub(repair, text)
        return (
            repaired,
            len(repaired_matches),
            repaired_matches,
            repaired_targets,
            repaired_reasons,
            repaired_contexts,
        )

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
        original_text = text
        (
            repaired_text,
            repaired_count,
            repaired_match_texts,
            repaired_targets,
            repaired_reasons,
            repaired_contexts,
        ) = (
            self._repair_target_before_descriptor(text)
        )
        text = repaired_text
        (
            text,
            target_repair_count,
            target_repair_match_texts,
            target_repair_targets,
            target_repair_reasons,
            target_repair_contexts,
        ) = self._repair_targets_after_matches(text)

        if self._is_in_excluded_section(section_path, in_excluded_section):
            if repaired_count or target_repair_count:
                reason_parts: list[str] = []
                if repaired_count:
                    reason_parts.append(f"repaired {repaired_count} descriptor order issue(s)")
                if target_repair_count:
                    reason_parts.append(f"updated {target_repair_count} target phrase(s)")
                unique_targets = list(dict.fromkeys(target_repair_targets))
                for target in repaired_targets:
                    if target not in unique_targets:
                        unique_targets.append(target)
                return RuleDecision(
                    self.id,
                    True,
                    original_text,
                    text,
                    "; ".join(reason_parts),
                    candidates=len(matches),
                    match_text=" | ".join((repaired_match_texts + target_repair_match_texts)[:5]),
                    selected_target=" | ".join(unique_targets),
                    selector_reason=" | ".join(dict.fromkeys(repaired_reasons + target_repair_reasons)),
                    context_excerpt=" | ".join((repaired_contexts + target_repair_contexts)[:3]),
                )
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

        text_norm = normalize_text(text)
        if target_repair_count == 0 and self.guards.skip_if_target_exists_in_paragraph and any(
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
        if target_repair_count and self.guards.skip_if_target_exists_in_paragraph:
            return RuleDecision(
                self.id,
                True,
                original_text,
                text,
                f"updated {target_repair_count} target phrase(s)",
                candidates=len(list(self.pattern.finditer(text))),
                already_expanded=1,
                match_text=" | ".join(target_repair_match_texts[:5]),
                selected_target=" | ".join(dict.fromkeys(target_repair_targets)),
                selector_reason=" | ".join(dict.fromkeys(target_repair_reasons)),
                context_excerpt=" | ".join(target_repair_contexts[:3]),
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
        (
            modified,
            post_repaired_count,
            post_repaired_match_texts,
            post_repaired_targets,
            post_repaired_reasons,
            post_repaired_contexts,
        ) = self._repair_target_before_descriptor(modified)
        if changed == 0 and repaired_count == 0 and target_repair_count == 0 and post_repaired_count == 0:
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
        if target_repair_count:
            reason_parts.append(f"updated {target_repair_count} target phrase(s)")
        if post_repaired_count:
            reason_parts.append(f"repaired {post_repaired_count} post-expansion descriptor issue(s)")
        if changed:
            reason_parts.append(f"expanded {changed} match(es)")
        unique_targets = list(dict.fromkeys(target_repair_targets + post_repaired_targets + changed_targets))
        for target in repaired_targets:
            if target not in unique_targets:
                unique_targets.append(target)
        unique_reasons = list(
            dict.fromkeys(target_repair_reasons + repaired_reasons + post_repaired_reasons + selector_reasons)
        )
        return RuleDecision(
            self.id,
            True,
            original_text,
            modified,
            "; ".join(reason_parts),
            candidates=candidates,
            already_expanded=already_expanded,
            skipped=skipped,
            match_text=" | ".join(
                (
                    repaired_match_texts
                    + target_repair_match_texts
                    + post_repaired_match_texts
                    + changed_match_texts
                )[:5]
            ),
            selected_target=" | ".join(unique_targets),
            selector_reason=" | ".join(unique_reasons),
            context_excerpt=" | ".join(
                (repaired_contexts + target_repair_contexts + post_repaired_contexts + changed_contexts)[:3]
            ),
        )


def load_rules(config: dict[str, Any]) -> list[ReplacementRule]:
    return [ReplacementRule(rule_config) for rule_config in config.get("rules", [])]