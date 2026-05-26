from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openai import OpenAI

from .docx_io import normalize_text


@dataclass
class LlmResult:
    changed: bool
    modified_text: str
    reason: str


class LlmRefiner:
    def __init__(self, config: dict[str, Any], project_root: Path) -> None:
        self.config = config
        self.project_root = project_root
        llm_cfg = config.get("llm_refine", {})
        self.max_tokens = int(llm_cfg.get("max_tokens", 700))
        self.min_overlap = float(llm_cfg.get("min_original_overlap_ratio", 0.65))
        self.max_length_ratio = float(llm_cfg.get("max_length_ratio", 1.8))
        self.candidate_pattern = re.compile(str(llm_cfg.get("candidate_regex", r"(?i)operador")))
        self.gate_prompt = self._read_prompt(llm_cfg.get("gate_prompt", "prompts/operator_gate.md"))
        self.constructor_prompt = self._read_prompt(
            llm_cfg.get("constructor_prompt", "prompts/operator_constructor.md")
        )
        self.client, self.model, self.temperature = self._build_client(config)

    def _read_prompt(self, relative_path: str) -> str:
        prompt_path = self.project_root / relative_path
        if not prompt_path.exists():
            raise FileNotFoundError(f"Prompt not found: {prompt_path}")
        return prompt_path.read_text(encoding="utf-8")

    @staticmethod
    def _build_client(config: dict[str, Any]) -> tuple[OpenAI | None, str, float]:
        models_cfg = config.get("models", {})
        active_name = models_cfg.get("active", "openai_api")
        profile = models_cfg.get("profiles", {}).get(active_name)
        if not profile:
            raise ValueError(f"Active model profile not found: {active_name}")

        provider = str(profile.get("provider", "openai"))
        model = os.getenv(str(profile.get("model_env", ""))) or str(profile.get("default_model", ""))
        temperature = float(profile.get("temperature", 0.0))

        if provider == "openai":
            api_key = os.getenv(str(profile.get("api_key_env", "OPENAI_API_KEY")))
            if not api_key:
                return None, model, temperature
            return OpenAI(api_key=api_key), model, temperature

        if provider == "openai_compatible":
            base_url = os.getenv(str(profile.get("base_url_env", ""))) or str(
                profile.get("default_base_url", "")
            )
            api_key = os.getenv(str(profile.get("api_key_env", ""))) or str(
                profile.get("default_api_key", "ollama")
            )
            return OpenAI(api_key=api_key, base_url=base_url), model, temperature

        raise ValueError(f"Unsupported model provider: {provider}")

    def can_run(self) -> bool:
        return self.client is not None and bool(self.model)

    def has_candidate(self, text: str) -> bool:
        return bool(self.candidate_pattern.search(text))

    def _json_completion(self, system_prompt: str, user_prompt: str) -> dict[str, Any]:
        if not self.client:
            raise RuntimeError("LLM client is not configured. Set OPENAI_API_KEY or use simple-only mode.")
        response = self.client.chat.completions.create(
            model=self.model,
            temperature=self.temperature,
            max_tokens=self.max_tokens,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        )
        content = response.choices[0].message.content or "{}"
        return json.loads(content)

    def _validate(self, original: str, modified: str, target_phrase: str) -> str | None:
        if normalize_text(original) == normalize_text(modified):
            return "llm returned unchanged text"
        if len(modified) > max(40, int(len(original) * self.max_length_ratio)):
            return "llm response is too long"
        if modified.count(". ") > original.count(". "):
            return "llm added a new sentence"
        if normalize_text(target_phrase) not in normalize_text(modified):
            return "target phrase missing from llm response"

        original_tokens = set(normalize_text(original).split())
        modified_tokens = set(normalize_text(modified).split())
        if original_tokens:
            overlap_ratio = len(original_tokens & modified_tokens) / len(original_tokens)
            if overlap_ratio < self.min_overlap:
                return "llm changed too much of the paragraph"
        return None

    def refine(self, text: str, target_phrase: str) -> LlmResult:
        gate_payload = self._json_completion(
            self.gate_prompt,
            "TEXTO DEL PARRAFO:\n" + text,
        )
        if not bool(gate_payload.get("needs_change")):
            return LlmResult(False, text, str(gate_payload.get("reason", "gate rejected")))

        constructor_payload = self._json_completion(
            self.constructor_prompt,
            "TEXTO ORIGINAL DEL PARRAFO:\n" + text,
        )
        modified = str(constructor_payload.get("modified_text", "")).strip()
        if not modified:
            return LlmResult(False, text, "constructor returned empty text")

        validation_error = self._validate(text, modified, target_phrase)
        if validation_error:
            return LlmResult(False, text, validation_error)
        return LlmResult(True, modified, str(gate_payload.get("reason", "llm refined")))