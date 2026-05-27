from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openai import OpenAI

from .docx_io import normalize_text
from .rules import ReplacementRule


@dataclass
class LlmResult:
    changed: bool
    modified_text: str
    reason: str


class LlmRefiner:
    def __init__(self, config: dict[str, Any], project_root: Path) -> None:
        self.config = config
        self.project_root = project_root
        self.client, self.model, self.temperature, self.use_response_format = self._build_client(config)
        llm_cfg = config.get("llm_refine", {})
        self.repair_retries = int(llm_cfg.get("repair_retries", 1))
        self._prompt_cache: dict[str, str] = {}

    def _read_prompt(self, relative_path: str) -> str:
        prompt_path = self.project_root / relative_path
        if not prompt_path.exists():
            raise FileNotFoundError(f"Prompt not found: {prompt_path}")
        cache_key = str(prompt_path)
        if cache_key not in self._prompt_cache:
            self._prompt_cache[cache_key] = prompt_path.read_text(encoding="utf-8")
        return self._prompt_cache[cache_key]

    @staticmethod
    def _resolve_model_profile(config: dict[str, Any]) -> tuple[str, dict[str, Any]]:
        models_cfg = config.get("models", {})
        llm_cfg = config.get("llm_refine", {})
        model_key = llm_cfg.get("model") or models_cfg.get("active") or "deepseek-r1-distill-14b-local"

        # Backward-compatible shape from the first implementation.
        if "profiles" in models_cfg:
            profile = models_cfg.get("profiles", {}).get(model_key)
            if not profile:
                raise ValueError(f"Active model profile not found: {model_key}")
            return str(model_key), profile

        profile = models_cfg.get(model_key)
        if not profile:
            available = ", ".join(k for k in models_cfg if isinstance(models_cfg.get(k), dict))
            raise ValueError(f"Model '{model_key}' not found in config models. Available: {available}")
        return str(model_key), profile

    @classmethod
    def _build_client(cls, config: dict[str, Any]) -> tuple[OpenAI | None, str, float, bool]:
        llm_cfg = config.get("llm_refine", {})
        model_key, profile = cls._resolve_model_profile(config)

        provider = str(profile.get("provider", "openai"))
        model = (
            os.getenv(str(profile.get("deployment_env", "")))
            or os.getenv(str(profile.get("model_env", "")))
            or str(profile.get("deployment", profile.get("default_model", model_key)))
        )
        temperature = float(llm_cfg.get("temperature", profile.get("temperature", 0.0)))
        timeout = float(llm_cfg.get("timeout_seconds", profile.get("timeout_seconds", 180)))

        if provider == "openai":
            api_key = os.getenv(str(profile.get("api_key_env", "OPENAI_API_KEY")))
            if not api_key:
                return None, model, temperature, True
            return OpenAI(api_key=api_key, timeout=timeout), model, temperature, True

        if provider in {"openai_compatible", "ollama"}:
            base_url = (
                os.getenv(str(profile.get("base_url_env", "")))
                or str(profile.get("base_url", profile.get("default_base_url", "")))
            )
            api_key = (
                os.getenv(str(profile.get("api_key_env", "")))
                or str(profile.get("api_key", profile.get("default_api_key", "ollama")))
            )
            if not base_url:
                raise ValueError(f"base_url is required for model profile '{model_key}'")
            # Ollama/DeepSeek local often emits reasoning or markdown fences; parse JSON manually.
            return OpenAI(api_key=api_key, base_url=base_url, timeout=timeout), model, temperature, False

        raise ValueError(f"Unsupported model provider: {provider}")

    def can_run(self) -> bool:
        return self.client is not None and bool(self.model)

    @staticmethod
    def _strip_reasoning(text: str) -> str:
        return re.sub(r"<think>.*?</think>", "", text, flags=re.IGNORECASE | re.DOTALL).strip()

    @classmethod
    def _extract_json_payload(cls, text: str) -> dict[str, Any]:
        cleaned = cls._strip_reasoning(text)
        fenced = re.findall(r"```(?:json)?\s*(.*?)```", cleaned, flags=re.IGNORECASE | re.DOTALL)
        candidates = [*fenced, cleaned]
        decoder = json.JSONDecoder()
        last_payload: object | None = None
        for candidate in candidates:
            payload = candidate.strip()
            if not payload:
                continue
            try:
                parsed = json.loads(payload)
                if isinstance(parsed, dict):
                    return parsed
            except json.JSONDecodeError:
                pass
            for index, char in enumerate(payload):
                if char not in "{[":
                    continue
                try:
                    parsed, _ = decoder.raw_decode(payload[index:])
                    last_payload = parsed
                    if isinstance(parsed, dict):
                        return parsed
                except json.JSONDecodeError:
                    continue
        if isinstance(last_payload, dict):
            return last_payload
        raise ValueError(f"Could not extract JSON object from model response: {text[:500]!r}")

    @staticmethod
    def _json_system_prompt(system_prompt: str) -> str:
        return (
            f"{system_prompt.strip()}\n\n"
            "Debes responder SOLO con un objeto JSON valido, sin markdown, sin cercos ``` "
            "y sin texto adicional. Si generas razonamiento interno, no lo incluyas."
        )

    @staticmethod
    def _repair_prompt(raw_response: str) -> str:
        return (
            "Tu respuesta anterior no se pudo parsear como JSON. "
            "Devuelve SOLO un objeto JSON valido, sin markdown ni explicaciones.\n\n"
            f"Respuesta anterior:\n{raw_response.strip()}"
        )

    def _json_completion(self, system_prompt: str, user_prompt: str, *, max_tokens: int) -> dict[str, Any]:
        if not self.client:
            raise RuntimeError("LLM client is not configured. Set OPENAI_API_KEY or use simple-only mode.")
        system_content = self._json_system_prompt(system_prompt)
        messages = [
            {"role": "system", "content": system_content},
            {"role": "user", "content": user_prompt},
        ]
        last_content = ""
        for attempt_index in range(self.repair_retries + 1):
            kwargs: dict[str, Any] = {
                "model": self.model,
                "temperature": self.temperature,
                "max_tokens": max_tokens,
                "messages": messages,
            }
            if self.use_response_format:
                kwargs["response_format"] = {"type": "json_object"}
            response = self.client.chat.completions.create(**kwargs)
            last_content = response.choices[0].message.content or "{}"
            try:
                return self._extract_json_payload(last_content)
            except Exception:
                if attempt_index >= self.repair_retries:
                    raise
                messages = [
                    *messages,
                    {"role": "assistant", "content": last_content},
                    {"role": "user", "content": self._repair_prompt(last_content)},
                ]
        raise ValueError(f"Could not parse JSON response: {last_content[:500]!r}")

    def _validate(self, original: str, modified: str, rule: ReplacementRule) -> str | None:
        if normalize_text(original) == normalize_text(modified):
            return "llm returned unchanged text"
        if len(modified) > max(40, int(len(original) * rule.llm.max_length_ratio)):
            return "llm response is too long"
        if modified.count(". ") > original.count(". "):
            return "llm added a new sentence"
        if rule.llm.validate_target_in_output and normalize_text(rule.target_phrase) not in normalize_text(modified):
            return "target phrase missing from llm response"

        original_tokens = set(normalize_text(original).split())
        modified_tokens = set(normalize_text(modified).split())
        if original_tokens:
            overlap_ratio = len(original_tokens & modified_tokens) / len(original_tokens)
            if overlap_ratio < rule.llm.min_original_overlap_ratio:
                return "llm changed too much of the paragraph"
        return None

    def refine_for_rule(self, text: str, rule: ReplacementRule) -> LlmResult:
        if not rule.llm.gate_prompt or not rule.llm.constructor_prompt:
            return LlmResult(False, text, "llm prompts not configured for rule")
        gate_prompt = self._read_prompt(rule.llm.gate_prompt)
        constructor_prompt = self._read_prompt(rule.llm.constructor_prompt)
        gate_payload = self._json_completion(
            gate_prompt,
            f"REGLA: {rule.id}\nFRASE OBJETIVO: {rule.target_phrase}\n\nTEXTO DEL PARRAFO:\n{text}",
            max_tokens=rule.llm.max_tokens,
        )
        if not bool(gate_payload.get("needs_change")):
            return LlmResult(False, text, str(gate_payload.get("reason", "gate rejected")))

        constructor_payload = self._json_completion(
            constructor_prompt,
            f"REGLA: {rule.id}\nFRASE OBJETIVO: {rule.target_phrase}\n\nTEXTO ORIGINAL DEL PARRAFO:\n{text}",
            max_tokens=rule.llm.max_tokens,
        )
        modified = str(constructor_payload.get("modified_text", "")).strip()
        if not modified:
            return LlmResult(False, text, "constructor returned empty text")

        validation_error = self._validate(text, modified, rule)
        if validation_error:
            return LlmResult(False, text, validation_error)
        return LlmResult(True, modified, str(gate_payload.get("reason", "llm refined")))