from __future__ import annotations

from pathlib import Path
from typing import Any
import os

import yaml


def _deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    merged = dict(base)
    for key, value in override.items():
        existing = merged.get(key)
        if isinstance(existing, dict) and isinstance(value, dict):
            merged[key] = _deep_merge(existing, value)
        else:
            merged[key] = value
    return merged


def _local_override_path(config_path: Path) -> Path:
    return config_path.with_name(f"{config_path.stem}.local{config_path.suffix}")


def load_config(path: str | Path) -> dict[str, Any]:
    config_path = Path(path)
    if not config_path.exists():
        raise FileNotFoundError(f"Config file not found: {config_path}")
    with config_path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle) or {}
    if not isinstance(data, dict):
        raise ValueError(f"Config root must be a mapping: {config_path}")

    local_override_path = _local_override_path(config_path)
    if local_override_path.exists():
        with local_override_path.open("r", encoding="utf-8") as handle:
            local_data = yaml.safe_load(handle) or {}
        if not isinstance(local_data, dict):
            raise ValueError(f"Config root must be a mapping: {local_override_path}")
        data = _deep_merge(data, local_data)

    return data


def load_env_file(project_root: Path, filename: str = ".env") -> None:
    env_path = project_root / filename
    if not env_path.exists():
        return
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key:
            os.environ.setdefault(key, value)


def resolve_path(project_root: Path, value: str | Path) -> Path:
    path = Path(value)
    return path if path.is_absolute() else project_root / path


def project_root_from_config(config_path: str | Path) -> Path:
    return Path(config_path).resolve().parent