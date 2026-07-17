from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path
from typing import Any


DEFAULT_CONFIG: dict[str, Any] = {
    "engine_preference": "auto",
    "last_input_path": "",
    "default_output_dir": "",
    "error_policy": "stop_on_error",
    "log_verbosity": "normal",
    "log_file": "batch_edit.log",
    "saved_operation_presets": [],
}


class ConfigManager:
    def __init__(self, config_path: str | Path) -> None:
        self.config_path = Path(config_path)

    def load(self) -> dict[str, Any]:
        if not self.config_path.is_file():
            return deepcopy(DEFAULT_CONFIG)
        try:
            data = json.loads(self.config_path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return deepcopy(DEFAULT_CONFIG)
        config = deepcopy(DEFAULT_CONFIG)
        if isinstance(data, dict):
            config.update(data)
        return config

    def save(self, config: dict[str, Any]) -> None:
        payload = deepcopy(DEFAULT_CONFIG)
        payload.update(config)
        self.config_path.write_text(
            json.dumps(payload, indent=2, sort_keys=True),
            encoding="utf-8",
        )
