import json
import os
from typing import Any, Dict

class AppConfig:
    """Manages reading and writing application preferences to a JSON file."""

    def __init__(self, config_path: str):
        self.config_path = config_path

    def leer_config(self) -> Dict[str, Any]:
        """Reads configuration from the JSON file."""
        try:
            with open(self.config_path, encoding="utf8") as fh:
                return json.load(fh)
        except Exception:
            return {}

    def guardar_config(self, datos: Dict[str, Any]) -> None:
        """Saves configuration dictionary to the JSON file."""
        try:
            with open(self.config_path, "w", encoding="utf8") as fh:
                json.dump(datos, fh, ensure_ascii=False, indent=2)
        except Exception:
            pass
