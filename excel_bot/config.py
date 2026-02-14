import json
import os
from importlib import resources
from typing import Optional


REQUIRED_CONFIG_KEYS = [
    ("paths", "input_dir"),
    ("paths", "output_dir"),
    ("files", "input_extension"),
    ("files", "cleaned_output"),
    ("files", "report_output"),
    ("filters", "exclude_status"),
    ("filters", "include_status"),
    ("filters", "min_quantity"),
    ("filters", "min_unit_price"),
    ("columns", "quantity"),
    ("columns", "unit_price"),
    ("columns", "status"),
    ("columns", "category"),
    ("columns", "region"),
]


def validate_config(cfg: dict) -> None:
    missing = []
    for section, key in REQUIRED_CONFIG_KEYS:
        if key not in cfg.get(section, {}):
            missing.append(f"{section}.{key}")
    if missing:
        raise KeyError(f"Missing config: {', '.join(missing)}")


def _load_config_from_path(path: str) -> dict:
    with open(path, "r", encoding="utf-8-sig") as f:
        cfg = json.load(f)
    validate_config(cfg)
    return cfg


def _load_config_from_package() -> dict:
    with resources.open_text(__package__, "config.json", encoding="utf-8") as f:
        cfg = json.load(f)
    validate_config(cfg)
    return cfg


def load_config(path: Optional[str] = None) -> dict:
    resolved_path = path or os.environ.get("EXCEL_BOT_CONFIG")
    if resolved_path:
        return _load_config_from_path(resolved_path)

    cwd_path = os.path.join(os.getcwd(), "config.json")
    if os.path.exists(cwd_path):
        return _load_config_from_path(cwd_path)

    return _load_config_from_package()
