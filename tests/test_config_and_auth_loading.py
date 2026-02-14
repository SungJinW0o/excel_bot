import json
import shutil
import sys
import uuid
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from excel_bot.auth import load_users
from excel_bot.config import load_config


def _make_temp_dir() -> Path:
    base_dir = Path(__file__).resolve().parent / "tmp"
    base_dir.mkdir(exist_ok=True)
    temp_dir = base_dir / f"run_{uuid.uuid4().hex}"
    temp_dir.mkdir()
    return temp_dir


def test_load_config_accepts_utf8_bom():
    temp_dir = _make_temp_dir()
    try:
        config_path = temp_dir / "config.json"
        config_data = {
            "paths": {"input_dir": "input_data", "output_dir": "output_data"},
            "files": {
                "input_extension": ".xlsx",
                "cleaned_output": "cleaned_master.xlsx",
                "report_output": "summary_report.xlsx",
            },
            "columns": {
                "quantity": "Quantity",
                "unit_price": "UnitPrice",
                "status": "Status",
                "category": "Category",
                "region": "Region",
            },
            "filters": {
                "min_quantity": 1,
                "min_unit_price": 0.01,
                "exclude_status": ["Cancelled"],
                "include_status": ["Completed"],
            },
        }
        config_path.write_text(json.dumps(config_data), encoding="utf-8-sig")

        loaded = load_config(str(config_path))
        assert loaded["paths"]["input_dir"] == "input_data"
        assert loaded["files"]["report_output"] == "summary_report.xlsx"
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def test_load_users_accepts_utf8_bom():
    temp_dir = _make_temp_dir()
    try:
        users_path = temp_dir / "users.json"
        users_data = [
            {
                "id": "u1",
                "email": "analyst1@example.com",
                "role": "analyst",
                "status": "active",
            }
        ]
        users_path.write_text(json.dumps(users_data), encoding="utf-8-sig")

        loaded = load_users(str(users_path))
        assert "analyst1@example.com" in loaded
        assert loaded["analyst1@example.com"].role == "analyst"
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
