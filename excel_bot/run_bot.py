#!/usr/bin/env python3
"""
Cross-platform Excel Bot Runner with CLI
Author: SungJinWoo
"""

import argparse
import json
import os
import platform
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, Optional


def _read_last_event(log_path: Path) -> Optional[Dict[str, Any]]:
    if not log_path.exists():
        return None
    try:
        with log_path.open("r", encoding="utf-8") as f:
            lines = [line.strip() for line in f if line.strip()]
        if not lines:
            return None
        last = lines[-1]
        try:
            payload = json.loads(last)
            if isinstance(payload, dict):
                return payload
            return {"raw": str(payload)}
        except json.JSONDecodeError:
            return {"raw": last}
    except Exception:
        return None


def _format_last_event(event: Dict[str, Any]) -> str:
    raw = event.get("raw")
    if raw:
        return str(raw)
    event_type = str(event.get("type", "UNKNOWN"))
    level = str(event.get("level", "INFO"))
    timestamp = str(event.get("timestamp", ""))
    return f"{event_type} ({level}) at {timestamp}"


def _resolve_python(root: Path) -> str:
    venv_dir = root / ".venv"
    if venv_dir.exists():
        if platform.system() == "Windows":
            candidate = venv_dir / "Scripts" / "python.exe"
        else:
            candidate = venv_dir / "bin" / "python"
        if candidate.exists():
            print(f"Using virtual environment: {candidate}")
            return str(candidate)
        print("Virtual environment detected but Python not found. Using system Python.")
    else:
        print("No .venv folder detected, using system Python.")
    return sys.executable


def main() -> int:
    parser = argparse.ArgumentParser(description="Run the Excel Bot (DRY_RUN optional).")
    parser.add_argument(
        "--dry-run",
        dest="dry_run",
        choices=["true", "false"],
        default="true",
        help="Enable DRY_RUN mode (default: true).",
    )
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="Do not open the report or log files after the run.",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Alias for --no-open (useful for servers or CI).",
    )
    args = parser.parse_args()

    dry_run = args.dry_run
    mode_label = "DRY RUN (safe test)" if dry_run == "true" else "LIVE RUN"
    print(f"Mode: {mode_label}")
    os.environ["DRY_RUN"] = dry_run
    open_files = not (args.no_open or args.headless)

    # 1. Set working directory
    root = Path.cwd()

    # 2. Create input/output folders
    input_dir = root / "input_data"
    output_dir = root / "output_data"
    input_dir.mkdir(exist_ok=True)
    output_dir.mkdir(exist_ok=True)

    # 3. Select Python (venv if present)
    python_exe = _resolve_python(root)

    # 4. Run the bot
    try:
        print("\nRunning Excel Bot pipeline...\n")
        subprocess.run([python_exe, "-m", "excel_bot.bot_main"], check=True)
    except subprocess.CalledProcessError as exc:
        print(f"Bot exited with code {exc.returncode}")
        return exc.returncode

    # 5. Open summary report
    report_path = output_dir / "summary_report.xlsx"
    if open_files and report_path.exists():
        print(f"\nOpening report: {report_path}")
        if platform.system() == "Windows":
            os.startfile(report_path)  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            subprocess.run(["open", str(report_path)])
        else:
            subprocess.run(["xdg-open", str(report_path)])

    # 6. Print log/events info
    log_env = os.environ.get("EXCEL_BOT_LOG_PATH")
    log_path = Path(log_env) if log_env else (root / "logs" / "events.jsonl")
    last_event = _read_last_event(log_path)
    print("\nRun summary")
    print(f"- Mode: {mode_label}")
    print(f"- Input folder: {input_dir}")
    print(f"- Output folder: {output_dir}")
    print(f"- Report file: {report_path}")
    print(f"- Log file: {log_path}")
    if last_event:
        print(f"- Last event: {_format_last_event(last_event)}")
    else:
        print("- Last event: none yet")

    # 7. Open log file if it exists
    if open_files and log_path.exists():
        print(f"\nOpening log file: {log_path}")
        if platform.system() == "Windows":
            os.startfile(log_path)  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            subprocess.run(["open", str(log_path)])
        else:
            subprocess.run(["xdg-open", str(log_path)])

    print("\nExcel Bot finished successfully.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
