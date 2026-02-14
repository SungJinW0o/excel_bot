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


def _python_has_runtime_dependencies(python_exe: str) -> bool:
    try:
        subprocess.run(
            [python_exe, "-c", "import pandas, openpyxl"],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return True
    except Exception:
        return False


def _read_last_event(log_path: Path) -> Optional[Dict[str, Any]]:
    if not log_path.exists():
        return None
    try:
        with log_path.open("r", encoding="utf-8") as f:
            last = ""
            for line in f:
                stripped = line.strip()
                if stripped:
                    last = stripped
        if not last:
            return None
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
    is_windows = platform.system() == "Windows"
    for folder_name in (".venv", "venv"):
        venv_dir = root / folder_name
        if not venv_dir.exists():
            continue
        if is_windows:
            candidate = venv_dir / "Scripts" / "python.exe"
        else:
            candidate = venv_dir / "bin" / "python"
        if candidate.exists():
            candidate_str = str(candidate)
            if _python_has_runtime_dependencies(candidate_str):
                print(f"Using virtual environment: {candidate}")
                return candidate_str
            print(
                f"Virtual environment '{folder_name}' is missing required runtime packages "
                "(pandas/openpyxl). Trying next Python."
            )
            continue
        print(
            f"Virtual environment folder '{folder_name}' detected but Python was not found. "
            "Using system Python."
        )
    print("No usable virtual environment detected (.venv or venv), using system Python.")
    return sys.executable


def _find_excel_executable_windows() -> Optional[str]:
    env_candidate = os.environ.get("EXCEL_BOT_EXCEL_PATH")
    if env_candidate:
        env_path = Path(env_candidate).expanduser()
        if env_path.exists():
            return str(env_path)

    try:
        import winreg  # type: ignore[import-not-found]

        key_paths = [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
        ]
        for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            for key_path in key_paths:
                try:
                    with winreg.OpenKey(hive, key_path) as key:
                        exe_path, _ = winreg.QueryValueEx(key, None)
                    candidate = Path(str(exe_path))
                    if candidate.exists():
                        return str(candidate)
                except OSError:
                    continue
    except Exception:
        pass

    base_dirs = [
        os.environ.get("ProgramFiles"),
        os.environ.get("ProgramFiles(x86)"),
    ]
    office_dirs = [
        ("Microsoft Office", "root", "Office16"),
        ("Microsoft Office", "Office16"),
        ("Microsoft Office", "Office15"),
        ("Microsoft Office", "Office14"),
    ]
    for base in base_dirs:
        if not base:
            continue
        for office_dir in office_dirs:
            candidate = Path(base).joinpath(*office_dir, "EXCEL.EXE")
            if candidate.exists():
                return str(candidate)

    return None


def _open_with_default_app(path: Path) -> bool:
    try:
        if platform.system() == "Windows":
            os.startfile(path)  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
        return True
    except Exception as exc:
        print(f"Failed to open file '{path}': {exc}")
        return False


def _open_with_excel_windows(path: Path) -> bool:
    excel_exe = _find_excel_executable_windows()
    if not excel_exe:
        return False

    try:
        subprocess.Popen([excel_exe, str(path)])
        return True
    except Exception as exc:
        print(f"Excel launch failed for '{path}': {exc}")
        return False


def _open_report_file(path: Path) -> bool:
    is_excel_file = path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}
    if platform.system() == "Windows" and is_excel_file:
        if _open_with_excel_windows(path):
            return True
        print("Excel executable not found or failed. Falling back to default file opener.")
    return _open_with_default_app(path)


def _open_log_file(path: Path) -> bool:
    return _open_with_default_app(path)


def _run_pipeline(python_exe: str) -> int:
    if getattr(sys, "frozen", False):
        try:
            from excel_bot import bot_main

            bot_main.main()
            return 0
        except SystemExit as exc:
            if isinstance(exc.code, int):
                return exc.code
            return 1
        except Exception as exc:
            print(f"Bot failed: {exc}")
            return 1

    try:
        subprocess.run([python_exe, "-m", "excel_bot.bot_main"], check=True)
        return 0
    except subprocess.CalledProcessError as exc:
        return exc.returncode


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
    print("\nRunning Excel Bot pipeline...\n")
    pipeline_exit_code = _run_pipeline(python_exe)
    if pipeline_exit_code != 0:
        print(f"Bot exited with code {pipeline_exit_code}")
        if pipeline_exit_code == 2:
            print("Run stopped because no valid input data was available after validation.")
            print("Check input files and review output_data/data_quality_issues.csv if present.")
        return pipeline_exit_code

    # 5. Open summary report
    report_path = output_dir / "summary_report.xlsx"
    if open_files and report_path.exists():
        print(f"\nOpening report: {report_path}")
        _open_report_file(report_path)

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
        _open_log_file(log_path)

    print("\nExcel Bot finished successfully.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
