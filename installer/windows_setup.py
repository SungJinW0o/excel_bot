#!/usr/bin/env python3
"""Windows installer entrypoint for Excel Bot.

This script is bundled into setup.exe using PyInstaller.
It installs project files into %%LOCALAPPDATA%%\\ExcelBot and prepares a venv.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path


def _payload_root() -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / "payload"


def _target_dir() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "ExcelBot"
    return Path.home() / "AppData" / "Local" / "ExcelBot"


def _run(cmd: list[str], cwd: Path | None = None) -> None:
    print("> " + " ".join(cmd))
    subprocess.run(cmd, check=True, cwd=str(cwd) if cwd else None)


def _find_python_launcher() -> list[str]:
    candidates = [
        ["py", "-3"],
        ["python"],
    ]
    for candidate in candidates:
        try:
            subprocess.run(
                candidate + ["--version"],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            return candidate
        except Exception:
            continue
    raise RuntimeError(
        "Python 3.9+ is required but was not found. Install Python and rerun setup."
    )


def _copy_payload(src: Path, dst: Path) -> None:
    dst.mkdir(parents=True, exist_ok=True)
    for item in src.iterdir():
        target = dst / item.name
        if item.is_dir():
            shutil.copytree(item, target, dirs_exist_ok=True)
        else:
            shutil.copy2(item, target)


def _ensure_virtualenv(target: Path, launcher: list[str]) -> None:
    venv_dir = target / ".venv"
    venv_python = venv_dir / "Scripts" / "python.exe"

    if not venv_python.exists():
        print("Creating virtual environment...")
        _run(launcher + ["-m", "venv", str(venv_dir)])

    print("Installing dependencies into virtual environment...")
    _run([str(venv_python), "-m", "pip", "install", "--upgrade", "pip"])
    _run([str(venv_python), "-m", "pip", "install", "pandas>=2.0", "openpyxl>=3.1"])


def _maybe_pause() -> None:
    if os.name == "nt" and sys.stdin and sys.stdin.isatty():
        try:
            input("\nPress Enter to exit setup...")
        except EOFError:
            pass


def main() -> int:
    try:
        print("Excel Bot Setup")
        payload = _payload_root()
        if not payload.exists():
            raise FileNotFoundError(f"Installer payload not found: {payload}")

        target = _target_dir()
        print(f"Installing to: {target}")
        _copy_payload(payload, target)

        launcher = _find_python_launcher()
        _ensure_virtualenv(target, launcher)

        launch_bat = target / "launch_excel_bot.bat"
        print("\nInstallation complete.")
        print(f"Launch command: {launch_bat}")
        return 0
    except Exception as exc:
        print(f"\nERROR: {exc}")
        return 1
    finally:
        _maybe_pause()


if __name__ == "__main__":
    raise SystemExit(main())
