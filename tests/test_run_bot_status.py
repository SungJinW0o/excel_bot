import os
import subprocess
import sys
import shutil
import uuid
from pathlib import Path
from unittest.mock import patch

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import excel_bot.run_bot as run_bot


def test_run_bot_reports_no_valid_data_exit_code(capsys):
    base_dir = Path(__file__).resolve().parent / "tmp"
    base_dir.mkdir(exist_ok=True)
    temp_dir = base_dir / f"run_{uuid.uuid4().hex}"
    temp_dir.mkdir()

    previous_cwd = Path.cwd()
    os.chdir(temp_dir)
    try:
        with (
            patch.object(sys, "argv", ["run_bot.py", "--dry-run", "true", "--headless"]),
            patch.object(run_bot, "_resolve_python", return_value=sys.executable),
            patch.object(
                run_bot.subprocess,
                "run",
                side_effect=subprocess.CalledProcessError(
                    returncode=2,
                    cmd=[sys.executable, "-m", "excel_bot.bot_main"],
                ),
            ),
        ):
            exit_code = run_bot.main()
    finally:
        os.chdir(previous_cwd)
        shutil.rmtree(temp_dir, ignore_errors=True)

    output = capsys.readouterr().out
    assert exit_code == 2
    assert "no valid input data" in output.lower()


def test_run_pipeline_uses_in_process_mode_when_frozen():
    with (
        patch.object(run_bot.sys, "frozen", True, create=True),
        patch.object(run_bot.subprocess, "run") as subprocess_run_mock,
        patch("excel_bot.bot_main.main", return_value=None) as bot_main_mock,
    ):
        code = run_bot._run_pipeline(sys.executable)

    assert code == 0
    bot_main_mock.assert_called_once_with()
    subprocess_run_mock.assert_not_called()


def test_resolve_python_supports_venv_folder():
    base_dir = Path(__file__).resolve().parent / "tmp"
    base_dir.mkdir(exist_ok=True)
    temp_dir = base_dir / f"run_{uuid.uuid4().hex}"
    scripts_dir = temp_dir / "venv" / "Scripts"
    scripts_dir.mkdir(parents=True)
    candidate = scripts_dir / "python.exe"
    candidate.write_text("", encoding="utf-8")

    try:
        with (
            patch.object(run_bot.platform, "system", return_value="Windows"),
            patch.object(run_bot, "_python_has_runtime_dependencies", return_value=True),
        ):
            resolved = run_bot._resolve_python(temp_dir)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    assert resolved == str(candidate)


def test_resolve_python_falls_back_when_venv_missing_runtime_deps():
    base_dir = Path(__file__).resolve().parent / "tmp"
    base_dir.mkdir(exist_ok=True)
    temp_dir = base_dir / f"run_{uuid.uuid4().hex}"
    scripts_dir = temp_dir / "venv" / "Scripts"
    scripts_dir.mkdir(parents=True)
    candidate = scripts_dir / "python.exe"
    candidate.write_text("", encoding="utf-8")

    try:
        with (
            patch.object(run_bot.platform, "system", return_value="Windows"),
            patch.object(run_bot, "_python_has_runtime_dependencies", return_value=False),
        ):
            resolved = run_bot._resolve_python(temp_dir)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    assert resolved == sys.executable


def test_read_last_event_uses_last_non_empty_line():
    base_dir = Path(__file__).resolve().parent / "tmp"
    base_dir.mkdir(exist_ok=True)
    temp_dir = base_dir / f"run_{uuid.uuid4().hex}"
    temp_dir.mkdir()
    log_path = temp_dir / "events.jsonl"
    log_path.write_text(
        '{"type":"PIPELINE_STARTED"}\n'
        '\n'
        '{"type":"PIPELINE_COMPLETED","level":"INFO","timestamp":"2026-02-14T00:00:00+00:00"}\n',
        encoding="utf-8",
    )

    try:
        event = run_bot._read_last_event(log_path)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    assert event is not None
    assert event["type"] == "PIPELINE_COMPLETED"


def test_read_last_event_returns_raw_for_non_json_line():
    base_dir = Path(__file__).resolve().parent / "tmp"
    base_dir.mkdir(exist_ok=True)
    temp_dir = base_dir / f"run_{uuid.uuid4().hex}"
    temp_dir.mkdir()
    log_path = temp_dir / "events.jsonl"
    log_path.write_text("first\n\nnot-json-line\n", encoding="utf-8")

    try:
        event = run_bot._read_last_event(log_path)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    assert event == {"raw": "not-json-line"}
