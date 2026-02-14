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
