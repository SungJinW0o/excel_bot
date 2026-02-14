import shutil
import subprocess
import sys
import uuid
from pathlib import Path
from unittest.mock import patch

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import installer.windows_setup as windows_setup


def test_ensure_pip_repairs_broken_pip():
    venv_python = Path(r"C:\temp\ExcelBot\.venv\Scripts\python.exe")
    bad_pip = subprocess.CalledProcessError(returncode=1, cmd=["pip", "--version"])
    ok_pip = subprocess.CompletedProcess(args=["pip", "--version"], returncode=0)

    with (
        patch.object(windows_setup, "_run") as run_mock,
        patch.object(windows_setup.subprocess, "run", side_effect=[bad_pip, ok_pip]),
    ):
        windows_setup._ensure_pip(venv_python)

    run_mock.assert_called_once_with([str(venv_python), "-m", "ensurepip", "--upgrade"])


def test_ensure_pip_skips_repair_when_pip_is_healthy():
    venv_python = Path(r"C:\temp\ExcelBot\.venv\Scripts\python.exe")
    ok_pip = subprocess.CompletedProcess(args=["pip", "--version"], returncode=0)

    with (
        patch.object(windows_setup, "_run") as run_mock,
        patch.object(windows_setup.subprocess, "run", return_value=ok_pip),
    ):
        windows_setup._ensure_pip(venv_python)

    run_mock.assert_not_called()


def test_pip_install_with_repair_retries_after_failure():
    venv_python = Path(r"C:\temp\ExcelBot\.venv\Scripts\python.exe")
    pip_args = ["install", "pandas>=2.0"]
    expected_cmd = [str(venv_python), "-m", "pip"] + pip_args

    with (
        patch.object(
            windows_setup,
            "_run",
            side_effect=[
                subprocess.CalledProcessError(returncode=1, cmd=expected_cmd),
                None,
            ],
        ) as run_mock,
        patch.object(windows_setup, "_ensure_pip") as ensure_pip_mock,
    ):
        windows_setup._pip_install_with_repair(venv_python, pip_args)

    ensure_pip_mock.assert_called_once_with(venv_python)
    assert run_mock.call_count == 2
    assert run_mock.call_args_list[0].args[0] == expected_cmd
    assert run_mock.call_args_list[1].args[0] == expected_cmd


def test_ensure_virtualenv_runs_expected_steps():
    base_dir = Path(__file__).resolve().parent / "tmp"
    base_dir.mkdir(exist_ok=True)
    temp_dir = base_dir / f"run_{uuid.uuid4().hex}"
    temp_dir.mkdir()
    launcher = ["py", "-3"]

    try:
        expected_venv_python = temp_dir / ".venv" / "Scripts" / "python.exe"
        with (
            patch.object(windows_setup, "_run") as run_mock,
            patch.object(windows_setup, "_ensure_pip") as ensure_pip_mock,
            patch.object(windows_setup, "_pip_install_with_repair") as pip_install_mock,
        ):
            windows_setup._ensure_virtualenv(temp_dir, launcher)

        run_mock.assert_called_once_with(launcher + ["-m", "venv", str(temp_dir / ".venv")])
        ensure_pip_mock.assert_called_once_with(expected_venv_python)
        assert pip_install_mock.call_count == 2
        assert pip_install_mock.call_args_list[0].args == (
            expected_venv_python,
            ["install", "--upgrade", "pip"],
        )
        assert pip_install_mock.call_args_list[1].args == (
            expected_venv_python,
            ["install", "pandas>=2.0", "openpyxl>=3.1", "PySide6>=6.7"],
        )
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
