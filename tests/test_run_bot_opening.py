import sys
from pathlib import Path
from unittest.mock import patch

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import excel_bot.run_bot as run_bot


def test_open_report_prefers_excel_on_windows():
    report = Path("summary_report.xlsx")
    with (
        patch.object(run_bot.platform, "system", return_value="Windows"),
        patch.object(run_bot, "_find_excel_executable_windows", return_value=r"C:\Office\EXCEL.EXE"),
        patch.object(run_bot.subprocess, "Popen") as popen_mock,
        patch.object(run_bot.os, "startfile", create=True) as startfile_mock,
    ):
        opened = run_bot._open_report_file(report)

    assert opened is True
    popen_mock.assert_called_once_with([r"C:\Office\EXCEL.EXE", str(report)])
    startfile_mock.assert_not_called()


def test_open_report_falls_back_to_default_when_excel_missing():
    report = Path("summary_report.xlsx")
    with (
        patch.object(run_bot.platform, "system", return_value="Windows"),
        patch.object(run_bot, "_find_excel_executable_windows", return_value=None),
        patch.object(run_bot.os, "startfile", create=True) as startfile_mock,
    ):
        opened = run_bot._open_report_file(report)

    assert opened is True
    startfile_mock.assert_called_once_with(report)


def test_open_report_uses_xdg_open_on_linux():
    report = Path("summary_report.xlsx")
    with (
        patch.object(run_bot.platform, "system", return_value="Linux"),
        patch.object(run_bot.subprocess, "run") as run_mock,
    ):
        opened = run_bot._open_report_file(report)

    assert opened is True
    run_mock.assert_called_once_with(["xdg-open", str(report)], check=False)


def test_open_log_uses_default_opener_not_excel():
    log_file = Path("logs/events.jsonl")
    with (
        patch.object(run_bot.platform, "system", return_value="Windows"),
        patch.object(run_bot.os, "startfile", create=True) as startfile_mock,
        patch.object(run_bot, "_find_excel_executable_windows") as excel_find_mock,
    ):
        opened = run_bot._open_log_file(log_file)

    assert opened is True
    startfile_mock.assert_called_once_with(log_file)
    excel_find_mock.assert_not_called()
