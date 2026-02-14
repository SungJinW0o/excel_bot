import importlib
import json
import os
import shutil
import sys
import uuid
from pathlib import Path
from unittest.mock import patch

import pytest

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from excel_bot.events import EVENTS


def _reload_notifications():
    import excel_bot.notifications

    return importlib.reload(excel_bot.notifications)


def _expected_admin_recipients() -> list[str]:
    users_path = ROOT_DIR / "users.json"
    if not users_path.exists():
        return []
    with users_path.open("r", encoding="utf-8") as f:
        users = json.load(f)
    recipients = []
    for user in users:
        if user.get("status") != "active":
            continue
        if user.get("role") != "admin":
            continue
        email = user.get("email")
        if email:
            recipients.append(email)
    return recipients


def _expected_subject(subject: str) -> str:
    return f"[DRY RUN] {subject}"


def _make_temp_dir() -> Path:
    base_dir = Path(__file__).resolve().parent / "tmp"
    base_dir.mkdir(exist_ok=True)
    temp_dir = base_dir / f"run_{uuid.uuid4().hex}"
    temp_dir.mkdir()
    return temp_dir


def test_send_email_dry_run_includes_attachments(capsys):
    os.environ["DRY_RUN"] = "true"
    notifications = _reload_notifications()

    EVENTS.clear()

    temp_dir = _make_temp_dir()
    try:
        attachment_a = temp_dir / "cleaned_master.xlsx"
        attachment_b = temp_dir / "summary_report.xlsx"
        attachment_a.write_text("a", encoding="utf-8")
        attachment_b.write_text("b", encoding="utf-8")

        notifications.send_email(
            subject="Test Email",
            body="Hello world",
            recipients=["test@example.com"],
            attachments=[str(attachment_a), str(attachment_b)],
        )
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    output = capsys.readouterr().out
    event = EVENTS[-1]
    assert event["type"] == "EMAIL_SENT"
    assert event["payload"]["dry_run"] is True
    assert event["payload"]["recipients"] == ["test@example.com"]
    assert event["payload"]["subject"] == _expected_subject("Test Email")
    assert "cleaned_master.xlsx" in event["payload"]["attachments"]
    assert "summary_report.xlsx" in event["payload"]["attachments"]
    assert "[DRY_RUN] Email subject: [DRY RUN] Test Email" in output
    assert "Hello world" in output
    assert "NOTE: This is a dry run. No email was actually sent." in output


def test_notify_pipeline_started_dry_run_emits_event(capsys):
    os.environ["DRY_RUN"] = "true"
    notifications = _reload_notifications()

    EVENTS.clear()

    notifications.notify_pipeline_started()

    output = capsys.readouterr().out
    event = EVENTS[-1]
    assert event["type"] == "EMAIL_SENT"
    assert event["payload"]["dry_run"] is True
    assert event["payload"]["recipients"] == _expected_admin_recipients()
    assert event["payload"]["subject"] == _expected_subject("Pipeline Started")
    assert "The data pipeline has begun execution." in output


def test_notify_data_cleaned_dry_run_includes_attachment(capsys):
    os.environ["DRY_RUN"] = "true"
    notifications = _reload_notifications()

    EVENTS.clear()

    temp_dir = _make_temp_dir()
    try:
        cleaned_file = temp_dir / "cleaned_master.xlsx"
        cleaned_file.write_text("data", encoding="utf-8")

        notifications.notify_data_cleaned(str(cleaned_file))
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    output = capsys.readouterr().out
    event = EVENTS[-1]
    assert event["type"] == "EMAIL_SENT"
    assert event["payload"]["dry_run"] is True
    assert event["payload"]["recipients"] == _expected_admin_recipients()
    assert event["payload"]["subject"] == _expected_subject("Data Cleaned")
    assert event["payload"]["attachments"] == ["cleaned_master.xlsx"]
    assert "Cleaned data is ready:" in output
    assert "cleaned_master.xlsx" in output


def test_notify_pipeline_failed_dry_run_emits_event(capsys):
    os.environ["DRY_RUN"] = "true"
    notifications = _reload_notifications()

    EVENTS.clear()

    error_message = "Simulated pipeline error"
    notifications.notify_pipeline_failed(error_message)

    output = capsys.readouterr().out
    event = EVENTS[-1]
    assert event["type"] == "EMAIL_SENT"
    assert event["payload"]["dry_run"] is True
    assert event["payload"]["recipients"] == _expected_admin_recipients()
    assert event["payload"]["subject"] == _expected_subject("Pipeline Failed")
    assert event["payload"]["attachments"] == []
    assert "The data pipeline encountered an error:" in output
    assert error_message in output


def test_send_email_respects_runtime_dry_run_without_reload(capsys):
    os.environ["DRY_RUN"] = "true"
    notifications = _reload_notifications()

    EVENTS.clear()

    os.environ["DRY_RUN"] = "false"
    with patch.object(notifications.smtplib, "SMTP", side_effect=RuntimeError("smtp blocked")):
        notifications.send_email(
            subject="Runtime DRY_RUN toggle",
            body="Body",
            recipients=["test@example.com"],
            smtp_host="smtp.test.local",
            smtp_port=587,
            smtp_user="user@example.com",
            smtp_pass="secret",
            sender="sender@example.com",
        )

    output = capsys.readouterr().out
    event = EVENTS[-1]
    assert event["type"] == "EMAIL_FAILED"
    assert "Runtime DRY_RUN toggle" in event["payload"]["subject"]
    assert "This is a dry run" not in output


def test_send_email_missing_config_non_strict_emits_failed_without_raising(capsys):
    os.environ["DRY_RUN"] = "false"
    os.environ.pop("EXCEL_BOT_STRICT_EMAIL", None)
    for key in ["SMTP_HOST", "SMTP_PORT", "SMTP_USER", "SMTP_PASS", "SMTP_SENDER"]:
        os.environ.pop(key, None)

    notifications = _reload_notifications()
    EVENTS.clear()

    notifications.send_email(
        subject="Missing SMTP",
        body="body",
        recipients=["admin@example.com"],
    )

    output = capsys.readouterr().out
    event = EVENTS[-1]
    assert event["type"] == "EMAIL_FAILED"
    assert "SMTP configuration incomplete" in event["payload"]["error"]
    assert "SMTP configuration incomplete" in output


def test_send_email_missing_config_strict_mode_raises():
    os.environ["DRY_RUN"] = "false"
    os.environ["EXCEL_BOT_STRICT_EMAIL"] = "true"
    for key in ["SMTP_HOST", "SMTP_PORT", "SMTP_USER", "SMTP_PASS", "SMTP_SENDER"]:
        os.environ.pop(key, None)

    notifications = _reload_notifications()
    EVENTS.clear()

    with pytest.raises(RuntimeError, match="SMTP configuration incomplete"):
        notifications.send_email(
            subject="Missing SMTP strict",
            body="body",
            recipients=["admin@example.com"],
        )

    event = EVENTS[-1]
    assert event["type"] == "EMAIL_FAILED"
