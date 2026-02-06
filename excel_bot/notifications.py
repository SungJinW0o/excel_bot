import os
import smtplib
from email.message import EmailMessage
from typing import List, Optional

from .auth import load_users
from .events import DEFAULT_LOG_PATH, emit_event


SMTP_HOST = os.getenv("SMTP_HOST", "smtp.example.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASS = os.getenv("SMTP_PASS")
SMTP_SENDER = os.getenv("SMTP_SENDER", SMTP_USER)
DRY_RUN = os.getenv("DRY_RUN", "false").lower() in ("1", "true", "yes")


def get_recipients_by_role(role: str = "admin") -> List[str]:
    users = load_users()
    recipients = []
    for user in users.values():
        if user.status != "active":
            continue
        if user.role == role and user.email:
            recipients.append(user.email)
    return recipients


def send_email(
    subject: str,
    body: str,
    recipients: List[str],
    attachments: Optional[List[str]] = None,
    smtp_host: str = SMTP_HOST,
    smtp_port: int = SMTP_PORT,
    smtp_user: Optional[str] = SMTP_USER,
    smtp_pass: Optional[str] = SMTP_PASS,
    sender: Optional[str] = SMTP_SENDER,
) -> None:
    subject_prefix = "[DRY RUN] " if DRY_RUN else ""
    subject_with_prefix = f"{subject_prefix}{subject}"
    attachment_names = [os.path.basename(path) for path in (attachments or [])]
    if DRY_RUN:
        body = body + "\n\nNOTE: This is a dry run. No email was actually sent."
        print(f"[DRY_RUN] Email would be sent to: {recipients}")
        print(f"[DRY_RUN] Email subject: {subject_with_prefix}")
        print(f"[DRY_RUN] Email body:\n{body}")
        emit_event(
            "EMAIL_SENT",
            user_id="system",
            payload={
                "recipients": recipients,
                "dry_run": True,
                "subject": subject_with_prefix,
                "attachments": attachment_names,
            },
            log_path=DEFAULT_LOG_PATH,
        )
        return

    required_vars = ["SMTP_HOST", "SMTP_PORT", "SMTP_USER", "SMTP_PASS"]
    missing = [var for var in required_vars if not os.getenv(var)]
    if missing or smtp_host == "smtp.example.com":
        details = missing[:]
        if smtp_host == "smtp.example.com":
            details.append("SMTP_HOST default")
        raise RuntimeError(
            "SMTP configuration incomplete. Missing or default values: "
            + ", ".join(details)
        )

    if not recipients:
        print("No recipients found, skipping email.")
        return

    msg = EmailMessage()
    msg["From"] = sender or smtp_user or "no-reply@example.com"
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject_with_prefix
    msg.set_content(body)

    attachments = attachments or []
    for filepath in attachments:
        if not os.path.exists(filepath):
            continue
        with open(filepath, "rb") as f:
            data = f.read()
            filename = os.path.basename(filepath)
            msg.add_attachment(
                data,
                maintype="application",
                subtype="octet-stream",
                filename=filename,
            )

    try:
        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.starttls()
            if smtp_user and smtp_pass:
                server.login(smtp_user, smtp_pass)
            server.send_message(msg)
        print(f"Email sent to {len(recipients)} recipient(s).")
        emit_event(
            "EMAIL_SENT",
            user_id="system",
            payload={
                "recipients": recipients,
                "subject": subject,
                "attachments": attachment_names,
            },
            log_path=DEFAULT_LOG_PATH,
        )
    except Exception as exc:
        print("Failed to send email:", exc)
        emit_event(
            "EMAIL_FAILED",
            user_id="system",
            payload={
                "error": str(exc),
                "recipients": recipients,
                "subject": subject,
                "attachments": attachment_names,
            },
            log_path=DEFAULT_LOG_PATH,
        )


def notify_pipeline_completed(cleaned_file: str, report_file: str) -> None:
    recipients = get_recipients_by_role("admin")
    subject = "Pipeline Completed"
    body = (
        "The data pipeline has completed successfully.\n\n"
        f"Cleaned file: {cleaned_file}\n"
        f"Report file: {report_file}"
    )

    attachments = [path for path in [cleaned_file, report_file] if os.path.exists(path)]

    send_email(
        subject=subject,
        body=body,
        recipients=recipients,
        attachments=attachments,
        smtp_host=SMTP_HOST,
        smtp_port=SMTP_PORT,
        smtp_user=SMTP_USER,
        smtp_pass=SMTP_PASS,
        sender=SMTP_SENDER,
    )


def notify_pipeline_started() -> None:
    recipients = get_recipients_by_role("admin")
    subject = "Pipeline Started"
    body = "The data pipeline has begun execution."

    send_email(
        subject=subject,
        body=body,
        recipients=recipients,
        smtp_host=SMTP_HOST,
        smtp_port=SMTP_PORT,
        smtp_user=SMTP_USER,
        smtp_pass=SMTP_PASS,
        sender=SMTP_SENDER,
    )


def notify_data_cleaned(cleaned_file: str) -> None:
    recipients = get_recipients_by_role("admin")
    subject = "Data Cleaned"
    body = f"Cleaned data is ready: {cleaned_file}"

    attachments = [cleaned_file] if os.path.exists(cleaned_file) else []

    send_email(
        subject=subject,
        body=body,
        recipients=recipients,
        attachments=attachments,
        smtp_host=SMTP_HOST,
        smtp_port=SMTP_PORT,
        smtp_user=SMTP_USER,
        smtp_pass=SMTP_PASS,
        sender=SMTP_SENDER,
    )


def notify_pipeline_failed(error_msg: str) -> None:
    recipients = get_recipients_by_role("admin")
    subject = "Pipeline Failed"
    body = f"The data pipeline encountered an error:\n{error_msg}"

    send_email(
        subject=subject,
        body=body,
        recipients=recipients,
        smtp_host=SMTP_HOST,
        smtp_port=SMTP_PORT,
        smtp_user=SMTP_USER,
        smtp_pass=SMTP_PASS,
        sender=SMTP_SENDER,
    )
