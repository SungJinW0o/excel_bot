import os
import smtplib
from email.message import EmailMessage
from typing import List, Optional

from .auth import load_users
from .events import emit_event


def _is_dry_run() -> bool:
    return os.getenv("DRY_RUN", "false").lower() in ("1", "true", "yes")


def _smtp_host() -> str:
    return os.getenv("SMTP_HOST", "smtp.example.com")


def _smtp_port() -> int:
    return int(os.getenv("SMTP_PORT", "587"))


def _smtp_user() -> Optional[str]:
    return os.getenv("SMTP_USER")


def _smtp_pass() -> Optional[str]:
    return os.getenv("SMTP_PASS")


def _smtp_sender() -> Optional[str]:
    return os.getenv("SMTP_SENDER", _smtp_user())


def _strict_email_mode() -> bool:
    return os.getenv("EXCEL_BOT_STRICT_EMAIL", "false").lower() in ("1", "true", "yes")


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
    smtp_host: Optional[str] = None,
    smtp_port: Optional[int] = None,
    smtp_user: Optional[str] = None,
    smtp_pass: Optional[str] = None,
    sender: Optional[str] = None,
) -> None:
    if smtp_host is None:
        smtp_host = _smtp_host()
    if smtp_port is None:
        smtp_port = _smtp_port()
    if smtp_user is None:
        smtp_user = _smtp_user()
    if smtp_pass is None:
        smtp_pass = _smtp_pass()
    if sender is None:
        sender = _smtp_sender()

    dry_run = _is_dry_run()
    subject_prefix = "[DRY RUN] " if dry_run else ""
    subject_with_prefix = f"{subject_prefix}{subject}"
    attachment_names = [os.path.basename(path) for path in (attachments or [])]
    if dry_run:
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
        )
        return

    missing = []
    if not smtp_host:
        missing.append("SMTP_HOST")
    if not smtp_port:
        missing.append("SMTP_PORT")
    if not smtp_user:
        missing.append("SMTP_USER")
    if not smtp_pass:
        missing.append("SMTP_PASS")
    if missing or smtp_host == "smtp.example.com":
        details = missing[:]
        if smtp_host == "smtp.example.com":
            details.append("SMTP_HOST default")
        message = (
            "SMTP configuration incomplete. Missing or default values: "
            + ", ".join(details)
        )
        print(message)
        emit_event(
            "EMAIL_FAILED",
            user_id="system",
            payload={
                "error": message,
                "recipients": recipients,
                "subject": subject,
                "attachments": attachment_names,
            },
        )
        if _strict_email_mode():
            raise RuntimeError(message)
        return

    if not recipients:
        print("No recipients found, skipping email.")
        emit_event(
            "EMAIL_SKIPPED",
            user_id="system",
            payload={"reason": "no_recipients", "subject": subject_with_prefix},
        )
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
        )
        if _strict_email_mode():
            raise


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
    )


def notify_pipeline_started() -> None:
    recipients = get_recipients_by_role("admin")
    subject = "Pipeline Started"
    body = "The data pipeline has begun execution."

    send_email(
        subject=subject,
        body=body,
        recipients=recipients,
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
    )


def notify_pipeline_failed(error_msg: str) -> None:
    recipients = get_recipients_by_role("admin")
    subject = "Pipeline Failed"
    body = f"The data pipeline encountered an error:\n{error_msg}"

    send_email(
        subject=subject,
        body=body,
        recipients=recipients,
    )
