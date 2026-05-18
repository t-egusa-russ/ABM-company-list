"""SMTP経由でレポートメールを送信する"""

from __future__ import annotations

import os
import smtplib
import ssl
from email.message import EmailMessage
from typing import Iterable


def _split_emails(raw: str) -> list[str]:
    return [e.strip() for e in raw.replace(";", ",").split(",") if e.strip()]


def send_report_email(
    subject: str,
    html_body: str,
    attachment_bytes: bytes | None = None,
    attachment_filename: str = "weekly_subsidy_report.xlsx",
    *,
    smtp_host: str | None = None,
    smtp_port: int | None = None,
    smtp_user: str | None = None,
    smtp_password: str | None = None,
    mail_from: str | None = None,
    mail_to: str | Iterable[str] | None = None,
    use_ssl: bool | None = None,
) -> None:
    """SMTPでレポートメールを送信する。各引数はNoneの場合は環境変数を読む。

    必須環境変数: SMTP_HOST, SMTP_USER, SMTP_PASSWORD, MAIL_FROM, MAIL_TO
    任意環境変数: SMTP_PORT（デフォルト587）, SMTP_USE_SSL（"true"でSMTPS:465）
    """
    host = smtp_host or os.getenv("SMTP_HOST", "")
    port = smtp_port or int(os.getenv("SMTP_PORT", "587"))
    user = smtp_user or os.getenv("SMTP_USER", "")
    password = smtp_password or os.getenv("SMTP_PASSWORD", "")
    sender = mail_from or os.getenv("MAIL_FROM", user)

    if isinstance(mail_to, str):
        recipients = _split_emails(mail_to)
    elif mail_to:
        recipients = list(mail_to)
    else:
        recipients = _split_emails(os.getenv("MAIL_TO", ""))

    if use_ssl is None:
        use_ssl = os.getenv("SMTP_USE_SSL", "false").lower() in ("1", "true", "yes")

    missing = [
        name for name, val in {
            "SMTP_HOST": host, "SMTP_USER": user, "SMTP_PASSWORD": password,
            "MAIL_FROM": sender,
        }.items() if not val
    ]
    if not recipients:
        missing.append("MAIL_TO")
    if missing:
        raise RuntimeError(f"SMTP設定が不足: {', '.join(missing)}")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg.set_content("このメールはHTML形式です。HTMLが表示できる環境でご覧ください。")
    msg.add_alternative(html_body, subtype="html")

    if attachment_bytes:
        msg.add_attachment(
            attachment_bytes,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=attachment_filename,
        )

    context = ssl.create_default_context()
    if use_ssl:
        with smtplib.SMTP_SSL(host, port, context=context, timeout=30) as smtp:
            smtp.login(user, password)
            smtp.send_message(msg)
    else:
        with smtplib.SMTP(host, port, timeout=30) as smtp:
            smtp.ehlo()
            smtp.starttls(context=context)
            smtp.ehlo()
            smtp.login(user, password)
            smtp.send_message(msg)
