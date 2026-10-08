"""Email a short report after an automatic run (uses the same Gmail app password as the OTP)."""

import smtplib
from email.message import EmailMessage

from .common import warn
from .site import load_credentials


def send(cfg, subject, body):
    n = cfg.get("notify", {})
    if not n.get("enabled"):
        return
    creds = load_credentials(cfg.path(cfg["export"]["credentials_file"]))
    sender, password = creds.get("email_address"), creds.get("email_app_password")
    to = n.get("to") or sender
    if not (sender and password and to):
        warn("Report email not sent: email_address / email_app_password missing in credentials.json.")
        return
    msg = EmailMessage()
    msg["Subject"], msg["From"], msg["To"] = subject, sender, to
    msg.set_content(body)
    try:
        with smtplib.SMTP_SSL(n.get("smtp_host", "smtp.gmail.com"), int(n.get("smtp_port", 465)), timeout=30) as s:
            s.login(sender, password)
            s.send_message(msg)
    except Exception as e:
        warn(f"Report email not sent: {e}")
