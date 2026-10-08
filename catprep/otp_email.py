"""Read the seller-admin login OTP from an email inbox (IMAP).

credentials.json needs:
    "email_address": "you@gmail.com",
    "email_app_password": "xxxx xxxx xxxx xxxx"     (Gmail: Google account > Security > App passwords)

config.toml [otp] sets the server, which sender/subject to look for and the code pattern.
Only emails that arrived after the Sign In click are used, so an old code is never reused.
"""

import email
import imaplib
import re
import time
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime

from .common import StepError, info


def _text_of(msg):
    parts = []
    for part in msg.walk() if msg.is_multipart() else [msg]:
        ctype = part.get_content_type()
        if ctype in ("text/plain", "text/html"):
            payload = part.get_payload(decode=True) or b""
            text = payload.decode(part.get_content_charset() or "utf-8", errors="replace")
            if ctype == "text/html":
                text = re.sub(r"<style.*?</style>|<[^>]+>", " ", text, flags=re.S | re.I)
            parts.append(text)
    return " ".join(parts)


NEAR_WORD = r"(?i)(?:otp|code|pin|password|verification)\D{0,40}?\b(\d{4,8})\b"
ANY_CODE = r"\b(\d{4,8})\b"


def find_code(text, pattern=""):
    """With no pattern: digits right after 'OTP'/'code'/... first, else any 4-8 digit number."""
    for pat in ([pattern] if pattern else [NEAR_WORD, ANY_CODE]):
        m = re.search(pat, text)
        if m:
            return m.group(1) if m.groups() else m.group(0)
    return None


def wait_for_otp(creds, otp_cfg, since):
    address = creds.get("email_address")
    password = creds.get("email_app_password")
    if not address or not password:
        raise StepError("Email OTP: add email_address and email_app_password to credentials.json.")
    host = otp_cfg.get("imap_host", "imap.gmail.com")
    sender = otp_cfg.get("sender", "")
    subject = otp_cfg.get("subject_contains", "")
    pattern = otp_cfg.get("code_pattern", "")
    timeout = int(otp_cfg.get("wait_seconds", 120))

    info(f"Waiting up to {timeout}s for the OTP email in {address} ...")
    end = time.time() + timeout
    try:
        imap = imaplib.IMAP4_SSL(host)
        imap.login(address, password)
    except Exception as e:
        raise StepError(f"Email OTP: cannot log in to {host} as {address}: {e}")
    try:
        while time.time() < end:
            imap.select("INBOX")
            criteria = ["SINCE", time.strftime("%d-%b-%Y", time.localtime(since - 86400))]
            if sender:
                criteria += ["FROM", f'"{sender}"']
            status, data = imap.search(None, *criteria)
            ids = data[0].split() if status == "OK" and data and data[0] else []
            for msg_id in reversed(ids[-20:]):  # newest first
                status, fetched = imap.fetch(msg_id, "(RFC822)")
                if status != "OK":
                    continue
                msg = email.message_from_bytes(fetched[0][1])
                try:
                    sent = parsedate_to_datetime(msg.get("Date")).timestamp()
                except Exception:
                    continue
                if sent < since:
                    break  # older than the Sign In click
                subj = str(make_header(decode_header(msg.get("Subject", ""))))
                if subject and subject.lower() not in subj.lower():
                    continue
                code = find_code(subj + " " + _text_of(msg), pattern)
                if code:
                    return code
            time.sleep(5)
    finally:
        try:
            imap.logout()
        except Exception:
            pass
    raise StepError(f"Email OTP: no OTP email arrived within {timeout}s.")
