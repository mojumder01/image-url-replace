"""Email OTP reader tests with a fake IMAP mailbox."""

import os
import sys
import time
import unittest
from email.message import EmailMessage
from email.utils import formatdate
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from catprep.common import StepError  # noqa: E402
from catprep.otp_email import wait_for_otp  # noqa: E402

CREDS = {"email_address": "me@gmail.com", "email_app_password": "app pass"}
CFG = {"sender": "no-reply@cartup.com", "subject_contains": "OTP", "wait_seconds": 1}


def make_mail(subject, body, sent, sender="no-reply@cartup.com", html=False):
    m = EmailMessage()
    m["Subject"], m["From"], m["Date"] = subject, sender, formatdate(sent, localtime=True)
    m.set_content(body, subtype="html" if html else "plain")
    return m.as_bytes()


class FakeImap:
    def __init__(self, mails):
        self.mails = mails  # oldest first

    def login(self, *a):
        pass

    def select(self, *a):
        pass

    def search(self, charset, *criteria):
        ids = [str(i + 1).encode() for i, m in enumerate(self.mails)
               if "FROM" not in criteria or criteria[criteria.index("FROM") + 1].strip('"').encode() in m]
        return "OK", [b" ".join(ids)]

    def fetch(self, msg_id, what):
        return "OK", [(None, self.mails[int(msg_id) - 1])]

    def logout(self):
        pass


class OtpTest(unittest.TestCase):
    def run_with(self, mails, since):
        with mock.patch("imaplib.IMAP4_SSL", return_value=FakeImap(mails)), mock.patch("time.sleep"):
            return wait_for_otp(CREDS, CFG, since)

    def test_newest_code_after_sign_in(self):
        now = time.time()
        mails = [make_mail("Your OTP", "Your OTP is 111111", now - 600),            # old: ignored
                 make_mail("Newsletter", "Sale 50% code 999999", now + 2, "shop@x.com"),
                 make_mail("Your OTP", "<p>Your login OTP is <b>482913</b>.</p>", now + 5, html=True)]
        self.assertEqual(self.run_with(mails, since=now - 30), "482913")

    def test_old_code_never_used(self):
        now = time.time()
        with self.assertRaises(StepError):
            self.run_with([make_mail("Your OTP", "Your OTP is 111111", now - 600)], since=now - 30)

    def test_code_next_to_otp_word_wins(self):
        from catprep.otp_email import find_code
        self.assertEqual(find_code("Order 2026 placed. Your OTP code: 5521 (valid 5 minutes)"), "5521")
        self.assertEqual(find_code("Use 778899 to sign in"), "778899")

    def test_missing_email_settings(self):
        with self.assertRaises(StepError):
            wait_for_otp({}, CFG, time.time())


if __name__ == "__main__":
    unittest.main()
