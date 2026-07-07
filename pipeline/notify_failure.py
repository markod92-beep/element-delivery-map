"""notify_failure.py - send a single failure-notification email.

Factored out of smoke_test.py so the daily_refresh.ps1 catch block can notify
on ANY step failure (not just the smoke test). Uses the EXACT same Gmail SMTP
mechanism and GMAIL_APP_PASSWORD env var that smoke_test.py's _send_email uses,
so there is one shared, auditable mail path.

Silently no-ops (exit 0) if GMAIL_APP_PASSWORD is not set — matching
smoke_test.py — so a missing secret never masks the original failure.

Usage:
    python notify_failure.py --subject "Delivery Map refresh FAILED at <step>" \
        --body-file refresh-v2\\LAST_RUN_FAILED.flag
    python notify_failure.py --subject "..." --body "inline body text"
"""
from __future__ import annotations

import argparse
import os
import smtplib
import ssl
import sys
from email.message import EmailMessage
from pathlib import Path


def send_email(subject: str, body: str) -> None:
    """Email via Gmail SMTP. Silently no-ops if GMAIL_APP_PASSWORD not set.

    Mirrors smoke_test.py::_send_email exactly (same server, port, env vars).
    """
    pwd = os.environ.get("GMAIL_APP_PASSWORD")
    if not pwd:
        print("  [notify] GMAIL_APP_PASSWORD not set — email skipped.", file=sys.stderr)
        return
    sender = os.environ.get("GMAIL_SENDER", "markod92@gmail.com")
    recipient = os.environ.get("QA_NOTIFY_EMAIL", "markod92@gmail.com")
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(body)
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context()) as s:
            s.login(sender, pwd)
            s.send_message(msg)
        print(f"  [notify] failure email sent to {recipient}.", file=sys.stderr)
    except Exception as e:
        # Don't let email failure mask the actual pipeline failure.
        print(f"  [notify] email send failed: {e}", file=sys.stderr)


def main() -> int:
    p = argparse.ArgumentParser(description="Send a failure-notification email.")
    p.add_argument("--subject", required=True, help="Email subject line.")
    g = p.add_mutually_exclusive_group(required=True)
    g.add_argument("--body", help="Inline email body text.")
    g.add_argument("--body-file", help="Path to a file whose contents become the body.")
    args = p.parse_args()

    if args.body_file:
        try:
            body = Path(args.body_file).read_text(encoding="utf-8")
        except Exception as e:
            body = f"(could not read body file {args.body_file}: {e})"
    else:
        body = args.body or ""

    send_email(args.subject, body)
    return 0


if __name__ == "__main__":
    sys.exit(main())
