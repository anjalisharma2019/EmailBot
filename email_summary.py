#!/usr/bin/env python3
from __future__ import annotations
"""
Email Weekly Summary Tool
Connects to Gmail via IMAP, fetches primary inbox emails from the last 7 days,
and generates a summary report with per-email briefs.
 
Setup:
  1. Enable IMAP in Gmail: Settings → See all settings → Forwarding and POP/IMAP → Enable IMAP
  2. Create an App Password (if 2FA is on):
     https://myaccount.google.com/apppasswords
  3. Install dependencies:
     pip install python-dotenv
 
Usage:
  python email_summary.py
"""
 
import imaplib
import email.message
import email.header
import email.utils
import email.parser
from email.header import decode_header
from email.utils import parsedate_to_datetime
import os
import re
import textwrap
from datetime import datetime, timezone, timedelta
from dotenv import load_dotenv
 
# ── Config ──────────────────────────────────────────────────────────────────
 
load_dotenv()  # loads EMAIL and APP_PASSWORD from a .env file if present
 
EMAIL_ADDRESS = os.getenv("GMAIL_ADDRESS") or input("Gmail address: ").strip()
APP_PASSWORD  = os.getenv("GMAIL_APP_PASSWORD") or input("App password (no spaces): ").strip()
 
IMAP_HOST   = "imap.gmail.com"
IMAP_PORT   = 993
CUTOFF_DAYS = 7          # only emails newer than this
MAX_BODY_CHARS = 2000    # characters to read per email body for the summary
 
# ── Helpers ──────────────────────────────────────────────────────────────────
 
def decode_str(raw) -> str:
    """Decode an encoded email header value to a plain string."""
    if raw is None:
        return ""
    parts = decode_header(raw)
    decoded = []
    for part, enc in parts:
        if isinstance(part, bytes):
            try:
                decoded.append(part.decode(enc or "utf-8", errors="replace"))
            except (LookupError, TypeError):
                decoded.append(part.decode("latin-1", errors="replace"))
        else:
            decoded.append(part)
    return "".join(decoded)
 
 
def get_text_body(msg: email.message.Message) -> str:
    """Extract plain-text body from an email message."""
    body_parts = []
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get("Content-Disposition", ""))
            if ct == "text/plain" and "attachment" not in cd:
                charset = part.get_content_charset() or "utf-8"
                try:
                    body_parts.append(part.get_payload(decode=True).decode(charset, errors="replace"))
                except Exception:
                    pass
    else:
        charset = msg.get_content_charset() or "utf-8"
        try:
            body_parts.append(msg.get_payload(decode=True).decode(charset, errors="replace"))
        except Exception:
            pass
    return "\n".join(body_parts)
 
 
def clean_body(text: str) -> str:
    """Strip quoted replies, excess whitespace, and HTML artifacts."""
    # Remove common reply markers
    text = re.sub(r"(?m)^(>.*|On .{10,80} wrote:)$", "", text)
    # Collapse blank lines
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()
 
 
def one_line_summary(subject: str, body: str) -> str:
    """
    Build a brief (≤2-sentence) summary from subject + body.
    This is a simple heuristic — no external AI call needed.
    """
    # Prefer the first substantive paragraph of the body
    paragraphs = [p.strip() for p in body.split("\n\n") if len(p.strip()) > 40]
    excerpt = paragraphs[0] if paragraphs else body[:300]
    # Collapse newlines inside the excerpt
    excerpt = re.sub(r"\s+", " ", excerpt).strip()
    # Truncate to ~200 chars
    if len(excerpt) > 200:
        excerpt = excerpt[:197] + "..."
    if excerpt:
        return excerpt
    return "(No readable body)"
 
 
def fetch_primary_emails(mail: imaplib.IMAP4_SSL, cutoff: datetime) -> list[dict]:
    """
    Select [Gmail]/All Mail is NOT what we want — we use INBOX which in Gmail
    corresponds to the Primary tab when combined with the NOT X-GM-LABELS filter.
    Gmail exposes category labels as: ^Promotions, ^Social, ^Updates, ^Forums.
    We search INBOX and then discard anything tagged with those labels.
    """
    mail.select("INBOX")
 
    # Build IMAP date string (SINCE uses DD-Mon-YYYY)
    since_str = cutoff.strftime("%d-%b-%Y")
    status, data = mail.search(None, f'SINCE "{since_str}"')
    if status != "OK":
        return []
 
    uids = data[0].split()
    if not uids:
        return []
 
    emails = []
    for uid in uids:
        status, raw = mail.fetch(uid, "(RFC822)")
        if status != "OK":
            continue
 
        raw_bytes = raw[0][1]
        msg = email.parser.BytesParser().parsebytes(raw_bytes)
 
        # ── Filter out non-primary categories via X-GM-LABELS ──
        # Gmail sets X-GM-LABELS header for category labels
        labels_header = msg.get("X-GM-LABELS", "")
        skip_labels = {"\\Spam", "\\Trash", "Promotions", "Social", "Updates", "Forums"}
        if any(lbl in labels_header for lbl in skip_labels):
            continue
 
        # Also skip if Gmail-added category header is present (some clients see this)
        category = msg.get("X-Gmail-Labels", "")
        if any(lbl in category for lbl in {"Promotions", "Social", "Updates", "Forums"}):
            continue
 
        # ── Parse metadata ──
        subject = decode_str(msg.get("Subject", "(No subject)"))
        sender  = decode_str(msg.get("From", "(Unknown)"))
        date_str = msg.get("Date", "")
 
        try:
            dt = parsedate_to_datetime(date_str)
            # Normalise to UTC
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            else:
                dt = dt.astimezone(timezone.utc)
        except Exception:
            dt = datetime.now(timezone.utc)
 
        # Skip if somehow older than cutoff (can happen with timezone edge cases)
        if dt < cutoff:
            continue
 
        body  = get_text_body(msg)
        body  = clean_body(body)
        brief = one_line_summary(subject, body)
 
        emails.append({
            "uid":     uid.decode(),
            "subject": subject,
            "sender":  sender,
            "datetime": dt,
            "brief":   brief,
        })
 
    # Sort newest → oldest
    emails.sort(key=lambda e: e["datetime"], reverse=True)
    return emails
 
 
def render_report(emails: list[dict]) -> str:
    """Format the final human-readable report."""
    now   = datetime.now(timezone.utc)
    lines = []
 
    lines.append("=" * 70)
    lines.append("  WEEKLY EMAIL SUMMARY REPORT")
    lines.append(f"  Generated : {now.strftime('%A, %d %B %Y  %H:%M UTC')}")
    lines.append(f"  Period    : last {CUTOFF_DAYS} days  |  Total emails: {len(emails)}")
    lines.append("=" * 70)
    lines.append("")
 
    if not emails:
        lines.append("  No primary-inbox emails found in the last 7 days.")
        return "\n".join(lines)
 
    # ── Overall summary block ──
    lines.append("OVERVIEW")
    lines.append("-" * 70)
    sender_counts: dict[str, int] = {}
    for e in emails:
        # Extract just the name/address part
        m = re.search(r"<(.+?)>", e["sender"])
        key = m.group(1) if m else e["sender"]
        sender_counts[key] = sender_counts.get(key, 0) + 1
 
    top_senders = sorted(sender_counts.items(), key=lambda x: x[1], reverse=True)[:5]
    lines.append(f"  You received {len(emails)} email(s) in your Primary inbox this week.")
    lines.append(f"  Date range : {emails[-1]['datetime'].strftime('%d %b')} → {emails[0]['datetime'].strftime('%d %b %Y')}")
    lines.append("")
    lines.append("  Top senders:")
    for addr, cnt in top_senders:
        lines.append(f"    • {addr}  ({cnt} email{'s' if cnt > 1 else ''})")
    lines.append("")
 
    # ── Per-email entries ──
    lines.append("INDIVIDUAL EMAILS")
    lines.append("-" * 70)
    for i, e in enumerate(emails, 1):
        dt_local = e["datetime"].strftime("%a, %d %b %Y  %H:%M UTC")
        lines.append(f"[{i:03d}]  {dt_local}")
        lines.append(f"  From    : {e['sender']}")
        lines.append(f"  Subject : {e['subject']}")
        # Wrap the brief to 66 chars, indented
        wrapped = textwrap.fill(e["brief"], width=66, initial_indent="  Summary : ",
                                subsequent_indent="           ")
        lines.append(wrapped)
        lines.append("")
 
    lines.append("=" * 70)
    lines.append("  END OF REPORT")
    lines.append("=" * 70)
 
    return "\n".join(lines)
 
 
# ── Main ─────────────────────────────────────────────────────────────────────
 
def main():
    cutoff = datetime.now(timezone.utc) - timedelta(days=CUTOFF_DAYS)
 
    print(f"\nConnecting to {IMAP_HOST} …")
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(EMAIL_ADDRESS, APP_PASSWORD)
    except imaplib.IMAP4.error as exc:
        print(f"\n❌  Login failed: {exc}")
        print("   Make sure IMAP is enabled in Gmail and you're using an App Password.")
        return
 
    print(f"✓  Logged in as {EMAIL_ADDRESS}")
    print(f"   Fetching Primary inbox emails since {cutoff.strftime('%d %b %Y %H:%M UTC')} …\n")
 
    try:
        emails = fetch_primary_emails(mail, cutoff)
    finally:
        mail.logout()
 
    report = render_report(emails)
    print(report)
 
    # ── Save report to file ──
    timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_file   = f"email_summary_{timestamp}.txt"
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(report)
    print(f"\n📄  Report saved to: {out_file}")
 
 
if __name__ == "__main__":
    main()
 