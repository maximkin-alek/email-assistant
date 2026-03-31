from __future__ import annotations

import datetime as dt
import email
from email.message import Message
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime

from bs4 import BeautifulSoup


def _decode_rfc2047(value: str | None) -> str | None:
    if not value:
        return None
    try:
        return str(make_header(decode_header(value))).strip() or None
    except Exception:
        return value.strip() or None


def _clean_subject(subject: str | None) -> str | None:
    if not subject:
        return None
    s = subject.strip()
    # Частая проблема: в теме длинный токен/хэш без пробелов.
    if len(s) >= 60 and (" " not in s):
        allowed = sum(ch.isalnum() or ch in "-_=" for ch in s)
        if allowed / max(1, len(s)) > 0.95:
            return s[:60] + "…"
    return s


def _html_to_text(html: str) -> str:
    soup = BeautifulSoup(html, "lxml")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    text = soup.get_text("\n")
    lines = [line.strip() for line in text.splitlines()]
    lines = [line for line in lines if line]
    return "\n".join(lines)


def extract_text_from_message(msg: Message) -> str:
    if msg.is_multipart():
        parts = []
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = part.get_content_disposition()
            if disp == "attachment":
                continue
            if ctype == "text/plain":
                payload = part.get_payload(decode=True) or b""
                charset = part.get_content_charset() or "utf-8"
                parts.append(payload.decode(charset, errors="replace"))
            elif ctype == "text/html":
                payload = part.get_payload(decode=True) or b""
                charset = part.get_content_charset() or "utf-8"
                parts.append(_html_to_text(payload.decode(charset, errors="replace")))
        return "\n\n".join([p.strip() for p in parts if p and p.strip()]).strip()

    ctype = msg.get_content_type()
    payload = msg.get_payload(decode=True) or b""
    charset = msg.get_content_charset() or "utf-8"
    raw = payload.decode(charset, errors="replace")
    if ctype == "text/html":
        return _html_to_text(raw).strip()
    return raw.strip()


def parse_eml(raw_eml: bytes) -> dict:
    msg = email.message_from_bytes(raw_eml)
    message_id = (msg.get("Message-Id") or msg.get("Message-ID") or "").strip()
    subject = _clean_subject(_decode_rfc2047(msg.get("Subject")))
    from_email = _decode_rfc2047(msg.get("From"))

    date_header = msg.get("Date")
    parsed_date: dt.datetime | None = None
    if date_header:
        try:
            parsed_date = parsedate_to_datetime(date_header)
            if parsed_date and parsed_date.tzinfo is None:
                parsed_date = parsed_date.replace(tzinfo=dt.UTC)
        except Exception:
            parsed_date = None

    body_text = extract_text_from_message(msg) or None
    snippet = None
    if body_text:
        snippet = body_text[:400].replace("\r", "").strip() or None

    return {
        "provider_message_id": message_id or None,
        "from_email": from_email,
        "subject": subject,
        "date": parsed_date,
        "body_text": body_text,
        "snippet": snippet,
    }

