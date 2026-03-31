from __future__ import annotations

import datetime as dt
import email
import re
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


def _extract_links_and_images_from_html(html: str, limit: int = 50) -> tuple[list[str], list[str]]:
    soup = BeautifulSoup(html, "lxml")
    links: list[str] = []
    images: list[str] = []
    seen_l: set[str] = set()
    seen_i: set[str] = set()
    for a in soup.find_all("a"):
        href = (a.get("href") or "").strip()
        if href.startswith("http") and href not in seen_l:
            seen_l.add(href)
            links.append(href)
            if len(links) >= limit:
                break
    for img in soup.find_all("img"):
        src = (img.get("src") or "").strip()
        # Только внешние картинки. cid: без вложений мы не восстановим.
        if src.startswith("http") and src not in seen_i:
            seen_i.add(src)
            images.append(src)
            if len(images) >= limit:
                break
    return links, images


def _extract_links_from_text(text: str, limit: int = 50) -> list[str]:
    urls = re.findall(r"https?://[^\s<>()\"']+", text or "")
    out: list[str] = []
    seen: set[str] = set()
    for u in urls:
        u = u.strip().rstrip(".,;:!?)\"]}")
        if not u or u in seen:
            continue
        seen.add(u)
        out.append(u)
        if len(out) >= limit:
            break
    return out


def extract_parts_from_message(msg: Message) -> tuple[str, str | None, list[str], list[str]]:
    if msg.is_multipart():
        parts: list[str] = []
        html_parts: list[str] = []
        links: list[str] = []
        images: list[str] = []
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
                html = payload.decode(charset, errors="replace")
                html_parts.append(html)
                parts.append(_html_to_text(html))
                l2, i2 = _extract_links_and_images_from_html(html)
                links.extend(l2)
                images.extend(i2)
        text = "\n\n".join([p.strip() for p in parts if p and p.strip()]).strip()
        html_out = "\n\n".join([h.strip() for h in html_parts if h and h.strip()]).strip() or None
        if text:
            links.extend(_extract_links_from_text(text))
        # дедупликация
        links = list(dict.fromkeys([x for x in links if x]))[:50]
        images = list(dict.fromkeys([x for x in images if x]))[:50]
        return text, html_out, links, images

    ctype = msg.get_content_type()
    payload = msg.get_payload(decode=True) or b""
    charset = msg.get_content_charset() or "utf-8"
    raw = payload.decode(charset, errors="replace")
    if ctype == "text/html":
        text = _html_to_text(raw).strip()
        links, images = _extract_links_and_images_from_html(raw)
        if text:
            links.extend(_extract_links_from_text(text))
        links = list(dict.fromkeys([x for x in links if x]))[:50]
        images = list(dict.fromkeys([x for x in images if x]))[:50]
        return text, raw.strip() or None, links, images
    text = raw.strip()
    links = _extract_links_from_text(text) if text else []
    return text, None, links, []


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

    body_text, body_html, links, images = extract_parts_from_message(msg)
    body_text = body_text or None
    snippet = None
    if body_text:
        snippet = body_text[:400].replace("\r", "").strip() or None

    return {
        "provider_message_id": message_id or None,
        "from_email": from_email,
        "subject": subject,
        "date": parsed_date,
        "body_text": body_text,
        "body_html": body_html,
        "extracted_links": links,
        "extracted_images": images,
        "snippet": snippet,
    }


def extract_attachments_from_eml(raw_eml: bytes, limit: int = 20, max_bytes: int = 2_000_000) -> list[dict]:
    """
    Возвращает список вложений (включая inline), подходящих для MVP.
    Ограничения по количеству/размеру — чтобы не раздувать БД.
    """
    msg = email.message_from_bytes(raw_eml)
    out: list[dict] = []
    for part in msg.walk():
        ctype = (part.get_content_type() or "").lower()
        disp = (part.get_content_disposition() or "").lower()
        filename = part.get_filename()
        cid = (part.get("Content-ID") or "").strip()
        if cid.startswith("<") and cid.endswith(">"):
            cid = cid[1:-1].strip()

        is_attachment = disp == "attachment"
        is_inline = disp == "inline" or (cid != "")

        if not (is_attachment or is_inline):
            continue
        if ctype in {"text/plain", "text/html"} and not filename:
            continue

        data = part.get_payload(decode=True) or b""
        if not data:
            continue
        if len(data) > max_bytes:
            continue

        out.append(
            {
                "filename": filename,
                "content_type": ctype or None,
                "size_bytes": len(data),
                "content_id": cid or None,
                "is_inline": bool(is_inline),
                "data": data,
            }
        )
        if len(out) >= limit:
            break
    return out

