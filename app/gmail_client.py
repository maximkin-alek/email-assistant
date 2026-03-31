from __future__ import annotations

from dataclasses import dataclass

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from email.header import decode_header, make_header
import base64
from bs4 import BeautifulSoup

from app.email_parsing import _extract_links_from_text, _extract_links_and_images_from_html, _html_to_text


GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]


@dataclass(frozen=True)
class GmailProfile:
    email_address: str | None
    history_id: str | None


def build_gmail_service(creds: Credentials):
    return build("gmail", "v1", credentials=creds, cache_discovery=False)


def get_profile(service) -> GmailProfile:
    prof = service.users().getProfile(userId="me").execute()
    return GmailProfile(
        email_address=prof.get("emailAddress"),
        history_id=str(prof.get("historyId")) if prof.get("historyId") is not None else None,
    )


def extract_headers(payload: dict) -> dict[str, str]:
    headers = payload.get("headers") or []
    out: dict[str, str] = {}
    for h in headers:
        name = (h.get("name") or "").strip().lower()
        value = (h.get("value") or "").strip()
        if name and value and name not in out:
            # Gmail может отдавать RFC2047-закодированные заголовки.
            try:
                out[name] = str(make_header(decode_header(value))).strip()
            except Exception:
                out[name] = value
    return out


def _walk_parts(payload: dict) -> list[dict]:
    out: list[dict] = []
    if not payload:
        return out
    out.append(payload)
    for p in payload.get("parts") or []:
        if isinstance(p, dict):
            out.extend(_walk_parts(p))
    return out


def _decode_body_data(part: dict) -> str | None:
    body = part.get("body") or {}
    data = body.get("data")
    if not data or not isinstance(data, str):
        return None
    try:
        raw = base64.urlsafe_b64decode(data.encode("ascii"))
        return raw.decode("utf-8", errors="replace")
    except Exception:
        return None


def extract_bodies_from_gmail_payload(payload: dict) -> tuple[str | None, str | None, list[str], list[str]]:
    """
    Возвращает (text, html, links, images) из Gmail message payload (format=full).
    """
    text_parts: list[str] = []
    html_parts: list[str] = []
    links: list[str] = []
    images: list[str] = []

    for part in _walk_parts(payload):
        mime = (part.get("mimeType") or "").lower()
        if mime not in {"text/plain", "text/html"}:
            continue
        content = _decode_body_data(part)
        if not content:
            continue
        if mime == "text/plain":
            text_parts.append(content)
        else:
            html_parts.append(content)
            l2, i2 = _extract_links_and_images_from_html(content)
            links.extend(l2)
            images.extend(i2)

    html = "\n\n".join([h.strip() for h in html_parts if h and h.strip()]).strip() or None
    text = "\n\n".join([t.strip() for t in text_parts if t and t.strip()]).strip() or None
    if not text and html:
        text = _html_to_text(html).strip() or None
    if text:
        links.extend(_extract_links_from_text(text))

    # дедуп
    links = list(dict.fromkeys([x for x in links if x]))[:50]
    images = list(dict.fromkeys([x for x in images if x]))[:50]
    return text, html, links, images


def extract_attachments_from_gmail_message(service, message_id: str, payload: dict, limit: int = 20, max_bytes: int = 2_000_000) -> list[dict]:
    """
    Gmail attachments/inline images: скачиваем по attachmentId.
    """
    out: list[dict] = []
    for part in _walk_parts(payload):
        filename = (part.get("filename") or "").strip() or None
        mime = (part.get("mimeType") or "").strip() or None
        headers = part.get("headers") or []
        hdr_map = {}
        for h in headers:
            n = (h.get("name") or "").lower().strip()
            v = (h.get("value") or "").strip()
            if n and v and n not in hdr_map:
                hdr_map[n] = v
        cid = (hdr_map.get("content-id") or "").strip()
        if cid.startswith("<") and cid.endswith(">"):
            cid = cid[1:-1].strip()
        disp = (hdr_map.get("content-disposition") or "").lower()
        is_inline = ("inline" in disp) or bool(cid)
        is_attachment = ("attachment" in disp) or bool(filename)

        body = part.get("body") or {}
        attach_id = body.get("attachmentId")
        size = body.get("size")
        if not attach_id:
            continue
        if size and isinstance(size, int) and size > max_bytes:
            continue
        if not (is_attachment or is_inline):
            continue

        a = service.users().messages().attachments().get(userId="me", messageId=message_id, id=attach_id).execute()
        data = a.get("data")
        if not data or not isinstance(data, str):
            continue
        raw = base64.urlsafe_b64decode(data.encode("ascii"))
        if len(raw) > max_bytes:
            continue

        out.append(
            {
                "filename": filename,
                "content_type": mime,
                "size_bytes": len(raw),
                "content_id": cid or None,
                "is_inline": bool(is_inline),
                "data": raw,
            }
        )
        if len(out) >= limit:
            break
    return out

