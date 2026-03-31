from __future__ import annotations

from dataclasses import dataclass

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from email.header import decode_header, make_header


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

