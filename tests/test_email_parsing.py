from __future__ import annotations

from app.email_parsing import parse_eml


def test_parse_simple_eml_extracts_text_and_snippet() -> None:
    raw = (
        b"From: Alice <alice@example.com>\r\n"
        b"To: me@example.com\r\n"
        b"Subject: Test subject\r\n"
        b"Message-ID: <abc123@example.com>\r\n"
        b"Date: Mon, 30 Mar 2026 10:00:00 +0000\r\n"
        b"Content-Type: text/plain; charset=utf-8\r\n"
        b"\r\n"
        b"Hello!\r\n"
        b"Please confirm the code: 123456\r\n"
    )
    parsed = parse_eml(raw)
    assert parsed["provider_message_id"] == "<abc123@example.com>"
    assert parsed["from_email"].startswith("Alice")
    assert parsed["subject"] == "Test subject"
    assert "confirm" in (parsed["body_text"] or "").lower()
    assert parsed["snippet"] is not None

