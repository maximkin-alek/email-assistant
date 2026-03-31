from __future__ import annotations

from app.imap_client import _parse_uid_search


def test_parse_uid_search_handles_empty() -> None:
    assert _parse_uid_search([]) == []
    assert _parse_uid_search([b""]) == []
    assert _parse_uid_search([b"   "]) == []


def test_parse_uid_search_parses_uids() -> None:
    assert _parse_uid_search([b"1 2 300"]) == [1, 2, 300]

