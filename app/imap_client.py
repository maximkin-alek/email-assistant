from __future__ import annotations

import imaplib
import socket
from dataclasses import dataclass


@dataclass(frozen=True)
class ImapConfig:
    host: str
    port: int
    username: str
    password: str
    folder: str = "INBOX"
    timeout_s: int = 20


def _parse_uid_search(data: list[bytes | None]) -> list[int]:
    # IMAP search обычно возвращает: [b'1 2 3'] или [b'']
    if not data:
        return []
    chunk = data[0] or b""
    chunk = chunk.strip()
    if not chunk:
        return []
    return [int(x) for x in chunk.split() if x.isdigit()]


class ImapSession:
    def __init__(self, cfg: ImapConfig):
        self.cfg = cfg
        self.imap: imaplib.IMAP4_SSL | None = None

    def __enter__(self) -> "ImapSession":
        socket.setdefaulttimeout(self.cfg.timeout_s)
        self.imap = imaplib.IMAP4_SSL(self.cfg.host, self.cfg.port)
        self.imap.login(self.cfg.username, self.cfg.password)
        self.imap.select(self.cfg.folder)
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self.imap is None:
            return
        try:
            self.imap.logout()
        except Exception:
            try:
                self.imap.shutdown()
            except Exception:
                pass
        finally:
            self.imap = None

    def uid_search_new(self, last_uid: int | None) -> list[int]:
        assert self.imap is not None
        start_uid = (last_uid or 0) + 1
        uid_query = f"UID {start_uid}:*"
        status, data = self.imap.uid("search", None, uid_query)
        if status != "OK":
            return []
        return _parse_uid_search(data)

    def fetch_rfc822(self, uid: int) -> bytes | None:
        assert self.imap is not None
        status, data = self.imap.uid("fetch", str(uid), "(RFC822)")
        if status != "OK" or not data:
            return None
        for item in data:
            if isinstance(item, tuple) and len(item) == 2 and isinstance(item[1], (bytes, bytearray)):
                return bytes(item[1])
        return None

    def fetch_rfc822_and_flags(self, uid: int) -> tuple[bytes | None, set[str]]:
        """
        Возвращает (raw_rfc822, flags). flags — набор строк вида "\\Seen".
        """
        assert self.imap is not None
        status, data = self.imap.uid("fetch", str(uid), "(RFC822 FLAGS)")
        if status != "OK" or not data:
            return None, set()
        raw: bytes | None = None
        flags: set[str] = set()
        for item in data:
            if isinstance(item, tuple) and len(item) == 2:
                meta = item[0]
                payload = item[1]
                if isinstance(payload, (bytes, bytearray)) and payload:
                    raw = bytes(payload)
                if isinstance(meta, (bytes, bytearray)):
                    # Пример meta: b'123 (RFC822 {..} FLAGS (\\Seen))'
                    txt = bytes(meta).decode("utf-8", errors="ignore")
                    if "FLAGS" in txt:
                        start = txt.find("FLAGS")
                        if start != -1:
                            seg = txt[start:]
                            lp = seg.find("(")
                            rp = seg.find(")")
                            if lp != -1 and rp != -1 and rp > lp:
                                inside = seg[lp + 1 : rp]
                                for f in inside.split():
                                    if f:
                                        flags.add(f.strip())
        return raw, flags

    def mark_seen(self, uid: int) -> bool:
        """
        Пометить письмо как прочитанное (UID STORE +FLAGS \\Seen).
        """
        assert self.imap is not None
        try:
            status, _ = self.imap.uid("store", str(uid), "+FLAGS", "(\\Seen)")
            return status == "OK"
        except Exception:
            return False


def fetch_new_uids(cfg: ImapConfig, last_uid: int | None, limit: int = 50) -> list[int]:
    with ImapSession(cfg) as s:
        uids = s.uid_search_new(last_uid=last_uid)
    if limit and len(uids) > limit:
        uids = uids[-limit:]
    return uids


def fetch_rfc822_by_uid(cfg: ImapConfig, uid: int) -> bytes | None:
    with ImapSession(cfg) as s:
        return s.fetch_rfc822(uid)

