from __future__ import annotations

import datetime as dt

from sqlalchemy import select

from app.db import session_scope
from app.crypto import decrypt_str, encrypt_str
import json
from google.oauth2.credentials import Credentials

from app.email_parsing import extract_attachments_from_eml, parse_eml
from app.gmail_client import (
    build_gmail_service,
    extract_attachments_from_gmail_message,
    extract_bodies_from_gmail_payload,
    extract_headers,
    get_profile,
)
from app.imap_client import ImapConfig, ImapSession
from app.models import AppSetting, EmailAttachment, EmailMessage, Mailbox
from app.ai_client import classify_and_summarize
from app.app_state import (
    AiRunStatus,
    AiTestResult,
    AiTestStatus,
    get_ai_stop_flag,
    get_ai_test_status,
    now_iso,
    set_ai_run_status,
    set_ai_stop_flag,
    set_ai_test_result,
    set_ai_test_status,
)
from app.queue import get_queue
from app.settings import settings
from rq.registry import DeferredJobRegistry, FailedJobRegistry, FinishedJobRegistry, ScheduledJobRegistry, StartedJobRegistry
import logging

log = logging.getLogger("email-assistant")


def sync_remote_mark_read(email_id: int) -> None:
    """
    Синхронизация статуса "прочитано" обратно в оригинальный ящик.
    Gmail: removeLabelIds=["UNREAD"]
    IMAP: UID STORE +FLAGS (\\Seen)
    """
    with session_scope() as s:
        msg = s.get(EmailMessage, email_id)
        if not msg:
            return
        if not msg.is_read:
            return
        mb = s.get(Mailbox, msg.mailbox_id)
        if not mb or not mb.is_enabled:
            return

        if mb.provider == "imap":
            # Мы умеем помечать прочитанным только если provider_message_id содержит UID.
            pmid = (msg.provider_message_id or "").strip().lower()
            uid = None
            if pmid.startswith("uid:"):
                tail = pmid.split(":", 1)[1].strip()
                if tail.isdigit():
                    uid = int(tail)
            if not uid:
                return
            if not (mb.imap_host_enc and mb.imap_user_enc and mb.imap_password_enc and mb.imap_port):
                return
            cfg = ImapConfig(
                host=decrypt_str(mb.imap_host_enc),
                port=int(mb.imap_port),
                username=decrypt_str(mb.imap_user_enc),
                password=decrypt_str(mb.imap_password_enc),
                folder=mb.imap_folder or "INBOX",
            )
            try:
                with ImapSession(cfg) as imap:
                    imap.mark_seen(uid)
            except Exception:
                return

        elif mb.provider == "gmail":
            if not mb.gmail_credentials_enc:
                return
            try:
                creds = Credentials.from_authorized_user_info(info=__import__("json").loads(decrypt_str(mb.gmail_credentials_enc)))
                service = build_gmail_service(creds)
                service.users().messages().modify(userId="me", id=msg.provider_message_id, body={"removeLabelIds": ["UNREAD"]}).execute()
                # если токен обновился — сохраним обратно
                with session_scope() as s2:
                    mb2 = s2.get(Mailbox, mb.id)
                    if mb2:
                        mb2.gmail_credentials_enc = encrypt_str(creds.to_json())
            except Exception:
                # важно не молчать: иначе кажется, что "прочитано" синхронизируется, но нет
                try:
                    log.exception(
                        "sync_remote_mark_read:gmail failed email_id=%s mid=%s mailbox_id=%s",
                        email_id,
                        msg.provider_message_id,
                        mb.id,
                    )
                except Exception:
                    pass
                return


def classify_basic(email_id: int) -> None:
    """
    Итерация 1: только базовая классификация без LLM.
    В итерации 3 подключим OpenRouter и заменим/усилим эту логику.
    """
    with session_scope() as s:
        msg = s.get(EmailMessage, email_id)
        if not msg:
            return

        text = f"{msg.subject or ''}\n{msg.from_email or ''}\n{msg.snippet or ''}".lower()
        score = 10
        category = "normal"

        if "unsubscribe" in text or "рассылка" in text or "newsletter" in text:
            category = "newsletter"
            score = 5

        urgent_markers = ["срочно", "urgent", "asap", "оплат", "счет", "invoice", "код", "подтверд"]
        if any(m in text for m in urgent_markers):
            category = "important"
            score = 80

        msg.category = category
        msg.score = score
        msg.summary = msg.snippet
        if not msg.date:
            msg.date = dt.datetime.now(dt.UTC)


def recompute_all_basic(limit: int = 200) -> int:
    with session_scope() as s:
        ids = list(s.scalars(select(EmailMessage.id).order_by(EmailMessage.id.desc()).limit(limit)))
    for email_id in ids:
        classify_basic(email_id)
    return len(ids)


def sync_imap_mailbox(mailbox_id: int, limit: int = 50) -> int:
    """
    Итерация 2: синхронизация IMAP (Яндекс/Mail.ru).
    Забираем новые письма по UID, сохраняем в БД и ставим базовую классификацию.
    """
    # Берём конфиг и помечаем “running”
    with session_scope() as s:
        mb = s.get(Mailbox, mailbox_id)
        if not mb or not mb.is_enabled or mb.provider != "imap":
            return 0

        mb.last_sync_at = dt.datetime.now(dt.UTC)
        mb.last_sync_status = "running"
        mb.last_sync_error = None
        mb.last_sync_count = 0

        if not (mb.imap_host_enc and mb.imap_user_enc and mb.imap_password_enc and mb.imap_port):
            mb.last_sync_status = "error"
            mb.last_sync_error = "IMAP настройки не заполнены"
            return 0

        cfg = ImapConfig(
            host=decrypt_str(mb.imap_host_enc),
            port=int(mb.imap_port),
            username=decrypt_str(mb.imap_user_enc),
            password=decrypt_str(mb.imap_password_enc),
            folder=mb.imap_folder or "INBOX",
            tls_verify=bool(getattr(mb, "imap_tls_verify", True)),
        )
        last_uid = mb.imap_last_uid

    inserted = 0
    max_uid = last_uid or 0
    q = get_queue()
    new_ids: list[int] = []

    try:
        with ImapSession(cfg) as imap:
            uids = imap.uid_search_new(last_uid=last_uid)
            if limit and len(uids) > limit:
                uids = uids[-limit:]

            for uid in uids:
                raw, flags = imap.fetch_rfc822_and_flags(uid)
                if not raw:
                    continue
                parsed = parse_eml(raw)
                provider_message_id = parsed["provider_message_id"] or f"uid:{uid}"
                is_read = "\\Seen" in flags

                with session_scope() as s:
                    exists = s.scalars(
                        select(EmailMessage.id).where(
                            EmailMessage.mailbox_id == mailbox_id,
                            EmailMessage.provider_message_id == provider_message_id,
                        )
                    ).first()
                    if exists:
                        max_uid = max(max_uid, uid)
                        continue

                    msg = EmailMessage(
                        mailbox_id=mailbox_id,
                        provider_message_id=provider_message_id,
                        thread_id=None,
                        from_email=parsed["from_email"],
                        subject=parsed["subject"],
                        date=parsed["date"],
                        snippet=parsed["snippet"],
                        body_text=parsed["body_text"],
                        body_html=parsed.get("body_html"),
                        extracted_links_json=json.dumps(parsed.get("extracted_links") or [], ensure_ascii=False),
                        extracted_images_json=json.dumps(parsed.get("extracted_images") or [], ensure_ascii=False),
                        is_read=is_read,
                    )
                    s.add(msg)
                    s.flush()
                    # вложения (включая inline cid картинки)
                    for att in extract_attachments_from_eml(raw, limit=20, max_bytes=2_000_000):
                        s.add(
                            EmailAttachment(
                                email_id=msg.id,
                                filename=att.get("filename"),
                                content_type=att.get("content_type"),
                                size_bytes=att.get("size_bytes"),
                                content_id=att.get("content_id"),
                                is_inline=bool(att.get("is_inline")),
                                data=att.get("data"),
                            )
                        )
                    inserted += 1
                    max_uid = max(max_uid, uid)
                    new_ids.append(msg.id)

                    q.enqueue(classify_basic, msg.id)

        with session_scope() as s:
            mb = s.get(Mailbox, mailbox_id)
            if mb:
                mb.imap_last_uid = max_uid if max_uid != (last_uid or 0) else mb.imap_last_uid
                mb.last_sync_at = dt.datetime.now(dt.UTC)
                mb.last_sync_status = "ok"
                mb.last_sync_count = inserted
                mb.last_sync_error = None

        # Автоматически запускаем AI только для реально новых писем.
        for email_id in new_ids:
            q.enqueue(ai_process_email, email_id)

        return inserted
    except Exception as e:
        import socket
        import ssl

        msg = str(e)[:2000]
        # Частая причина у новых ящиков: неверный host/порт/SSL или сеть блокирует соединение.
        if isinstance(e, socket.timeout) or "handshake operation timed out" in msg:
            msg = (
                "IMAP: таймаут TLS/SSL рукопожатия. Проверь host/порт (обычно 993), "
                "что IMAP включён и сеть/антивирус не блокирует соединение."
            )
        elif isinstance(e, ssl.SSLError):
            if "CERTIFICATE_VERIFY_FAILED" in msg or "certificate verify failed" in msg:
                msg = (
                    "IMAP: не удалось проверить сертификат (похоже, VPN/антивирус подменяет сертификаты). "
                    "Попробуй отключить VPN или выключи проверку TLS для этого ящика в настройках."
                )
            else:
                msg = f"IMAP: ошибка TLS/SSL: {msg}"
        with session_scope() as s:
            mb = s.get(Mailbox, mailbox_id)
            if mb:
                mb.last_sync_at = dt.datetime.now(dt.UTC)
                mb.last_sync_status = "error"
                mb.last_sync_count = 0
                mb.last_sync_error = msg
        return 0


def sync_gmail_mailbox(mailbox_id: int, limit: int = 50) -> int:
    """
    Итерация Gmail: импорт последних писем (минимально).
    """
    from google.oauth2.credentials import Credentials

    with session_scope() as s:
        mb = s.get(Mailbox, mailbox_id)
        if not mb or not mb.is_enabled or mb.provider != "gmail":
            return 0

        mb.last_sync_at = dt.datetime.now(dt.UTC)
        mb.last_sync_status = "running"
        mb.last_sync_error = None
        mb.last_sync_count = 0

        if not mb.gmail_credentials_enc:
            mb.last_sync_status = "error"
            mb.last_sync_error = "Gmail не подключён (нет токена)"
            return 0

        creds = Credentials.from_authorized_user_info(
            info=__import__("json").loads(decrypt_str(mb.gmail_credentials_enc)),
        )

    inserted = 0
    q = get_queue()
    new_ids: list[int] = []
    try:
        service = build_gmail_service(creds)
        prof = get_profile(service)

        with session_scope() as s:
            mb = s.get(Mailbox, mailbox_id)
            if mb:
                mb.gmail_email = prof.email_address
                if prof.history_id and not mb.gmail_last_history_id:
                    mb.gmail_last_history_id = prof.history_id

        # Инкрементально: если есть historyId — забираем только новые messageAdded.
        with session_scope() as s:
            mb2 = s.get(Mailbox, mailbox_id)
            start_hist = mb2.gmail_last_history_id if mb2 else None

        msg_ids: list[str] = []
        latest_hist: str | None = None
        if start_hist:
            try:
                hresp = (
                    service.users()
                    .history()
                    .list(userId="me", startHistoryId=start_hist, historyTypes=["messageAdded"], maxResults=limit)
                    .execute()
                )
                latest_hist = str(hresp.get("historyId")) if hresp.get("historyId") is not None else None
                history = hresp.get("history") or []
                for h in history:
                    for added in h.get("messagesAdded") or []:
                        m = added.get("message") or {}
                        mid = m.get("id")
                        if mid:
                            msg_ids.append(mid)
            except Exception:
                # если startHistoryId устарел/невалиден — падаем назад на list()
                msg_ids = []
                latest_hist = None

        if not msg_ids:
            resp = service.users().messages().list(userId="me", maxResults=limit).execute()
            msgs = resp.get("messages") or []
            msg_ids = [m.get("id") for m in msgs if m.get("id")]

        for mid in msg_ids:

            with session_scope() as s:
                exists_id = s.scalars(
                    select(EmailMessage.id).where(
                        EmailMessage.mailbox_id == mailbox_id,
                        EmailMessage.provider_message_id == mid,
                    )
                ).first()
                # Если письмо уже есть, но мы раньше сохраняли только metadata (без body),
                # то обновим поля контента/ссылок/картинок.
                if exists_id:
                    msg = s.get(EmailMessage, exists_id)
                    if msg:
                        need_body = (msg.body_text is None and msg.body_html is None and msg.extracted_links_json is None)
                        need_att = (
                            s.scalars(select(EmailAttachment.id).where(EmailAttachment.email_id == msg.id).limit(1)).first()
                            is None
                        )
                        if need_body or need_att:
                            full = service.users().messages().get(userId="me", id=mid, format="full").execute()
                            payload = full.get("payload") or {}
                            if need_body:
                                body_text, body_html, links, images = extract_bodies_from_gmail_payload(payload)
                                msg.body_text = body_text
                                msg.body_html = body_html
                                msg.extracted_links_json = json.dumps(links or [], ensure_ascii=False)
                                msg.extracted_images_json = json.dumps(images or [], ensure_ascii=False)
                            if need_att:
                                for att in extract_attachments_from_gmail_message(service, mid, payload, limit=20, max_bytes=2_000_000):
                                    s.add(
                                        EmailAttachment(
                                            email_id=msg.id,
                                            filename=att.get("filename"),
                                            content_type=att.get("content_type"),
                                            size_bytes=att.get("size_bytes"),
                                            content_id=att.get("content_id"),
                                            is_inline=bool(att.get("is_inline")),
                                            data=att.get("data"),
                                        )
                                    )
                    continue

            full = service.users().messages().get(userId="me", id=mid, format="full").execute()
            payload = full.get("payload") or {}
            label_ids = full.get("labelIds") or []
            is_read = True
            if isinstance(label_ids, list) and "UNREAD" in label_ids:
                is_read = False
            headers = extract_headers(payload)
            subject = headers.get("subject")
            from_email = headers.get("from")
            date_hdr = headers.get("date")
            thread_id = full.get("threadId")
            snippet = full.get("snippet")
            body_text, body_html, links, images = extract_bodies_from_gmail_payload(payload)

            parsed_date = None
            if date_hdr:
                try:
                    from email.utils import parsedate_to_datetime

                    parsed_date = parsedate_to_datetime(date_hdr)
                    if parsed_date and parsed_date.tzinfo is None:
                        parsed_date = parsed_date.replace(tzinfo=dt.UTC)
                except Exception:
                    parsed_date = None

            with session_scope() as s:
                msg = EmailMessage(
                    mailbox_id=mailbox_id,
                    provider_message_id=mid,
                    thread_id=thread_id,
                    from_email=from_email,
                    subject=subject,
                    date=parsed_date,
                    snippet=snippet,
                    body_text=body_text,
                    body_html=body_html,
                    extracted_links_json=json.dumps(links or [], ensure_ascii=False),
                    extracted_images_json=json.dumps(images or [], ensure_ascii=False),
                    is_read=is_read,
                )
                s.add(msg)
                s.flush()
                for att in extract_attachments_from_gmail_message(service, mid, payload, limit=20, max_bytes=2_000_000):
                    s.add(
                        EmailAttachment(
                            email_id=msg.id,
                            filename=att.get("filename"),
                            content_type=att.get("content_type"),
                            size_bytes=att.get("size_bytes"),
                            content_id=att.get("content_id"),
                            is_inline=bool(att.get("is_inline")),
                            data=att.get("data"),
                        )
                    )
                inserted += 1
                q.enqueue(classify_basic, msg.id)
                new_ids.append(msg.id)

        with session_scope() as s:
            mb = s.get(Mailbox, mailbox_id)
            if mb:
                mb.last_sync_at = dt.datetime.now(dt.UTC)
                mb.last_sync_status = "ok"
                mb.last_sync_count = inserted
                mb.last_sync_error = None
                # если токен обновился внутри google lib — сохраним обратно
                mb.gmail_credentials_enc = encrypt_str(creds.to_json())
                if latest_hist:
                    mb.gmail_last_history_id = latest_hist

        for email_id in new_ids:
            q.enqueue(ai_process_email, email_id)

        return inserted
    except Exception as e:
        with session_scope() as s:
            mb = s.get(Mailbox, mailbox_id)
            if mb:
                mb.last_sync_at = dt.datetime.now(dt.UTC)
                mb.last_sync_status = "error"
                mb.last_sync_count = 0
                mb.last_sync_error = str(e)[:2000]
        return 0


def ai_process_email(email_id: int) -> None:
    # Глобальная остановка: даём возможность быстро прервать длинный прогон.
    # В этом режиме считаем письмо НЕ обработанным.
    if get_ai_stop_flag():
        return
    with session_scope() as s:
        msg = s.get(EmailMessage, email_id)
        if not msg:
            return
        if msg.ai_done:
            return

        # Правила секретаря (простые): порог важности, whitelist/blacklist.
        def _get_setting(k: str, default: str = "") -> str:
            try:
                v = s.get(AppSetting, k)
                return (v.value or "").strip() if v else default
            except Exception:
                return default

        threshold_s = _get_setting("important_threshold", "70")
        try:
            threshold = int(threshold_s)
        except Exception:
            threshold = 70
        threshold = max(0, min(100, threshold))

        wl = _get_setting("sender_whitelist", "")
        bl = _get_setting("sender_blacklist", "")

        def _norm_lines(v: str) -> list[str]:
            out: list[str] = []
            for line in (v or "").splitlines():
                line = line.strip().lower()
                if not line or line.startswith("#"):
                    continue
                out.append(line)
            return out[:200]

        wl_list = _norm_lines(wl)
        bl_list = _norm_lines(bl)

        from_l = (msg.from_email or "").lower()
        subj_l = (msg.subject or "").lower()
        snippet_l = (msg.snippet or "").lower()
        body_l = (msg.body_text or "").lower()

        def _sender_matches(rule: str) -> bool:
            r = rule.strip().lower()
            if not r:
                return False
            if r.startswith("@") and ("@" in from_l):
                # @domain.tld
                return from_l.endswith(r)
            if r.startswith("domain:") and ("@" in from_l):
                dom = r.split(":", 1)[1].strip()
                return bool(dom and ("@" + dom) in from_l)
            if r.startswith("subject:"):
                term = r.split(":", 1)[1].strip()
                return bool(term and term in subj_l)
            return r in from_l

        def _looks_transactional_important() -> bool:
            """
            Защита от ложного "рассылка" по отправителю: даже у брендов бывают
            транзакционные/секьюрные письма (коды входа, подтверждения, счета).
            """
            text = f"{subj_l}\n{snippet_l}\n{body_l}"
            markers = [
                "код",
                "otp",
                "one-time",
                "одноразов",
                "подтверж",
                "verify",
                "verification",
                "вход",
                "login",
                "парол",
                "password",
                "security",
                "безопас",
                "2fa",
                "двухфактор",
                "счет",
                "invoice",
                "оплат",
                "payment",
            ]
            return any(m in text for m in markers)

        try:
            result = classify_and_summarize(
                subject=msg.subject,
                from_email=msg.from_email,
                snippet=msg.snippet,
                body_text=msg.body_text,
            )

            # применяем правила после AI (мягкий режим):
            # - не "ломаем" категорию в лоб, а сдвигаем score и только затем порогом переводим в important.
            score_i = int(result.score or 0)
            category_i = str(result.category or "normal")
            bl_hit = any(_sender_matches(x) for x in bl_list)
            if bl_hit and not _looks_transactional_important():
                # Считать рассылкой: понижаем важность и не даём стать important.
                score_i = min(score_i, 25)
                if category_i == "important":
                    category_i = "normal"
            if any(_sender_matches(x) for x in wl_list):
                # Приоритетные отправители: повышаем важность до порога.
                score_i = max(score_i, threshold)
            # порог важности
            if (category_i not in {"newsletter", "spam_candidate"}) and (int(score_i or 0) >= threshold):
                category_i = "important"

            msg.category = category_i
            msg.score = int(score_i)
            msg.summary = result.summary
            msg.ai_explanation = result.explanation
            msg.ai_model = result.model
            msg.ai_processed_at = dt.datetime.now(dt.UTC)
            msg.ai_done = True
        except Exception as e:
            # Не валим воркер пачкой: сохраняем ошибку на письмо, чтобы было видно в UI.
            raw = str(e)
            # Нормализуем типовые англ. сообщения в человеческий русский.
            if "402 Payment Required" in raw:
                raw = (
                    "OpenRouter вернул 402 (Payment Required): для этой модели/аккаунта нет доступного кредита или "
                    "не активирован биллинг. Даже 'free' модели могут требовать активный биллинг/кредит."
                )
            elif "OPENROUTER_API_KEY" in raw:
                raw = "Не задан OPENROUTER_API_KEY в .env (ключ OpenRouter)."
            elif "For more information check" in raw:
                raw = raw.split("For more information check", 1)[0].strip()
            elif "Expecting value: line 1 column 1 (char 0)" in raw:
                raw = (
                    "AI вернул ответ не в JSON-формате (или пустой). "
                    "Обычно это означает, что модель проигнорировала инструкцию 'верни JSON' "
                    "или прокси вернул пустой/текстовый ответ."
                )

            msg.ai_explanation = f"Ошибка AI: {raw[:500]}"
            msg.ai_model = None
            # ВАЖНО: не ставим ai_processed_at, чтобы можно было переобработать позже.
            msg.ai_processed_at = None
            msg.ai_done = False


def ai_run(limit: int = 50) -> dict:
    """
    Единая команда: запускает AI-обработку для писем, где ai_done = false.
    Обрабатывает последовательно в одном job, чтобы был понятный прогресс.
    """
    started_at = now_iso()
    set_ai_stop_flag(False)
    set_ai_run_status(
        AiRunStatus(
            running=True,
            started_at=started_at,
            updated_at=now_iso(),
            total=0,
            processed=0,
            ok=0,
            failed=0,
            message="Запущено",
        )
    )
    with session_scope() as s:
        ids = list(
            s.scalars(
                select(EmailMessage.id)
                .where(
                    EmailMessage.is_archived == False,  # noqa: E712
                    EmailMessage.ai_done == False,  # noqa: E712
                )
                .order_by(EmailMessage.date.desc().nullslast(), EmailMessage.id.desc())
                .limit(limit)
            )
        )

    total = len(ids)
    processed = 0
    ok = 0
    failed = 0
    set_ai_run_status(
        AiRunStatus(
            running=True,
            started_at=started_at,
            updated_at=now_iso(),
            total=total,
            processed=0,
            ok=0,
            failed=0,
            message="В процессе",
        )
    )

    for email_id in ids:
        if get_ai_stop_flag():
            set_ai_run_status(
                AiRunStatus(
                    running=False,
                    started_at=started_at,
                    updated_at=now_iso(),
                    finished_at=now_iso(),
                    total=total,
                    processed=processed,
                    ok=ok,
                    failed=failed,
                    message="Остановлено пользователем",
                )
            )
            return {"total": total, "processed": processed, "ok": ok, "failed": failed, "stopped": True}
        try:
            ai_process_email(email_id)
            ok += 1
        except Exception:
            failed += 1
        processed += 1
        if processed % 5 == 0 or processed == total:
            set_ai_run_status(
                AiRunStatus(
                    running=True,
                    started_at=started_at,
                    updated_at=now_iso(),
                    total=total,
                    processed=processed,
                    ok=ok,
                    failed=failed,
                    message="В процессе",
                )
            )

    set_ai_run_status(
        AiRunStatus(
            running=False,
            started_at=started_at,
            updated_at=now_iso(),
            finished_at=now_iso(),
            total=total,
            processed=processed,
            ok=ok,
            failed=failed,
            message="Готово",
        )
    )
    return {"total": total, "processed": processed, "ok": ok, "failed": failed}


def ai_stop() -> None:
    set_ai_stop_flag(True)
    q = get_queue()

    def is_ai_job_func(func_name: str | None) -> bool:
        if not func_name:
            return False
        # всё, что связано с AI
        return func_name.startswith("app.jobs.ai_") or func_name == "app.jobs.ai_process_email"

    removed = 0

    # 1) Очистка очереди (ожидающие job'ы)
    for job in list(q.jobs):
        if is_ai_job_func(getattr(job, "func_name", None)):
            try:
                q.remove(job)
            except Exception:
                pass
            try:
                job.delete()
            except Exception:
                pass
            removed += 1

    # 2) Очистка основных реестров (на случай deferred/scheduled/started/failed)
    registries = [
        StartedJobRegistry(q.name, connection=q.connection),
        ScheduledJobRegistry(q.name, connection=q.connection),
        DeferredJobRegistry(q.name, connection=q.connection),
        FailedJobRegistry(q.name, connection=q.connection),
        FinishedJobRegistry(q.name, connection=q.connection),
    ]
    for reg in registries:
        for job_id in list(reg.get_job_ids()):
            try:
                job = reg.job_class.fetch(job_id, connection=q.connection)
            except Exception:
                continue
            if is_ai_job_func(getattr(job, "func_name", None)):
                try:
                    reg.remove(job_id, delete_job=True)
                except Exception:
                    try:
                        reg.remove(job_id)
                    except Exception:
                        pass
                removed += 1

    # Обновим статус, чтобы в UI было видно, что стоп сработал.
    cur = None
    try:
        from app.app_state import get_ai_run_status

        cur = get_ai_run_status()
    except Exception:
        cur = None

    started_at = cur.started_at if cur else ""
    total = cur.total if cur else 0
    processed = cur.processed if cur else 0
    ok = cur.ok if cur else 0
    failed = cur.failed if cur else 0
    set_ai_run_status(
        AiRunStatus(
            running=False,
            started_at=started_at,
            finished_at=now_iso(),
            total=total,
            processed=processed,
            ok=ok,
            failed=failed,
            message=f"Остановлено. Очищено AI job'ов: {removed}",
        )
    )


def ai_retry_failed(limit: int = 50) -> int:
    """
    Переобработка писем, где сохранилась старая ошибка вида "AI error: ..." или "Ошибка AI: ...".
    """
    with session_scope() as s:
        q = (
            select(EmailMessage.id)
            .where(
                EmailMessage.ai_done == False,  # noqa: E712
                EmailMessage.ai_explanation.is_not(None),
            )
            .order_by(EmailMessage.id.desc())
            .limit(limit)
        )
        ids = list(s.scalars(q))

    qn = get_queue()
    for email_id in ids:
        qn.enqueue(ai_process_email, email_id)
    return len(ids)


def ai_retry_frozen_assignment_errors(limit: int = 500) -> int:
    """
    Разовая миграция: письма могли сохранить ошибку вида
    "Ошибка AI: cannot assign to field 'score/category'" из-за попытки менять frozen AiResult.
    Сбрасываем такие письма в состояние "нужно AI" и ставим в очередь.
    """
    changed = 0
    with session_scope() as s:
        msgs = list(
            s.scalars(
                select(EmailMessage)
                .where(
                    EmailMessage.ai_explanation.is_not(None),
                )
                .order_by(EmailMessage.id.desc())
                .limit(limit)
            )
        )
        ids: list[int] = []
        for m in msgs:
            expl = (m.ai_explanation or "")
            if "cannot assign to field" not in expl:
                continue
            m.ai_done = False
            m.ai_processed_at = None
            m.ai_model = None
            # очищаем ошибку, чтобы не засорять UI
            m.ai_explanation = None
            changed += 1
            ids.append(m.id)

    qn = get_queue()
    for email_id in ids:
        qn.enqueue(ai_process_email, email_id)
    return changed


def ai_reset_old_errors(limit: int = 500) -> int:
    """
    Миграция для уже существующих записей:
    - если в базе осталось "AI error: ..." — переводим на русский префикс
    - сбрасываем ai_processed_at в NULL, чтобы письмо можно было переобработать
    """
    changed = 0
    with session_scope() as s:
        q = (
            select(EmailMessage)
            .where(
                EmailMessage.ai_explanation.is_not(None),
            )
            .order_by(EmailMessage.id.desc())
            .limit(limit)
        )
        msgs = list(s.scalars(q))
        for m in msgs:
            if not m.ai_explanation:
                continue
            expl = m.ai_explanation
            if m.ai_explanation.startswith("AI error: "):
                m.ai_explanation = "Ошибка AI: " + m.ai_explanation[len("AI error: ") :]
                m.ai_processed_at = None
                m.ai_done = False
                changed += 1
            elif "OPENROUTER_API_KEY" in m.ai_explanation or "Payment Required" in m.ai_explanation:
                # На всякий случай: если там старый англ. текст, тоже разрешаем повторную обработку
                if not m.ai_explanation.startswith("Ошибка AI: "):
                    m.ai_explanation = "Ошибка AI: " + m.ai_explanation[:480]
                m.ai_processed_at = None
                m.ai_done = False
                changed += 1
            elif "For more information check" in m.ai_explanation:
                m.ai_explanation = m.ai_explanation.split("For more information check", 1)[0].strip()
                if not m.ai_explanation.startswith("Ошибка AI: "):
                    m.ai_explanation = "Ошибка AI: " + m.ai_explanation[:480]
                m.ai_processed_at = None
                m.ai_done = False
                changed += 1
            elif "Expecting value: line 1 column 1 (char 0)" in m.ai_explanation:
                m.ai_explanation = (
                    "Ошибка AI: AI вернул ответ не в JSON-формате (или пустой). "
                    "Нажми “Повторить AI для ошибок” после проверки модели."
                )
                m.ai_processed_at = None
                m.ai_done = False
                changed += 1
            # Нормализация самого частого англ. текста от httpx по 402
            if "Client error '402 Payment Required'" in expl:
                m.ai_explanation = (
                    "Ошибка AI: OpenRouter вернул 402 (Payment Required): для этой модели/аккаунта нет доступного кредита "
                    "или не активирован биллинг. Даже 'free' модели могут требовать активный биллинг/кредит."
                )
                m.ai_processed_at = None
                m.ai_done = False
                changed += 1
    return changed


def ai_reset_empty_explanations(limit: int = 500) -> int:
    """
    Письма могли быть помечены как обработанные, но AI не вернул explanation (раньше ставили 'Без объяснения').
    Сбрасываем такие записи в состояние "нужно переобработать".
    """
    changed = 0
    with session_scope() as s:
        q = (
            select(EmailMessage)
            .where(
                EmailMessage.ai_processed_at.is_not(None),
                EmailMessage.ai_explanation.is_not(None),
            )
            .order_by(EmailMessage.id.desc())
            .limit(limit)
        )
        msgs = list(s.scalars(q))
        for m in msgs:
            expl = (m.ai_explanation or "").strip()
            if expl == "Без объяснения" or expl == "AI не вернул explanation (переобработка не помогла)":
                m.ai_processed_at = None
                m.ai_explanation = None
                m.ai_model = None
                m.ai_done = False
                changed += 1

    qn = get_queue()
    # Переобрабатываем только то, что сбросили
    with session_scope() as s:
        ids = list(
            s.scalars(
                select(EmailMessage.id)
                .where(
                    EmailMessage.ai_done == False,  # noqa: E712
                    EmailMessage.ai_explanation.is_(None),
                )
                .order_by(EmailMessage.id.desc())
                .limit(min(limit, 500))
            )
        )
    for email_id in ids:
        qn.enqueue(ai_process_email, email_id)
    return changed


def ai_process_recent(limit: int = 50) -> int:
    """
    Обрабатываем последние письма, которые ещё не проходили AI.
    Дешёвый пред-фильтр: не трогаем рассылки с низким score из базовой модели.
    """
    with session_scope() as s:
        q = (
            select(EmailMessage.id)
            .where(EmailMessage.ai_done == False)  # noqa: E712
            .order_by(EmailMessage.date.desc().nullslast(), EmailMessage.id.desc())
            .limit(limit)
        )
        ids = list(s.scalars(q))

    qn = get_queue()
    for email_id in ids:
        qn.enqueue(ai_process_email, email_id)
    return len(ids)


def ai_reset_all(limit: int = 500) -> int:
    """
    Для тестирования/отладки: сбросить AI-состояние у последних писем.
    Делает письма "не обработано AI": ai_done=false, чистит ai_* поля.
    """
    changed = 0
    with session_scope() as s:
        msgs = list(s.scalars(select(EmailMessage).order_by(EmailMessage.id.desc()).limit(limit)))
        for m in msgs:
            m.ai_done = False
            m.ai_processed_at = None
            m.ai_explanation = None
            m.ai_model = None
            changed += 1
    return changed


def ai_reset_mailbox(mailbox_id: int, limit: int = 500) -> int:
    """
    Для тестирования/отладки: сбросить AI-состояние только для одного ящика.
    """
    changed = 0
    with session_scope() as s:
        msgs = list(
            s.scalars(
                select(EmailMessage)
                .where(EmailMessage.mailbox_id == mailbox_id)
                .order_by(EmailMessage.id.desc())
                .limit(limit)
            )
        )
        for m in msgs:
            m.ai_done = False
            m.ai_processed_at = None
            m.ai_explanation = None
            m.ai_model = None
            changed += 1
    return changed


def ai_test_model() -> None:
    started = now_iso()
    set_ai_test_status(AiTestStatus(running=True, started_at=started, message="В процессе"))
    try:
        r = classify_and_summarize(
            subject="Тестовое письмо",
            from_email="test@example.com",
            snippet="Пожалуйста, подтвердите встречу завтра в 10:00.",
            body_text=None,
        )
        finished = now_iso()
        set_ai_test_status(AiTestStatus(running=False, started_at=started, finished_at=finished, message="Готово"))
        set_ai_test_result(
            AiTestResult(
                ok=True,
                configured_base_url=settings.ai_base_url,
                configured_model=settings.ai_model,
                used_model=r.model,
                message=f"OK: category={r.category}, score={r.score}",
                tested_at=now_iso(),
            )
        )
    except Exception as e:
        finished = now_iso()
        set_ai_test_status(AiTestStatus(running=False, started_at=started, finished_at=finished, message="Ошибка"))
        # Сохраняем человеческое сообщение об ошибке
        set_ai_test_result(
            AiTestResult(
                ok=False,
                configured_base_url=settings.ai_base_url,
                configured_model=settings.ai_model,
                used_model="",
                message=str(e)[:500],
                tested_at=now_iso(),
            )
        )

