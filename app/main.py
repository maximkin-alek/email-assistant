from __future__ import annotations

import datetime as dt
import logging
import re
import json
import os
from urllib.parse import urlparse
from email.header import decode_header, make_header
from email.utils import parseaddr
import time

from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy import delete, func, or_, select, update
from rq.registry import DeferredJobRegistry, ScheduledJobRegistry, StartedJobRegistry

from app.db import Base, engine, session_scope
from app.crypto import encrypt_str
from app.jobs import (
    ai_run,
    ai_stop,
    ai_process_recent,
    ai_reset_all,
    ai_reset_mailbox,
    ai_reset_empty_explanations,
    ai_reset_old_errors,
    ai_retry_failed,
    ai_retry_frozen_assignment_errors,
    ai_test_model,
    recompute_all_basic,
    sync_remote_mark_read,
    sync_gmail_mailbox,
    sync_imap_mailbox,
)
from app.models import AppSetting, EmailAttachment, EmailMessage, Mailbox
from app.queue import get_queue
from app.schema import ensure_schema
from app.settings import settings
from app.oauth_state import consume_state, issue_state
from app.app_state import (
    AiRunStatus,
    AiTestStatus,
    get_ai_run_status,
    get_ai_test_result,
    get_ai_test_status,
    now_iso,
    set_ai_run_status,
    set_ai_stop_flag,
    set_ai_test_status,
)

app = FastAPI(title="Email Assistant")
templates = Jinja2Templates(directory="templates")
log = logging.getLogger("email-assistant")

app.mount("/static", StaticFiles(directory="static"), name="static")

# OAuthlib иногда поднимает предупреждение о расхождении scopes как исключение,
# из-за чего Gmail OAuth колбэк может падать без сохранения токена.
# Для нашего кейса это безопасно: мы расширяем права (readonly -> modify), и Google
# возвращает список фактически выданных scopes.
os.environ.setdefault("OAUTHLIB_RELAX_TOKEN_SCOPE", "1")


def _decode_rfc2047(value: str | None) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value))).strip()
    except Exception:
        return str(value).strip()


templates.env.filters["rfc2047"] = _decode_rfc2047

_WEEKDAYS_RU = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]


def _fmt_dt_with_weekday(value: dt.datetime | None) -> str:
    if not value:
        return ""
    try:
        # В БД чаще всего храним UTC; если tz отсутствует (naive) — считаем, что это UTC.
        if value.tzinfo is None:
            value = value.replace(tzinfo=dt.UTC)
        local = value.astimezone()
        wd = _WEEKDAYS_RU[local.weekday()]
        return f"{wd}, {local.strftime('%d.%m.%Y %H:%M')}"
    except Exception:
        try:
            wd = _WEEKDAYS_RU[value.weekday()]
            return f"{wd}, {value.strftime('%d.%m.%Y %H:%M')}"
        except Exception:
            return str(value)


templates.env.filters["fmt_dt"] = _fmt_dt_with_weekday


@app.middleware("http")
async def _slowlog_middleware(request: Request, call_next):
    t0 = time.monotonic()
    try:
        response = await call_next(request)
        return response
    finally:
        dt_ms = int((time.monotonic() - t0) * 1000)
        # Логируем только подозрительно медленные запросы, чтобы поймать "зависания" на перезагрузке.
        if dt_ms >= 1500:
            try:
                log.warning(f"slow_request {dt_ms}ms {request.method} {request.url.path}")
            except Exception:
                pass


_LEARN_CONFIRM_N = 2


def _extract_sender_key(from_email: str | None) -> str:
    raw = (from_email or "").strip()
    _, addr = parseaddr(raw)
    key = (addr or raw).strip().lower()
    return key


def _norm_lines(v: str) -> list[str]:
    out: list[str] = []
    for line in (v or "").splitlines():
        x = line.strip()
        if not x:
            continue
        if x not in out:
            out.append(x)
    return out[:500]


def _learn_on_manual_category_change(
    s,
    *,
    from_email: str | None,
    category: str,
) -> None:
    """
    "Обучение" без переобучения модели:
    после N одинаковых ручных действий добавляем отправителя в мягкое правило.
    """
    if category not in {"important", "newsletter"}:
        return

    sender_key = _extract_sender_key(from_email)
    if not sender_key:
        return

    cnt_key = f"learn_confirm:{category}"
    cur = s.get(AppSetting, cnt_key)
    try:
        counts = json.loads(cur.value or "{}") if cur and cur.value else {}
    except Exception:
        counts = {}
    if not isinstance(counts, dict):
        counts = {}

    n = int(counts.get(sender_key) or 0) + 1
    counts[sender_key] = n
    if not cur:
        s.add(AppSetting(key=cnt_key, value=json.dumps(counts, ensure_ascii=False)))
    else:
        cur.value = json.dumps(counts, ensure_ascii=False)

    if n < _LEARN_CONFIRM_N:
        return

    rule_key = "sender_whitelist" if category == "important" else "sender_blacklist"
    rule = s.get(AppSetting, rule_key)
    lines = _norm_lines(rule.value if rule else "")
    if sender_key not in lines:
        lines.append(sender_key)
        new_val = "\n".join(lines[:200])
        if not rule:
            s.add(AppSetting(key=rule_key, value=new_val))
        else:
            rule.value = new_val


def _extract_links(text: str | None, limit: int = 20) -> list[str]:
    if not text:
        return []
    # простой, но практичный экстрактор ссылок из plain text
    urls = re.findall(r"https?://[^\s<>()\"']+", text)
    seen: set[str] = set()
    out: list[str] = []
    for u in urls:
        u = u.strip().rstrip(".,;:!?)\"]}")
        if not u or u in seen:
            continue
        seen.add(u)
        out.append(u)
        if len(out) >= limit:
            break
    return out


def _sanitize_email_html(html: str | None) -> str | None:
    """
    Минимально безопасный рендер HTML-писем:
    - убираем script/style/noscript
    - запрещаем event handlers (on*)
    - оставляем только базовые теги
    - фильтруем href/src: только http(s), mailto, tel
    """
    if not html:
        return None
    try:
        from bs4 import BeautifulSoup
    except Exception:
        return None

    soup = BeautifulSoup(html, "lxml")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()

    allowed_tags = {
        "a",
        "p",
        "br",
        "div",
        "span",
        "strong",
        "b",
        "em",
        "i",
        "u",
        "ul",
        "ol",
        "li",
        "table",
        "thead",
        "tbody",
        "tr",
        "td",
        "th",
        "blockquote",
        "pre",
        "code",
        "hr",
        "img",
        "h1",
        "h2",
        "h3",
        "h4",
        "h5",
        "h6",
    }

    def ok_url(u: str) -> bool:
        try:
            p = urlparse(u)
            return p.scheme in {"http", "https", "mailto", "tel"}
        except Exception:
            return False

    for el in list(soup.find_all(True)):
        name = (el.name or "").lower()
        if name not in allowed_tags:
            el.unwrap()
            continue

        for attr in list(el.attrs.keys()):
            if attr.lower().startswith("on"):
                del el.attrs[attr]

        if name == "a":
            href = (el.get("href") or "").strip()
            if not href or not ok_url(href):
                el.unwrap()
            else:
                el.attrs = {"href": href, "target": "_blank", "rel": "noopener noreferrer"}
        elif name == "img":
            src = (el.get("src") or "").strip()
            if not src:
                el.decompose()
            elif src.lower().startswith("cid:"):
                # cid перепишем позже в /email/{id}/cid/{...}
                el.attrs = {"src": src, "alt": (el.get("alt") or "").strip(), "loading": "lazy"}
            elif not ok_url(src) or not src.startswith(("http://", "https://")):
                el.decompose()
            else:
                alt = (el.get("alt") or "").strip()
                el.attrs = {"src": src, "alt": alt, "loading": "lazy"}
        else:
            el.attrs = {}

    out = str(soup.body) if soup.body else str(soup)
    out = out.replace("<body>", "").replace("</body>", "").strip()
    return out or None


def _rewrite_cid_images(safe_html: str | None, email_id: int, cid_to_url: dict[str, str]) -> str | None:
    if not safe_html or not cid_to_url:
        return safe_html
    try:
        from bs4 import BeautifulSoup
    except Exception:
        return safe_html
    soup = BeautifulSoup(safe_html, "lxml")
    for img in soup.find_all("img"):
        src = (img.get("src") or "").strip()
        if not src.lower().startswith("cid:"):
            continue
        cid = src[4:].strip()
        cid = cid.strip("<>").strip()
        repl = cid_to_url.get(cid)
        if repl:
            img["src"] = repl
        else:
            # если не нашли вложение — уберём картинку
            img.decompose()
    out = str(soup.body) if soup.body else str(soup)
    out = out.replace("<body>", "").replace("</body>", "").strip()
    return out or safe_html


@app.on_event("startup")
def _startup() -> None:
    import threading

    log.warning("startup: create_all begin")
    Base.metadata.create_all(bind=engine)
    log.warning("startup: create_all ok")
    log.warning("startup: ensure_schema begin")
    done: list[bool] = []

    def _run_schema():
        try:
            ensure_schema()
        except Exception:
            pass
        done.append(True)

    t = threading.Thread(target=_run_schema, daemon=True)
    t.start()
    t.join(8.0)
    log.warning("startup: ensure_schema ok")
    def _run_quick(name: str, fn, join_s: float) -> None:
        log.warning(f"startup: {name} begin")

        def _wrap():
            try:
                fn()
            except Exception:
                pass

        tt = threading.Thread(target=_wrap, daemon=True)
        tt.start()
        tt.join(join_s)
        log.warning(f"startup: {name} ok")

    # Ящики добавляются через UI (/settings). Бутстрап из .env больше не используем.

    # Боевой авто-синк: раз в 5 минут запускаем sync по включенным ящикам.
    def _autosync_loop() -> None:
        import time

        from app.queue import get_redis

        while True:
            try:
                r = get_redis()
                # Redis-lock, чтобы не плодить параллельные синки
                if r.set("autosync:lock", "1", nx=True, ex=240):
                    q = get_queue()
                    with session_scope() as s:
                        ids = list(s.scalars(select(Mailbox.id).where(Mailbox.is_enabled == True)))  # noqa: E712
                        mbs = list(s.scalars(select(Mailbox).where(Mailbox.id.in_(ids))))
                    for mb in mbs:
                        if mb.provider == "imap":
                            q.enqueue(sync_imap_mailbox, mb.id)
                        elif mb.provider == "gmail":
                            q.enqueue(sync_gmail_mailbox, mb.id)
            except Exception:
                pass
            time.sleep(300)

    threading.Thread(target=_autosync_loop, daemon=True).start()


@app.get("/health")
def health() -> dict:
    return {"ok": True}


@app.get("/", response_class=HTMLResponse)
def index(
    request: Request,
    category: str | None = None,
    mailbox_id: str | None = None,
    q: str | None = None,
    unread: str | None = None,
    quick: str | None = None,
    sort: str | None = None,
    thread: str | None = None,
    subj: str | None = None,
    view: str | None = None,
    page: str | None = None,
    per: str | None = None,
    fragment: str | None = None,
) -> HTMLResponse:
    with session_scope() as s:
        mailboxes = list(s.scalars(select(Mailbox).order_by(Mailbox.id.asc())))
        mailbox_map = {m.id: m for m in mailboxes}
        # Цвета ящиков для UI (табы + полоска слева).
        palette = [
            "#1f77b4",  # blue
            "#ff7f0e",  # orange
            "#2ca02c",  # green
            "#d62728",  # red
            "#9467bd",  # purple
            "#8c564b",  # brown
            "#e377c2",  # pink
            "#17becf",  # cyan
            "#bcbd22",  # olive
            "#7f7f7f",  # gray
        ]
        mailbox_colors = {m.id: palette[(m.id - 1) % len(palette)] for m in mailboxes}
        last_sync_at = None
        for m in mailboxes:
            if m.last_sync_at and (last_sync_at is None or m.last_sync_at > last_sync_at):
                last_sync_at = m.last_sync_at

        view_v = (view or "").strip().lower()
        archived_view = view_v in {"archive", "archived"}

        base_qry = select(EmailMessage).where(EmailMessage.is_archived == (True if archived_view else False))  # noqa: E712
        qry = select(EmailMessage).where(EmailMessage.is_archived == (True if archived_view else False))  # noqa: E712
        if thread and thread.strip():
            t = thread.strip()
            base_qry = base_qry.where(EmailMessage.thread_id == t)
            qry = qry.where(EmailMessage.thread_id == t)
        if subj and subj.strip():
            # MVP: простая фильтрация по теме
            term_s = f"%{subj.strip()}%"
            base_qry = base_qry.where(EmailMessage.subject.ilike(term_s))
            qry = qry.where(EmailMessage.subject.ilike(term_s))
        if category:
            base_qry = base_qry.where(EmailMessage.category == category)
            qry = qry.where(EmailMessage.category == category)
        mailbox_id_int: int | None = None
        if mailbox_id and mailbox_id.strip().isdigit():
            mailbox_id_int = int(mailbox_id.strip())
            base_qry = base_qry.where(EmailMessage.mailbox_id == mailbox_id_int)
            qry = qry.where(EmailMessage.mailbox_id == mailbox_id_int)
        unread_only = bool(unread and unread.strip() in {"1", "true", "on", "yes"})
        if unread_only:
            base_qry = base_qry.where(EmailMessage.is_read == False)  # noqa: E712
            qry = qry.where(EmailMessage.is_read == False)  # noqa: E712
        if quick:
            qv = quick.strip().lower()
            if qv == "important":
                base_qry = base_qry.where(EmailMessage.category == "important")
                qry = qry.where(EmailMessage.category == "important")
            elif qv == "unread":
                base_qry = base_qry.where(EmailMessage.is_read == False)  # noqa: E712
                qry = qry.where(EmailMessage.is_read == False)  # noqa: E712
            elif qv == "today":
                now_local = dt.datetime.now().astimezone()
                start_local = now_local.replace(hour=0, minute=0, second=0, microsecond=0)
                start_utc = start_local.astimezone(dt.UTC)
                base_qry = base_qry.where(EmailMessage.date >= start_utc)
                qry = qry.where(EmailMessage.date >= start_utc)
            elif qv == "week":
                start_utc = dt.datetime.now(dt.UTC) - dt.timedelta(days=7)
                base_qry = base_qry.where(EmailMessage.date >= start_utc)
                qry = qry.where(EmailMessage.date >= start_utc)
        if q and q.strip():
            term = f"%{q.strip()}%"
            base_qry = base_qry.where(
                or_(
                    EmailMessage.subject.ilike(term),
                    EmailMessage.from_email.ilike(term),
                    EmailMessage.snippet.ilike(term),
                )
            )
            qry = qry.where(
                or_(
                    EmailMessage.subject.ilike(term),
                    EmailMessage.from_email.ilike(term),
                    EmailMessage.snippet.ilike(term),
                )
            )

        sort_v = (sort or "date").strip().lower()

        # Пагинация
        try:
            page_i = int((page or "").strip() or "1")
        except Exception:
            page_i = 1
        if page_i < 1:
            page_i = 1
        try:
            per_i = int((per or "").strip() or "50")
        except Exception:
            per_i = 50
        # разумные рамки, чтобы не убить БД/рендер
        if per_i < 10:
            per_i = 10
        if per_i > 200:
            per_i = 200
        offset_i = (page_i - 1) * per_i

        total = int(s.scalar(qry.with_only_columns(func.count()).order_by(None)) or 0)

        if sort_v == "score":
            qry = qry.order_by(
                EmailMessage.score.desc().nullslast(),
                EmailMessage.date.desc().nullslast(),
                EmailMessage.id.desc(),
            )
        else:
            sort_v = "date"
            qry = qry.order_by(EmailMessage.date.desc().nullslast(), EmailMessage.id.desc())

        qry = qry.offset(offset_i).limit(per_i)
        emails = list(s.scalars(qry))
        # "Сегодня" по локальному дню пользователя (на сервере в локальной TZ).
        now_local = dt.datetime.now().astimezone()
        start_local = now_local.replace(hour=0, minute=0, second=0, microsecond=0)
        start_utc = start_local.astimezone(dt.UTC)

        # Сводка дня (MVP, без AI): считаем только по входящим (не архив).
        summary_base = select(EmailMessage).where(EmailMessage.is_archived == False)  # noqa: E712
        summary_base = summary_base.where(EmailMessage.date >= start_utc)

        day_total = int(s.scalar(summary_base.with_only_columns(func.count()).order_by(None)) or 0)
        day_new = int(
            s.scalar(
                summary_base.where(EmailMessage.is_read == False).with_only_columns(func.count()).order_by(None)  # noqa: E712
            )
            or 0
        )
        day_important = int(
            s.scalar(summary_base.where(EmailMessage.category == "important").with_only_columns(func.count()).order_by(None)) or 0
        )
        day_newsletters = int(
            s.scalar(summary_base.where(EmailMessage.category == "newsletter").with_only_columns(func.count()).order_by(None)) or 0
        )

        # Нужен срез сегодняшних писем для кластеризации тем.
        # Лимитируем, чтобы не нагружать страницу при «днях с сотнями писем».
        today_emails = list(
            s.scalars(
                summary_base.order_by(EmailMessage.score.desc().nullslast(), EmailMessage.date.desc().nullslast(), EmailMessage.id.desc()).limit(400)
            )
        )

        def _sender_short(v: str | None) -> str:
            t = (v or "").strip()
            if not t:
                return ""
            m = re.search(r"<([^>]+)>", t)
            if m:
                return (m.group(1) or "").strip().lower()
            return t.lower()[:120]

        def _norm_subj_local(v: str | None) -> str:
            subj = (v or "").strip()
            subj = re.sub(r"^\s*(re|fw|fwd)\s*:\s*", "", subj, flags=re.IGNORECASE)
            subj = re.sub(r"\s+", " ", subj).strip().lower()
            return subj[:120]

        def _cluster_key(e: EmailMessage) -> tuple[str, str]:
            if e.thread_id:
                return ("thread", e.thread_id)
            return ("subject", _norm_subj_local(e.subject) or "(без темы)")

        clusters: dict[tuple[str, str], dict] = {}
        for e in today_emails:
            k = _cluster_key(e)
            c = clusters.get(k)
            subj_norm = k[1]
            if not c:
                clusters[k] = {
                    "kind": k[0],
                    "key": subj_norm,
                    "count": 1,
                    "unread": (0 if e.is_read else 1),
                    "important": (1 if e.category == "important" else 0),
                    "newsletter": (1 if e.category == "newsletter" else 0),
                    "top_subject": (e.subject or "(без темы)"),
                    "top_email_id": e.id,
                    "top_score": (e.score or 0),
                    "top_hint": (e.ai_explanation or e.summary or e.snippet or "").strip()[:240],
                    "senders": [s for s in [_sender_short(e.from_email)] if s],
                    "mailbox_ids": [e.mailbox_id] if e.mailbox_id else [],
                }
            else:
                c["count"] += 1
                if not e.is_read:
                    c["unread"] += 1
                if e.category == "important":
                    c["important"] += 1
                if e.category == "newsletter":
                    c["newsletter"] += 1
                sshort = _sender_short(e.from_email)
                if sshort and sshort not in c["senders"] and len(c["senders"]) < 3:
                    c["senders"].append(sshort)
                mbid = e.mailbox_id
                if mbid and mbid not in c.get("mailbox_ids", []) and len(c.get("mailbox_ids", [])) < 4:
                    c.setdefault("mailbox_ids", []).append(mbid)
                # top: хотим «самое полезное» письмо кластера
                score = (e.score or 0)
                cur = (c.get("top_score") or 0, 1 if c.get("unread") else 0, c.get("top_email_id") or 0)
                cand = (score, (1 if not e.is_read else 0), e.id)
                if cand > cur:
                    c["top_subject"] = (e.subject or "(без темы)")
                    c["top_email_id"] = e.id
                    c["top_score"] = score
                    c["top_hint"] = (e.ai_explanation or e.summary or e.snippet or "").strip()[:240]

        # Топ кластеров: сначала те, где есть непрочитанное и важное, потом по количеству.
        clusters_list = sorted(
            clusters.values(),
            key=lambda x: (
                -(1 if (x["unread"] > 0 and x["important"] > 0) else 0),
                -(1 if x["unread"] > 0 else 0),
                -(x["important"]),
                -(x["count"]),
                -(x.get("top_score") or 0),
            ),
        )[:8]

        # Что требует внимания (простые эвристики).
        urgent_markers = [
            "код",
            "подтверж",
            "вход",
            "парол",
            "security",
            "login",
            "otp",
            "2fa",
            "invoice",
            "счет",
            "оплат",
        ]
        unread_important = [e for e in today_emails if (not e.is_read) and e.category == "important"]
        unread_security = [
            e
            for e in today_emails
            if (not e.is_read)
            and any(m in f"{(e.subject or '')}\n{(e.snippet or '')}".lower() for m in urgent_markers)
        ]
        unread_high = [e for e in today_emails if (not e.is_read) and (e.score or 0) >= 70]

        attention: list[dict] = []
        if unread_important:
            attention.append(
                {
                    "title": f"Непрочитанные важные: {len(unread_important)}",
                    "hint": "Рекомендуется просмотреть в первую очередь.",
                    "href": str(request.url.include_query_params(quick="today", category="important", unread="1", page="1")),
                }
            )
        if unread_security:
            attention.append(
                {
                    "title": f"Коды/доступы/безопасность: {len(unread_security)}",
                    "hint": "Похоже на подтверждения входа/коды/важные действия.",
                    "href": str(request.url.include_query_params(quick="today", unread="1", page="1")),
                }
            )
        if unread_high and not unread_important:
            attention.append(
                {
                    "title": f"Высокая важность (score ≥ 70): {len(unread_high)}",
                    "hint": "Стоит проверить — высокий приоритет по оценке.",
                    "href": str(request.url.include_query_params(quick="today", unread="1", sort="score", page="1")),
                }
            )

        day_summary = {
            "day_total": day_total,
            "day_new": day_new,
            "day_important": day_important,
            "day_newsletters": day_newsletters,
            "clusters": clusters_list,
            "attention": attention[:4],
            "start_utc_iso": start_utc.isoformat(),
        }

        # Дайджест "Сегодня важное" (группируем по треду/теме)
        today_important = list(
            s.scalars(
                base_qry.where(
                    EmailMessage.ai_done == True,  # noqa: E712
                    EmailMessage.category == "important",
                    EmailMessage.date >= start_utc,
                )
                .order_by(
                    EmailMessage.score.desc().nullslast(),
                    EmailMessage.date.desc().nullslast(),
                    EmailMessage.id.desc(),
                )
                .limit(30)
            )
        )

    cat_ru = {
        "important": "Важно",
        "normal": "Обычное",
        "newsletter": "Рассылка",
        "spam_candidate": "Спам?",
    }

    def _u(**kwargs: str) -> str:
        # Удобно строить ссылки с сохранением текущих параметров.
        return str(request.url.include_query_params(**kwargs))

    # Пагинация: ссылки + мета
    pages_total = max(1, (total + per_i - 1) // per_i) if per_i else 1
    if page_i > pages_total:
        page_i = pages_total
    has_prev = page_i > 1
    has_next = page_i < pages_total
    prev_href = _u(page=str(page_i - 1)) if has_prev else ""
    next_href = _u(page=str(page_i + 1)) if has_next else ""

    view_tabs = [
        {"id": "", "name": "Входящие", "href": _u(view="", page="1")},
        {"id": "archive", "name": "Архив", "href": _u(view="archive", page="1")},
    ]

    mailbox_tabs = [{"id": "", "name": "Все", "color": "#64748b", "href": _u(mailbox_id="", page="1")}]
    for m in mailboxes:
        mailbox_tabs.append(
            {
                "id": str(m.id),
                "name": m.name,
                "color": mailbox_colors.get(m.id, "#64748b"),
                "href": _u(mailbox_id=str(m.id), page="1"),
            }
        )

    category_tabs = [
        {"id": "", "name": "Все", "href": _u(category="", page="1")},
        {"id": "important", "name": cat_ru["important"], "href": _u(category="important", page="1")},
        {"id": "normal", "name": cat_ru["normal"], "href": _u(category="normal", page="1")},
        {"id": "newsletter", "name": cat_ru["newsletter"], "href": _u(category="newsletter", page="1")},
        {"id": "spam_candidate", "name": cat_ru["spam_candidate"], "href": _u(category="spam_candidate", page="1")},
    ]

    sort_tabs = [
        {"id": "date", "name": "По дате", "href": _u(sort="date", page="1")},
        {"id": "score", "name": "По важности", "href": _u(sort="score", page="1")},
    ]

    quick_links = {
        "unread": _u(quick="unread", page="1"),
        "important": _u(quick="important", page="1"),
        "today": _u(quick="today", page="1"),
        "week": _u(quick="week", page="1"),
        "clear_quick": _u(quick="", page="1"),
    }

    def _norm_subj(v: str | None) -> str:
        s = (v or "").strip()
        s = re.sub(r"^\s*(re|fw|fwd)\s*:\s*", "", s, flags=re.IGNORECASE)
        s = re.sub(r"\s+", " ", s).strip().lower()
        return s[:120]

    def _group_key(e: EmailMessage) -> tuple[str, str]:
        if e.thread_id:
            return ("thread", e.thread_id)
        return ("subject", _norm_subj(e.subject) or "(без темы)")

    groups: dict[tuple[str, str], dict] = {}
    for e in today_important:
        k = _group_key(e)
        g = groups.get(k)
        if not g:
            groups[k] = {"kind": k[0], "key": k[1], "count": 1, "unread": (0 if e.is_read else 1), "top": e}
        else:
            g["count"] += 1
            if not e.is_read:
                g["unread"] += 1
            # top уже отсортирован запросом, но на всякий случай оставим максимальный score/id
            top = g["top"]
            if (e.score or 0, e.id) > (top.score or 0, top.id):
                g["top"] = e

    # "Сегодня важное" оставляем как actionable-блок: только группы, где есть непрочитанное.
    today_groups = [g for g in groups.values() if (g.get("unread") or 0) > 0]
    today_groups = sorted(today_groups, key=lambda x: (-(x["unread"]), -(x["count"]), -(x["top"].score or 0), -(x["top"].id)))

    if (fragment or "").strip().lower() in {"emails", "email_list", "list"}:
        return templates.TemplateResponse(
            request=request,
            name="partials/email_list.html",
            context={
                "emails": emails,
                "page": page_i,
                "per": per_i,
                "total": total,
                "pages_total": pages_total,
                "has_prev": has_prev,
                "has_next": has_next,
                "prev_href": prev_href,
                "next_href": next_href,
                "category": category or "",
                "cat_ru": cat_ru,
                "mailboxes": mailboxes,
                "mailbox_map": mailbox_map,
                "mailbox_colors": mailbox_colors,
                "mailbox_id": mailbox_id_int or "",
                "view": view_v,
                "archived_view": archived_view,
            },
        )

    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={
            "emails": emails,
            "page": page_i,
            "per": per_i,
            "total": total,
            "pages_total": pages_total,
            "has_prev": has_prev,
            "has_next": has_next,
            "prev_href": prev_href,
            "next_href": next_href,
            "today_groups": today_groups[:8],
            "day_summary": day_summary,
            "category": category or "",
            "cat_ru": cat_ru,
            "mailboxes": mailboxes,
            "mailbox_map": mailbox_map,
            "mailbox_colors": mailbox_colors,
            "mailbox_id": mailbox_id_int or "",
            "mailbox_tabs": mailbox_tabs,
            "view": view_v,
            "archived_view": archived_view,
            "view_tabs": view_tabs,
            "category_tabs": category_tabs,
            "sort_tabs": sort_tabs,
            "search_q": q or "",
            "unread_only": unread_only,
            "quick": (quick or "").strip().lower(),
            "sort": sort_v,
            "quick_links": quick_links,
            "thread": (thread or "").strip(),
            "subj": (subj or "").strip(),
            "last_sync_at": last_sync_at,
        },
    )


@app.post("/api/today-group")
def api_today_group(
    kind: str = Form(...),
    key: str = Form(...),
    action: str = Form(...),
) -> dict:
    """
    Быстрые действия для блока "Сегодня важное".
    kind: thread|subject
    action: mark_read|archive
    """
    kind = (kind or "").strip().lower()
    key = (key or "").strip()
    action = (action or "").strip().lower()
    if kind not in {"thread", "subject"}:
        return {"ok": False, "error": "invalid_kind"}
    if action not in {"mark_read", "archive"}:
        return {"ok": False, "error": "invalid_action"}

    now_local = dt.datetime.now().astimezone()
    start_local = now_local.replace(hour=0, minute=0, second=0, microsecond=0)
    start_utc = start_local.astimezone(dt.UTC)

    with session_scope() as s:
        q = (
            select(EmailMessage.id)
            .where(
                EmailMessage.is_archived == False,  # noqa: E712
                EmailMessage.ai_done == True,  # noqa: E712
                EmailMessage.category == "important",
                EmailMessage.date >= start_utc,
            )
            .order_by(EmailMessage.date.desc().nullslast(), EmailMessage.id.desc())
            .limit(80)
        )
        if kind == "thread":
            q = q.where(EmailMessage.thread_id == key)
        else:
            # subject: сравниваем по нормализованной теме (упрощённо)
            # NB: это MVP — точность достаточная, чтобы "пачкой" закрывать похожие письма.
            term = f"%{key.strip().lower()}%"
            q = q.where(EmailMessage.subject.ilike(term))
        ids = list(s.scalars(q))
        if not ids:
            return {"ok": True, "updated": [], "archived": [], "action": action}
        if action == "mark_read":
            s.execute(update(EmailMessage).where(EmailMessage.id.in_(ids)).values(is_read=True))
            q = get_queue()
            for eid in ids:
                q.enqueue(sync_remote_mark_read, eid)
            return {"ok": True, "updated": ids, "archived": [], "action": "mark_read"}
        s.execute(update(EmailMessage).where(EmailMessage.id.in_(ids)).values(is_archived=True))
        return {"ok": True, "updated": [], "archived": ids, "action": "archive"}


@app.post("/api/day-summary")
def api_day_summary_action(
    action: str = Form(...),
    mailbox_id: str | None = Form(default=None),
) -> dict:
    """
    Массовые действия из блока "Сводка дня" (верх /).
    action: mark_read_today_new | archive_today_newsletters
    mailbox_id: опционально, чтобы действовать в рамках выбранного ящика
    """
    action = (action or "").strip().lower()
    if action not in {"mark_read_today_new", "archive_today_newsletters"}:
        return {"ok": False, "error": "invalid_action"}
    mailbox_id_int: int | None = None
    if mailbox_id and mailbox_id.strip().isdigit():
        mailbox_id_int = int(mailbox_id.strip())

    now_local = dt.datetime.now().astimezone()
    start_local = now_local.replace(hour=0, minute=0, second=0, microsecond=0)
    start_utc = start_local.astimezone(dt.UTC)

    base = select(EmailMessage.id).where(EmailMessage.is_archived == False, EmailMessage.date >= start_utc)  # noqa: E712
    if mailbox_id_int:
        base = base.where(EmailMessage.mailbox_id == mailbox_id_int)

    updated: list[int] = []
    archived: list[int] = []
    with session_scope() as s:
        if action == "mark_read_today_new":
            ids = list(s.scalars(base.where(EmailMessage.is_read == False)))  # noqa: E712
            if ids:
                s.execute(update(EmailMessage).where(EmailMessage.id.in_(ids)).values(is_read=True))
                updated = ids
                q = get_queue()
                for eid in ids:
                    q.enqueue(sync_remote_mark_read, eid)
        else:
            ids = list(s.scalars(base.where(EmailMessage.category == "newsletter")))
            if ids:
                s.execute(update(EmailMessage).where(EmailMessage.id.in_(ids)).values(is_archived=True))
                archived = ids

    return {"ok": True, "updated": updated, "archived": archived, "action": action}


@app.post("/api/day-summary-ai")
def api_day_summary_ai() -> dict:
    """
    AI-выжимка дня: один запрос к модели на основе кластеров сегодняшней почты.
    Ничего не применяем автоматически — только текст/предложения.
    """
    from app.ai_client import summarize_day_digest

    now_local = dt.datetime.now().astimezone()
    start_local = now_local.replace(hour=0, minute=0, second=0, microsecond=0)
    start_utc = start_local.astimezone(dt.UTC)

    with session_scope() as s:
        base = select(EmailMessage).where(EmailMessage.is_archived == False, EmailMessage.date >= start_utc)  # noqa: E712
        day_total = int(s.scalar(base.with_only_columns(func.count()).order_by(None)) or 0)
        day_new = int(s.scalar(base.where(EmailMessage.is_read == False).with_only_columns(func.count()).order_by(None)) or 0)  # noqa: E712
        day_important = int(
            s.scalar(base.where(EmailMessage.category == "important").with_only_columns(func.count()).order_by(None)) or 0
        )
        day_newsletters = int(
            s.scalar(base.where(EmailMessage.category == "newsletter").with_only_columns(func.count()).order_by(None)) or 0
        )

        today_emails = list(
            s.scalars(
                base.order_by(EmailMessage.score.desc().nullslast(), EmailMessage.date.desc().nullslast(), EmailMessage.id.desc()).limit(400)
            )
        )

    # Собираем кластеры так же, как на главной (упрощённо, без дубля логики UI).
    def _sender_short(v: str | None) -> str:
        t = (v or "").strip()
        if not t:
            return ""
        m = re.search(r"<([^>]+)>", t)
        if m:
            return (m.group(1) or "").strip().lower()
        return t.lower()[:120]

    def _norm_subj_local(v: str | None) -> str:
        subj = (v or "").strip()
        subj = re.sub(r"^\s*(re|fw|fwd)\s*:\s*", "", subj, flags=re.IGNORECASE)
        subj = re.sub(r"\s+", " ", subj).strip().lower()
        return subj[:120]

    def _cluster_key(e: EmailMessage) -> tuple[str, str]:
        if e.thread_id:
            return ("thread", e.thread_id)
        return ("subject", _norm_subj_local(e.subject) or "(без темы)")

    clusters: dict[tuple[str, str], dict] = {}
    for e in today_emails:
        k = _cluster_key(e)
        c = clusters.get(k)
        if not c:
            clusters[k] = {
                "kind": k[0],
                "key": k[1],
                "count": 1,
                "unread": (0 if e.is_read else 1),
                "important": (1 if e.category == "important" else 0),
                "newsletter": (1 if e.category == "newsletter" else 0),
                "top_subject": (e.subject or "(без темы)"),
                "top_email_id": e.id,
                "top_score": (e.score or 0),
                "top_hint": (e.ai_explanation or e.summary or e.snippet or "").strip()[:240],
                "senders": [s for s in [_sender_short(e.from_email)] if s],
                "mailbox_ids": [e.mailbox_id] if e.mailbox_id else [],
            }
        else:
            c["count"] += 1
            if not e.is_read:
                c["unread"] += 1
            if e.category == "important":
                c["important"] += 1
            if e.category == "newsletter":
                c["newsletter"] += 1
            sshort = _sender_short(e.from_email)
            if sshort and sshort not in c["senders"] and len(c["senders"]) < 3:
                c["senders"].append(sshort)
            mbid = e.mailbox_id
            if mbid and mbid not in c.get("mailbox_ids", []) and len(c.get("mailbox_ids", [])) < 4:
                c.setdefault("mailbox_ids", []).append(mbid)
            score = (e.score or 0)
            cur = (c.get("top_score") or 0, (1 if c["unread"] > 0 else 0), c.get("top_email_id") or 0)
            cand = (score, (1 if not e.is_read else 0), e.id)
            if cand > cur:
                c["top_subject"] = (e.subject or "(без темы)")
                c["top_email_id"] = e.id
                c["top_score"] = score
                c["top_hint"] = (e.ai_explanation or e.summary or e.snippet or "").strip()[:240]

    clusters_list = sorted(
        clusters.values(),
        key=lambda x: (
            -(1 if (x["unread"] > 0 and x["important"] > 0) else 0),
            -(1 if x["unread"] > 0 else 0),
            -(x["important"]),
            -(x["count"]),
            -(x.get("top_score") or 0),
        ),
    )[:10]

    try:
        digest = summarize_day_digest(
            stats={
                "total": day_total,
                "unread": day_new,
                "important": day_important,
                "newsletters": day_newsletters,
            },
            clusters=clusters_list,
        )
        return {"ok": True, "digest": digest}
    except Exception as e:
        # Не ломаем UI: отдаём fallback, который всё равно полезен.
        return {
            "ok": True,
            "digest": {
                "headline": f"Сегодня: всего {day_total}, непрочитано {day_new}, важных {day_important}, рассылок {day_newsletters}.",
                "bullets": [
                    f"Темы дня: {', '.join([(c.get('top_subject') or '').strip()[:60] for c in clusters_list[:6] if (c.get('top_subject') or '').strip()])}"
                    or "Темы дня: —"
                ],
                "actions": [
                    {"id": "mark_read_today_new", "title": "Прочитано: все непрочитанные за сегодня"},
                    {"id": "archive_today_newsletters", "title": "В архив: рассылки за сегодня"},
                ],
                "error": str(e)[:300],
            },
        }


@app.post("/actions/email/{email_id}/archive")
def action_archive_email(email_id: int) -> RedirectResponse:
    with session_scope() as s:
        msg = s.get(EmailMessage, email_id)
        if msg:
            msg.is_archived = True
            if not msg.is_read:
                msg.is_read = True
                get_queue().enqueue(sync_remote_mark_read, msg.id)
    return RedirectResponse("/", status_code=303)


@app.post("/api/email/{email_id}/archive")
def api_archive_email(email_id: int) -> dict:
    with session_scope() as s:
        msg = s.get(EmailMessage, email_id)
        if msg:
            msg.is_archived = True
            if not msg.is_read:
                msg.is_read = True
                get_queue().enqueue(sync_remote_mark_read, msg.id)
    return {"ok": True, "email_id": email_id, "archived": True}


@app.post("/actions/email/{email_id}/set-category")
def action_set_category(email_id: int, category: str) -> RedirectResponse:
    if category not in {"important", "normal", "newsletter", "spam_candidate"}:
        return RedirectResponse("/", status_code=303)
    with session_scope() as s:
        msg = s.get(EmailMessage, email_id)
        if msg:
            msg.category = category
            _learn_on_manual_category_change(s, from_email=msg.from_email, category=category)
    return RedirectResponse("/", status_code=303)


@app.post("/api/email/{email_id}/set-category")
def api_set_category(email_id: int, category: str = Form(...)) -> dict:
    if category not in {"important", "normal", "newsletter", "spam_candidate"}:
        return {"ok": False, "error": "invalid_category"}
    with session_scope() as s:
        msg = s.get(EmailMessage, email_id)
        if not msg:
            return {"ok": False, "error": "not_found"}
        msg.category = category
        _learn_on_manual_category_change(s, from_email=msg.from_email, category=category)
    return {"ok": True, "email_id": email_id, "category": category}


@app.post("/actions/email/{email_id}/mark-read")
def action_mark_read(email_id: int) -> RedirectResponse:
    with session_scope() as s:
        s.execute(update(EmailMessage).where(EmailMessage.id == email_id).values(is_read=True))
    get_queue().enqueue(sync_remote_mark_read, email_id)
    return RedirectResponse("/", status_code=303)


@app.post("/api/email/{email_id}/mark-read")
def api_mark_read(email_id: int) -> dict:
    with session_scope() as s:
        s.execute(update(EmailMessage).where(EmailMessage.id == email_id).values(is_read=True))
    get_queue().enqueue(sync_remote_mark_read, email_id)
    return {"ok": True, "email_id": email_id}


@app.post("/actions/bulk")
def action_bulk(
    ids: list[int] = Form(default=[]),
    action: str = Form(...),
    category: str | None = Form(default=None),
) -> RedirectResponse:
    if not ids:
        return RedirectResponse("/", status_code=303)
    with session_scope() as s:
        if action == "mark_read":
            s.execute(update(EmailMessage).where(EmailMessage.id.in_(ids)).values(is_read=True))
            q = get_queue()
            for eid in ids:
                q.enqueue(sync_remote_mark_read, eid)
        elif action == "archive":
            s.execute(update(EmailMessage).where(EmailMessage.id.in_(ids)).values(is_archived=True))
        elif action == "unarchive":
            s.execute(update(EmailMessage).where(EmailMessage.id.in_(ids)).values(is_archived=False))
        elif action == "set_category":
            if category in {"important", "normal", "newsletter", "spam_candidate"}:
                msgs = list(s.scalars(select(EmailMessage).where(EmailMessage.id.in_(ids))))
                for msg in msgs:
                    msg.category = category
                    _learn_on_manual_category_change(s, from_email=msg.from_email, category=category)
    return RedirectResponse("/", status_code=303)


@app.post("/api/bulk")
def api_bulk(
    ids: list[int] = Form(default=[]),
    action: str = Form(...),
    category: str | None = Form(default=None),
) -> dict:
    if not ids:
        return {"ok": True, "updated": [], "archived": []}
    updated: list[int] = []
    archived: list[int] = []
    with session_scope() as s:
        if action == "mark_read":
            s.execute(update(EmailMessage).where(EmailMessage.id.in_(ids)).values(is_read=True))
            updated = ids
            q = get_queue()
            for eid in ids:
                q.enqueue(sync_remote_mark_read, eid)
        elif action == "archive":
            s.execute(update(EmailMessage).where(EmailMessage.id.in_(ids)).values(is_archived=True))
            archived = ids
        elif action == "unarchive":
            s.execute(update(EmailMessage).where(EmailMessage.id.in_(ids)).values(is_archived=False))
            updated = ids
        elif action == "set_category":
            if category in {"important", "normal", "newsletter", "spam_candidate"}:
                msgs = list(s.scalars(select(EmailMessage).where(EmailMessage.id.in_(ids))))
                for msg in msgs:
                    msg.category = category
                    _learn_on_manual_category_change(s, from_email=msg.from_email, category=category)
                updated = [m.id for m in msgs]
            else:
                return {"ok": False, "error": "invalid_category"}
    return {"ok": True, "updated": updated, "archived": archived, "action": action, "category": category}


@app.get("/email/{email_id}", response_class=HTMLResponse)
def email_view(request: Request, email_id: int) -> HTMLResponse:
    with session_scope() as s:
        msg = s.get(EmailMessage, email_id)
        if not msg:
            return RedirectResponse("/", status_code=303)
        # Открытие письма = просмотр. Помечаем прочитанным.
        if not msg.is_read:
            msg.is_read = True
            get_queue().enqueue(sync_remote_mark_read, msg.id)
        mb = s.get(Mailbox, msg.mailbox_id)

        # prev/next в контексте фильтров (если пришли со списка)
        qp = request.query_params
        category = (qp.get("category") or "").strip() or None
        mailbox_id = (qp.get("mailbox_id") or "").strip()
        q = (qp.get("q") or "").strip() or None
        quick = (qp.get("quick") or "").strip().lower() or None
        sort = (qp.get("sort") or "date").strip().lower()

        base = select(EmailMessage.id).where(EmailMessage.is_archived == False)  # noqa: E712
        if category:
            base = base.where(EmailMessage.category == category)
        if mailbox_id.isdigit():
            base = base.where(EmailMessage.mailbox_id == int(mailbox_id))
        if quick:
            if quick == "important":
                base = base.where(EmailMessage.category == "important")
            elif quick == "unread":
                base = base.where(EmailMessage.is_read == False)  # noqa: E712
            elif quick == "today":
                now_local = dt.datetime.now().astimezone()
                start_local = now_local.replace(hour=0, minute=0, second=0, microsecond=0)
                start_utc = start_local.astimezone(dt.UTC)
                base = base.where(EmailMessage.date >= start_utc)
            elif quick == "week":
                start_utc = dt.datetime.now(dt.UTC) - dt.timedelta(days=7)
                base = base.where(EmailMessage.date >= start_utc)
        if q:
            term = f"%{q}%"
            base = base.where(
                or_(
                    EmailMessage.subject.ilike(term),
                    EmailMessage.from_email.ilike(term),
                    EmailMessage.snippet.ilike(term),
                )
            )

        if sort == "score":
            base = base.order_by(EmailMessage.score.desc().nullslast(), EmailMessage.date.desc().nullslast(), EmailMessage.id.desc())
        else:
            base = base.order_by(EmailMessage.date.desc().nullslast(), EmailMessage.id.desc())

        ids = list(s.scalars(base.limit(250)))
        prev_id = None
        next_id = None
        try:
            idx = ids.index(email_id)
            prev_id = ids[idx - 1] if idx - 1 >= 0 else None
            next_id = ids[idx + 1] if idx + 1 < len(ids) else None
        except ValueError:
            pass

    with session_scope() as s:
        atts = list(s.scalars(select(EmailAttachment).where(EmailAttachment.email_id == email_id).order_by(EmailAttachment.id.asc())))
    cid_map = {}
    for a in atts:
        if a.content_id:
            cid_map[a.content_id.strip("<>").strip()] = f"/email/{email_id}/cid/{a.content_id.strip('<>').strip()}"

    safe_html = _sanitize_email_html(msg.body_html)
    safe_html = _rewrite_cid_images(safe_html, email_id, cid_map)

    return templates.TemplateResponse(
        request=request,
        name="email.html",
        context={
            "e": msg,
            "mb": mb,
            "prev_id": prev_id,
            "next_id": next_id,
            "safe_html": safe_html,
            "attachments": atts,
            "back_url": "/" + (("?" + str(request.url.query)) if request.url.query else ""),
        },
    )


@app.get("/email/{email_id}/cid/{cid}")
def email_cid(email_id: int, cid: str):
    c = (cid or "").strip().strip("<>").strip()
    with session_scope() as s:
        att = s.scalars(
            select(EmailAttachment).where(
                EmailAttachment.email_id == email_id,
                EmailAttachment.content_id == c,
            )
        ).first()
        if not att:
            return RedirectResponse(f"/email/{email_id}", status_code=302)
    return RedirectResponse(f"/attachments/{att.id}", status_code=302)


@app.get("/attachments/{attachment_id}")
def download_attachment(attachment_id: int):
    from fastapi.responses import Response
    from urllib.parse import quote

    with session_scope() as s:
        att = s.get(EmailAttachment, attachment_id)
        if not att or not att.data:
            return Response(status_code=404, content=b"")
        ct = att.content_type or "application/octet-stream"
        filename = att.filename or f"attachment-{att.id}"
        dispo = "inline" if (att.is_inline and ct.startswith("image/")) else "attachment"
        headers = {"Content-Disposition": f"{dispo}; filename*=UTF-8''{quote(filename)}"}
        return Response(content=att.data, media_type=ct, headers=headers)


@app.get("/settings", response_class=HTMLResponse)
def settings_page(request: Request) -> HTMLResponse:
    with session_scope() as s:
        mailboxes = list(s.scalars(select(Mailbox).order_by(Mailbox.id.asc())))
        rules = {r.key: (r.value or "") for r in list(s.scalars(select(AppSetting)))}
    q = get_queue()
    run = get_ai_run_status() or AiRunStatus(running=False, message="Ещё не запускалось")

    def _fmt_iso(s: str) -> str:
        if not s:
            return ""
        try:
            # Python понимает ISO с "+00:00"
            d = dt.datetime.fromisoformat(s.replace("Z", "+00:00"))
            return d.astimezone().strftime("%d.%m.%Y %H:%M:%S")
        except Exception:
            return s

    started_registry = StartedJobRegistry(q.name, connection=q.connection)
    scheduled_registry = ScheduledJobRegistry(q.name, connection=q.connection)
    deferred_registry = DeferredJobRegistry(q.name, connection=q.connection)

    queue_pending = q.count + len(scheduled_registry.get_job_ids()) + len(deferred_registry.get_job_ids())
    queue_started = len(started_registry.get_job_ids())

    # Человеческий статус без дублей и с детектором "зависшего" состояния.
    status_label = "Не выполняется"
    if run and run.message == "В очереди":
        status_label = "В очереди"
    elif run and run.running:
        status_label = "В процессе"

    status_detail = ""
    if run and run.message and run.message not in {"", status_label, "В очереди"}:
        status_detail = run.message

    # Если Redis-статус говорит "running", но в очередях ничего не выполняется,
    # значит job мог завершиться/упасть/быть очищен — показываем предупреждение.
    status_mismatch = bool(run and run.running and (queue_pending + queue_started == 0))

    qp = request.query_params
    ui_notice = (qp.get("notice") or "").strip()
    ui_error = (qp.get("error") or "").strip()

    return templates.TemplateResponse(
        request=request,
        name="settings.html",
        context={
            "mailboxes": mailboxes,
            "ai_test": get_ai_test_result(),
            "ai_test_status": get_ai_test_status(),
            "ai_run": run,
            "ai_run_status_label": status_label,
            "ai_run_status_detail": status_detail,
            "ai_run_status_mismatch": status_mismatch,
            "ai_run_started_h": _fmt_iso(run.started_at),
            "ai_run_finished_h": _fmt_iso(run.finished_at),
            "queue_pending": queue_pending,
            "queue_started": queue_started,
            "rules": rules,
            "ai_base_url": settings.ai_base_url,
            "ai_model": settings.ai_model,
            "ui_notice": ui_notice,
            "ui_error": ui_error,
        },
    )


@app.get("/api/ai-run-status")
def api_ai_run_status() -> dict:
    """
    Живой монитор AI-обработки для UI (/settings): статус из Redis + состояние очередей RQ.
    Также "разлипает" running-статус, если job исчезла, а heartbeat давно не обновлялся.
    """
    q = get_queue()
    run = get_ai_run_status() or AiRunStatus(running=False, message="Ещё не запускалось")

    started_registry = StartedJobRegistry(q.name, connection=q.connection)
    scheduled_registry = ScheduledJobRegistry(q.name, connection=q.connection)
    deferred_registry = DeferredJobRegistry(q.name, connection=q.connection)
    queue_pending = q.count + len(scheduled_registry.get_job_ids()) + len(deferred_registry.get_job_ids())
    queue_started = len(started_registry.get_job_ids())

    def _parse_iso(s: str) -> dt.datetime | None:
        if not s:
            return None
        try:
            return dt.datetime.fromisoformat(s.replace("Z", "+00:00"))
        except Exception:
            return None

    now = dt.datetime.now(dt.UTC)
    started_dt = _parse_iso(run.started_at)
    updated_dt = _parse_iso(getattr(run, "updated_at", "") or "")

    status_label = "Не выполняется"
    if run and run.message == "В очереди":
        status_label = "В очереди"
    elif run and run.running:
        status_label = "В процессе"

    status_detail = ""
    if run and run.message and run.message not in {"", status_label, "В очереди"}:
        status_detail = run.message

    status_mismatch = bool(run and run.running and (queue_pending + queue_started == 0))

    # Если job "running", но очереди пустые и heartbeat не обновлялся давно — считаем, что job умерла.
    # Это убирает зависания в UI без ручного refresh.
    if status_mismatch:
        last_hb = updated_dt or started_dt
        if last_hb and (now - last_hb).total_seconds() >= 30:
            run = AiRunStatus(
                running=False,
                started_at=run.started_at,
                updated_at=now_iso(),
                finished_at=now_iso(),
                total=run.total,
                processed=run.processed,
                ok=run.ok,
                failed=run.failed,
                message="Завершено (монитор)",
            )
            try:
                set_ai_run_status(run)
            except Exception:
                pass
            status_label = "Не выполняется"
            status_detail = run.message
            status_mismatch = False

    return {
        "ok": True,
        "run": {
            "running": bool(run.running),
            "started_at": run.started_at or "",
            "updated_at": getattr(run, "updated_at", "") or "",
            "finished_at": run.finished_at or "",
            "total": int(run.total or 0),
            "processed": int(run.processed or 0),
            "ok": int(run.ok or 0),
            "failed": int(run.failed or 0),
            "message": run.message or "",
        },
        "ui": {
            "status_label": status_label,
            "status_detail": status_detail,
            "status_mismatch": bool(status_mismatch),
            "queue_pending": int(queue_pending),
            "queue_started": int(queue_started),
        },
    }


@app.post("/actions/mailbox/add-imap")
def action_mailbox_add_imap(
    name: str = Form(...),
    host: str = Form(...),
    port: int = Form(default=993),
    tls_verify: str | None = Form(default=None),
    username: str = Form(...),
    password: str = Form(...),
    folder: str = Form(default="INBOX"),
) -> RedirectResponse:
    name = (name or "").strip()
    host = (host or "").strip()
    username = (username or "").strip()
    password = (password or "").strip()
    folder = (folder or "INBOX").strip() or "INBOX"
    try:
        port_i = int(port)
    except Exception:
        port_i = 993
    if not (name and host and username and password):
        return RedirectResponse("/settings?error=fill_required", status_code=303)
    if port_i <= 0 or port_i > 65535:
        return RedirectResponse("/settings?error=bad_port", status_code=303)

    with session_scope() as s:
        mb = Mailbox(provider="imap", name=name, is_enabled=True)
        mb.imap_host_enc = encrypt_str(host)
        mb.imap_port = port_i
        mb.imap_user_enc = encrypt_str(username)
        mb.imap_password_enc = encrypt_str(password)
        mb.imap_folder = folder
        mb.imap_last_uid = None
        mb.imap_tls_verify = bool(tls_verify)
        s.add(mb)
        s.flush()
        new_id = mb.id

    # сразу поставим в очередь синк, чтобы пользователь видел результат
    get_queue().enqueue(sync_imap_mailbox, new_id)

    return RedirectResponse("/settings?notice=imap_added", status_code=303)


@app.post("/actions/mailbox/add-gmail")
def action_mailbox_add_gmail() -> RedirectResponse:
    """
    Новый флоу: Gmail ящик создаётся в БД, затем запускается OAuth для конкретного mailbox_id.
    """
    with session_scope() as s:
        # Не создаём бесконечные "пустые" Gmail ящики при повторных кликах.
        # Переиспользуем незавершённый mailbox без токена/почты.
        mb = (
            s.scalars(
                select(Mailbox)
                .where(
                    Mailbox.provider == "gmail",
                    Mailbox.gmail_credentials_enc.is_(None),
                    Mailbox.gmail_email.is_(None),
                )
                .order_by(Mailbox.id.desc())
            ).first()
        )
        if not mb:
            mb = Mailbox(provider="gmail", name="Gmail", is_enabled=True)
            s.add(mb)
            s.flush()
        mailbox_id = mb.id
    return RedirectResponse(f"/connect/gmail?mailbox_id={mailbox_id}", status_code=303)


@app.post("/actions/rules-save")
def rules_save(
    important_threshold: str = Form(default="70"),
    sender_whitelist: str = Form(default=""),
    sender_blacklist: str = Form(default=""),
) -> RedirectResponse:
    with session_scope() as s:
        def _upsert(k: str, v: str) -> None:
            cur = s.get(AppSetting, k)
            if not cur:
                cur = AppSetting(key=k, value=v)
                s.add(cur)
            else:
                cur.value = v
                cur.updated_at = dt.datetime.now(dt.UTC)

        _upsert("important_threshold", (important_threshold or "").strip())
        _upsert("sender_whitelist", (sender_whitelist or "").strip())
        _upsert("sender_blacklist", (sender_blacklist or "").strip())

    return RedirectResponse("/settings", status_code=303)


@app.post("/actions/recompute")
def action_recompute() -> RedirectResponse:
    q = get_queue()
    q.enqueue(recompute_all_basic)
    return RedirectResponse("/", status_code=303)


@app.post("/actions/ai-process")
def action_ai_process() -> RedirectResponse:
    q = get_queue()
    q.enqueue(ai_process_recent, 50)
    return RedirectResponse("/settings", status_code=303)


@app.post("/actions/ai-run")
def action_ai_run() -> RedirectResponse:
    q = get_queue()
    # Ставим статус сразу, чтобы индикатор не был 0/0,
    # даже если очередь длинная и job начнётся позже.
    set_ai_run_status(AiRunStatus(running=False, started_at=now_iso(), message="В очереди"))
    set_ai_stop_flag(False)
    q.enqueue(ai_run, 200, at_front=True)
    return RedirectResponse("/settings", status_code=303)


@app.post("/actions/ai-stop")
def action_ai_stop() -> RedirectResponse:
    q = get_queue()
    # Ставим флаг сразу и дублируем job'ой (на случай, если воркер в другом процессе и мы хотим факт в очереди).
    set_ai_stop_flag(True)
    cur = get_ai_run_status()
    if cur and cur.running:
        set_ai_run_status(
            AiRunStatus(
                running=True,
                started_at=cur.started_at,
                total=cur.total,
                processed=cur.processed,
                ok=cur.ok,
                failed=cur.failed,
                message="Остановка запрошена",
            )
        )
    q.enqueue(ai_stop, at_front=True)
    return RedirectResponse("/settings", status_code=303)


@app.post("/actions/ai-reset-all")
def action_ai_reset_all() -> RedirectResponse:
    q = get_queue()
    q.enqueue(ai_reset_all, 500, at_front=True)
    return RedirectResponse("/settings", status_code=303)


@app.post("/actions/ai-reset-mailbox")
def action_ai_reset_mailbox(mailbox_id: int = Form(...)) -> RedirectResponse:
    q = get_queue()
    q.enqueue(ai_reset_mailbox, mailbox_id, 500, at_front=True)
    return RedirectResponse("/settings", status_code=303)


@app.post("/actions/ai-test")
def action_ai_test() -> RedirectResponse:
    q = get_queue()
    # Ставим статус сразу, чтобы было понятно, что тест запущен.
    set_ai_test_status(AiTestStatus(running=False, started_at=now_iso(), message="В очереди"))
    q.enqueue(ai_test_model)
    return RedirectResponse("/settings", status_code=303)


@app.post("/actions/ai-retry-failed")
def action_ai_retry_failed() -> RedirectResponse:
    q = get_queue()
    q.enqueue(ai_retry_failed, 100)
    return RedirectResponse("/settings", status_code=303)


@app.post("/actions/ai-retry-frozen")
def action_ai_retry_frozen() -> RedirectResponse:
    q = get_queue()
    q.enqueue(ai_retry_frozen_assignment_errors, 2000)
    return RedirectResponse("/settings", status_code=303)


@app.post("/actions/ai-reset-old-errors")
def action_ai_reset_old_errors() -> RedirectResponse:
    q = get_queue()
    q.enqueue(ai_reset_old_errors, 1000)
    return RedirectResponse("/settings", status_code=303)


@app.post("/actions/ai-reset-empty-explanations")
def action_ai_reset_empty_explanations() -> RedirectResponse:
    q = get_queue()
    q.enqueue(ai_reset_empty_explanations, 1000)
    return RedirectResponse("/settings", status_code=303)


@app.post("/actions/sync-all")
def action_sync_all() -> RedirectResponse:
    q = get_queue()
    with session_scope() as s:
        now = dt.datetime.now(dt.UTC)
        mailboxes = list(s.scalars(select(Mailbox).where(Mailbox.is_enabled == True)))  # noqa: E712
        # Помечаем "queued" сразу, чтобы UI видел запуск даже если воркер/очередь временно недоступны.
        for mb in mailboxes:
            mb.last_sync_status = "queued"
            mb.last_sync_error = None
            mb.last_sync_count = 0
            mb.last_sync_at = now
    for mailbox_id, provider in mailboxes:
        if provider == "gmail":
            q.enqueue(sync_gmail_mailbox, mailbox_id)
        else:
            q.enqueue(sync_imap_mailbox, mailbox_id)
    return RedirectResponse("/", status_code=303)


@app.post("/api/sync-all")
def api_sync_all() -> dict:
    """
    Запуск синка со страницы / (AJAX). Возвращает ok=true сразу, дальше синк идёт в фоне.
    """
    q = get_queue()
    with session_scope() as s:
        now = dt.datetime.now(dt.UTC)
        mailboxes = list(s.scalars(select(Mailbox).where(Mailbox.is_enabled == True)))  # noqa: E712
        for mb in mailboxes:
            mb.last_sync_status = "queued"
            mb.last_sync_error = None
            mb.last_sync_count = 0
            mb.last_sync_at = now
    for mailbox_id, provider in mailboxes:
        if provider == "gmail":
            q.enqueue(sync_gmail_mailbox, mailbox_id)
        else:
            q.enqueue(sync_imap_mailbox, mailbox_id)
    return {"ok": True}


@app.get("/api/sync-status")
def api_sync_status() -> dict:
    """
    Статус синхронизации для UI: идет ли сейчас sync и когда была последняя попытка.
    """
    syncing = False
    last_sync_at: dt.datetime | None = None
    with session_scope() as s:
        mbs = list(s.scalars(select(Mailbox).where(Mailbox.is_enabled == True)))  # noqa: E712
        for mb in mbs:
            if mb.last_sync_status in {"queued", "running"}:
                syncing = True
            if mb.last_sync_at and (last_sync_at is None or mb.last_sync_at > last_sync_at):
                last_sync_at = mb.last_sync_at
    return {
        "ok": True,
        "syncing": syncing,
        "last_sync_at": last_sync_at.isoformat() if last_sync_at else "",
    }


@app.post("/actions/sync/{mailbox_id}")
def action_sync_one(mailbox_id: int) -> RedirectResponse:
    q = get_queue()
    with session_scope() as s:
        mb = s.get(Mailbox, mailbox_id)
        if mb and mb.provider == "gmail":
            q.enqueue(sync_gmail_mailbox, mailbox_id)
        else:
            q.enqueue(sync_imap_mailbox, mailbox_id)
    return RedirectResponse("/settings", status_code=303)


@app.get("/connect/gmail")
def connect_gmail(mailbox_id: int | None = None) -> RedirectResponse:
    if not (settings.gmail_oauth_client_id and settings.gmail_oauth_client_secret):
        return RedirectResponse("/", status_code=303)

    from google_auth_oauthlib.flow import Flow

    from app.gmail_client import GMAIL_SCOPES

    client_config = {
        "web": {
            "client_id": settings.gmail_oauth_client_id,
            "client_secret": settings.gmail_oauth_client_secret,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
        }
    }

    # Привязываем OAuth к конкретному mailbox_id.
    if mailbox_id is None:
        with session_scope() as s:
            mb = Mailbox(provider="gmail", name="Gmail", is_enabled=True)
            s.add(mb)
            s.flush()
            mailbox_id = mb.id
    state = issue_state(f"gmail:{mailbox_id}")
    flow = Flow.from_client_config(client_config, scopes=GMAIL_SCOPES, redirect_uri=settings.gmail_oauth_redirect_uri)
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
        state=state,
    )
    return RedirectResponse(auth_url, status_code=302)


@app.get("/oauth2/google/callback")
def oauth2_google_callback(code: str | None = None, state: str | None = None) -> RedirectResponse:
    purpose = consume_state(state or "")
    mailbox_id: int | None = None
    if purpose and purpose.startswith("gmail:"):
        tail = purpose.split(":", 1)[1].strip()
        if tail.isdigit():
            mailbox_id = int(tail)
    if not code or not state or mailbox_id is None:
        return RedirectResponse("/", status_code=303)

    if not (settings.gmail_oauth_client_id and settings.gmail_oauth_client_secret):
        return RedirectResponse("/", status_code=303)

    from google_auth_oauthlib.flow import Flow

    from app.gmail_client import GMAIL_SCOPES

    client_config = {
        "web": {
            "client_id": settings.gmail_oauth_client_id,
            "client_secret": settings.gmail_oauth_client_secret,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
        }
    }
    flow = Flow.from_client_config(client_config, scopes=GMAIL_SCOPES, redirect_uri=settings.gmail_oauth_redirect_uri)
    fetch_err = None
    try:
        flow.fetch_token(code=code)
    except Exception as e:
        # oauthlib иногда поднимает Warning как Exception при "Scope has changed...".
        # В таком случае токена может не быть — аккуратно отработаем без 500.
        fetch_err = e

    try:
        creds = flow.credentials
    except Exception:
        # Если токена нет — не падаем 500, а возвращаемся в настройки с ошибкой.
        if fetch_err:
            log.warning(f"gmail_oauth_callback_error: {type(fetch_err).__name__}: {fetch_err}")
        return RedirectResponse("/settings?gmail_error=oauth_no_token", status_code=303)

    if not getattr(creds, "token", None):
        if fetch_err:
            log.warning(f"gmail_oauth_callback_error: {type(fetch_err).__name__}: {fetch_err}")
        return RedirectResponse("/settings?gmail_error=oauth_no_token", status_code=303)

    from app.gmail_client import build_gmail_service, get_profile

    with session_scope() as s:
        mb = s.get(Mailbox, mailbox_id)
        if not mb:
            return RedirectResponse("/settings?error=gmail_mailbox_missing", status_code=303)
        if mb.provider != "gmail":
            return RedirectResponse("/settings?error=gmail_mailbox_wrong_type", status_code=303)
        # Подтянем email для дедупликации/UI (best-effort).
        email_addr: str | None = None
        try:
            prof = get_profile(build_gmail_service(creds))
            email_addr = (prof.email_address or "").strip() or None
        except Exception:
            email_addr = None

        # Если такой Gmail уже подключён — не создаём дубль.
        if email_addr:
            existing = (
                s.scalars(
                    select(Mailbox)
                    .where(
                        Mailbox.provider == "gmail",
                        Mailbox.gmail_email == email_addr,
                        Mailbox.id != mailbox_id,
                    )
                    .order_by(Mailbox.id.asc())
                ).first()
            )
            if existing:
                # Если "текущий" mailbox пустой — можно удалить и сохранить токен в existing.
                mb2 = existing
                mb2.gmail_credentials_enc = encrypt_str(creds.to_json())
                mb2.gmail_last_history_id = None
                mb2.is_enabled = True
                mb2.name = f"Gmail: {email_addr}"
                # Удаляем промежуточный mailbox (без писем).
                try:
                    cnt = s.scalar(select(func.count()).select_from(EmailMessage).where(EmailMessage.mailbox_id == mailbox_id))
                except Exception:
                    cnt = 0
                if not cnt:
                    s.delete(mb)
                return RedirectResponse("/settings?notice=gmail_connected", status_code=303)

        mb.gmail_credentials_enc = encrypt_str(creds.to_json())
        # После переподключения/смены scope лучше сбросить historyId, чтобы инкрементальный sync не "пропускал" события.
        mb.gmail_last_history_id = None
        if email_addr:
            mb.gmail_email = email_addr
            mb.name = f"Gmail: {email_addr}"
        # Профиль (email/historyId) подтянем в sync job. На колбэке не падаем,
        # даже если Gmail API ещё не включён в проекте.
        mb.last_sync_status = "ok"
        mb.last_sync_error = None

    return RedirectResponse("/settings?notice=gmail_connected", status_code=303)


@app.post("/actions/mailbox/toggle/{mailbox_id}")
def action_mailbox_toggle(mailbox_id: int) -> RedirectResponse:
    with session_scope() as s:
        mb = s.get(Mailbox, mailbox_id)
        if not mb:
            return RedirectResponse("/settings?error=mailbox_not_found", status_code=303)
        mb.is_enabled = not bool(mb.is_enabled)
    return RedirectResponse("/settings", status_code=303)


@app.post("/actions/mailbox/tls-verify/{mailbox_id}")
def action_mailbox_tls_verify(mailbox_id: int) -> RedirectResponse:
    """
    Переключатель проверки TLS сертификата для IMAP.
    """
    with session_scope() as s:
        mb = s.get(Mailbox, mailbox_id)
        if not mb:
            return RedirectResponse("/settings?error=mailbox_not_found", status_code=303)
        if mb.provider != "imap":
            return RedirectResponse("/settings?error=mailbox_wrong_type", status_code=303)
        cur = bool(getattr(mb, "imap_tls_verify", True))
        mb.imap_tls_verify = not cur
    return RedirectResponse("/settings", status_code=303)


@app.post("/actions/mailbox/delete/{mailbox_id}")
def action_mailbox_delete(mailbox_id: int) -> RedirectResponse:
    with session_scope() as s:
        mb = s.get(Mailbox, mailbox_id)
        if not mb:
            return RedirectResponse("/settings?error=mailbox_not_found", status_code=303)
        # Удаляем ящик и письма, чтобы не оставлять "висячие" данные.
        msg_ids = select(EmailMessage.id).where(EmailMessage.mailbox_id == mailbox_id)
        s.execute(delete(EmailAttachment).where(EmailAttachment.email_id.in_(msg_ids)))
        s.execute(delete(EmailMessage).where(EmailMessage.mailbox_id == mailbox_id))
        s.delete(mb)
    return RedirectResponse("/settings?notice=mailbox_deleted", status_code=303)


@app.post("/actions/mailbox/import-env")
def action_mailbox_import_env() -> RedirectResponse:
    """
    Миграция для удобства: берём IMAP креды из .env (старый формат) и
    создаём/обновляем стандартные ящики в БД, чтобы не вбивать заново.
    """
    import os

    def _g(k: str) -> str:
        return (os.environ.get(k) or "").strip()

    candidates = [
        {
            "name": "Яндекс",
            "host": _g("YANDEX_IMAP_HOST") or "imap.yandex.ru",
            "port": int((_g("YANDEX_IMAP_PORT") or "993").strip() or "993"),
            "user": _g("YANDEX_IMAP_USER"),
            "password": _g("YANDEX_IMAP_PASSWORD"),
        },
        {
            "name": "Mail.ru",
            "host": _g("MAILRU_IMAP_HOST") or "imap.mail.ru",
            "port": int((_g("MAILRU_IMAP_PORT") or "993").strip() or "993"),
            "user": _g("MAILRU_IMAP_USER"),
            "password": _g("MAILRU_IMAP_PASSWORD"),
        },
    ]

    q = get_queue()
    touched = 0
    for c in candidates:
        if not (c["user"] and c["password"]):
            continue
        with session_scope() as s:
            mb = s.scalars(select(Mailbox).where(Mailbox.provider == "imap", Mailbox.name == c["name"])).first()
            if not mb:
                mb = Mailbox(provider="imap", name=c["name"], imap_folder="INBOX", imap_last_uid=None, is_enabled=True)
                s.add(mb)
                s.flush()
            mb.is_enabled = True
            mb.imap_host_enc = encrypt_str(c["host"])
            mb.imap_port = int(c["port"])
            mb.imap_user_enc = encrypt_str(c["user"])
            mb.imap_password_enc = encrypt_str(c["password"])
            if not (mb.imap_folder or "").strip():
                mb.imap_folder = "INBOX"
            mailbox_id = mb.id

        touched += 1
        q.enqueue(sync_imap_mailbox, mailbox_id)

    return RedirectResponse(f"/settings?notice=env_imported&count={touched}", status_code=303)

