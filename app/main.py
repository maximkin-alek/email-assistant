from __future__ import annotations

import datetime as dt
import logging
import re
import json
from urllib.parse import urlparse
from email.header import decode_header, make_header

from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy import or_, select, update
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
    ai_test_model,
    recompute_all_basic,
    sync_gmail_mailbox,
    sync_imap_mailbox,
)
from app.models import AppSetting, EmailAttachment, EmailMessage, Mailbox
from app.queue import get_queue
from app.schema import ensure_schema
from app.settings import settings
from app.oauth_state import consume_state, issue_state
from app.app_state import AiRunStatus, get_ai_run_status, now_iso, set_ai_run_status, set_ai_stop_flag, get_ai_test_result

app = FastAPI(title="Email Assistant")
templates = Jinja2Templates(directory="templates")
log = logging.getLogger("email-assistant")

app.mount("/static", StaticFiles(directory="static"), name="static")


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

    _run_quick("bootstrap imap", _bootstrap_imap_mailboxes, 3.0)
    _run_quick("bootstrap gmail", _bootstrap_gmail_mailbox, 3.0)


def _bootstrap_imap_mailboxes() -> None:
    """
    Простой bootstrap для итерации 2: если в .env заданы креды — создаём/обновляем mailboxes.
    UI настройки добавим позже, когда синхронизация устоится.
    """
    candidates: list[tuple[str, str, int, str | None, str | None]] = [
        ("Яндекс", settings.yandex_imap_host, settings.yandex_imap_port, settings.yandex_imap_user, settings.yandex_imap_password),
        ("Mail.ru", settings.mailru_imap_host, settings.mailru_imap_port, settings.mailru_imap_user, settings.mailru_imap_password),
    ]

    with session_scope() as s:
        for name, host, port, user, password in candidates:
            if not (user and password):
                continue
            mb = s.scalars(select(Mailbox).where(Mailbox.provider == "imap", Mailbox.name == name)).first()
            if not mb:
                mb = Mailbox(provider="imap", name=name, imap_folder="INBOX", imap_last_uid=None, is_enabled=True)
                s.add(mb)
                s.flush()

            mb.imap_host_enc = encrypt_str(host)
            mb.imap_port = int(port)
            mb.imap_user_enc = encrypt_str(user)
            mb.imap_password_enc = encrypt_str(password)
            if not mb.imap_folder:
                mb.imap_folder = "INBOX"


def _bootstrap_gmail_mailbox() -> None:
    with session_scope() as s:
        mb = s.scalars(select(Mailbox).where(Mailbox.provider == "gmail")).first()
        if not mb:
            mb = Mailbox(provider="gmail", name="Gmail", is_enabled=True)
            s.add(mb)


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
) -> HTMLResponse:
    with session_scope() as s:
        mailboxes = list(s.scalars(select(Mailbox).order_by(Mailbox.id.asc())))
        mailbox_map = {m.id: m for m in mailboxes}

        base_qry = select(EmailMessage).where(EmailMessage.is_archived == False)  # noqa: E712
        qry = (
            select(EmailMessage)
            .where(EmailMessage.is_archived == False)  # noqa: E712
            .limit(200)
        )
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
        if sort_v == "score":
            qry = qry.order_by(EmailMessage.score.desc().nullslast(), EmailMessage.date.desc().nullslast(), EmailMessage.id.desc())
        else:
            sort_v = "date"
            qry = qry.order_by(EmailMessage.date.desc().nullslast(), EmailMessage.id.desc())

        emails = list(s.scalars(qry))
        # Дайджест "Сегодня важное" (группируем по треду/теме)
        now_local = dt.datetime.now().astimezone()
        start_local = now_local.replace(hour=0, minute=0, second=0, microsecond=0)
        start_utc = start_local.astimezone(dt.UTC)
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

    quick_links = {
        "unread": _u(quick="unread"),
        "important": _u(quick="important"),
        "today": _u(quick="today"),
        "week": _u(quick="week"),
        "clear_quick": _u(quick=""),
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

    today_groups = sorted(groups.values(), key=lambda x: (-(x["unread"]), -(x["count"]), -(x["top"].score or 0), -(x["top"].id)))

    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={
            "emails": emails,
            "today_groups": today_groups[:8],
            "category": category or "",
            "cat_ru": cat_ru,
            "mailboxes": mailboxes,
            "mailbox_map": mailbox_map,
            "mailbox_id": mailbox_id_int or "",
            "search_q": q or "",
            "unread_only": unread_only,
            "quick": (quick or "").strip().lower(),
            "sort": sort_v,
            "quick_links": quick_links,
            "thread": (thread or "").strip(),
            "subj": (subj or "").strip(),
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
            return {"ok": True, "updated": ids, "archived": [], "action": "mark_read"}
        s.execute(update(EmailMessage).where(EmailMessage.id.in_(ids)).values(is_archived=True))
        return {"ok": True, "updated": [], "archived": ids, "action": "archive"}


@app.post("/actions/email/{email_id}/archive")
def action_archive_email(email_id: int) -> RedirectResponse:
    with session_scope() as s:
        s.execute(update(EmailMessage).where(EmailMessage.id == email_id).values(is_archived=True))
    return RedirectResponse("/", status_code=303)


@app.post("/api/email/{email_id}/archive")
def api_archive_email(email_id: int) -> dict:
    with session_scope() as s:
        s.execute(update(EmailMessage).where(EmailMessage.id == email_id).values(is_archived=True))
    return {"ok": True, "email_id": email_id, "archived": True}


@app.post("/actions/email/{email_id}/set-category")
def action_set_category(email_id: int, category: str) -> RedirectResponse:
    if category not in {"important", "normal", "newsletter", "spam_candidate"}:
        return RedirectResponse("/", status_code=303)
    with session_scope() as s:
        s.execute(update(EmailMessage).where(EmailMessage.id == email_id).values(category=category))
    return RedirectResponse("/", status_code=303)


@app.post("/api/email/{email_id}/set-category")
def api_set_category(email_id: int, category: str = Form(...)) -> dict:
    if category not in {"important", "normal", "newsletter", "spam_candidate"}:
        return {"ok": False, "error": "invalid_category"}
    with session_scope() as s:
        s.execute(update(EmailMessage).where(EmailMessage.id == email_id).values(category=category))
    return {"ok": True, "email_id": email_id, "category": category}


@app.post("/actions/email/{email_id}/mark-read")
def action_mark_read(email_id: int) -> RedirectResponse:
    with session_scope() as s:
        s.execute(update(EmailMessage).where(EmailMessage.id == email_id).values(is_read=True))
    return RedirectResponse("/", status_code=303)


@app.post("/api/email/{email_id}/mark-read")
def api_mark_read(email_id: int) -> dict:
    with session_scope() as s:
        s.execute(update(EmailMessage).where(EmailMessage.id == email_id).values(is_read=True))
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
        elif action == "archive":
            s.execute(update(EmailMessage).where(EmailMessage.id.in_(ids)).values(is_archived=True))
        elif action == "set_category":
            if category in {"important", "normal", "newsletter", "spam_candidate"}:
                s.execute(update(EmailMessage).where(EmailMessage.id.in_(ids)).values(category=category))
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
        elif action == "archive":
            s.execute(update(EmailMessage).where(EmailMessage.id.in_(ids)).values(is_archived=True))
            archived = ids
        elif action == "set_category":
            if category in {"important", "normal", "newsletter", "spam_candidate"}:
                s.execute(update(EmailMessage).where(EmailMessage.id.in_(ids)).values(category=category))
                updated = ids
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

    return templates.TemplateResponse(
        request=request,
        name="settings.html",
        context={
            "mailboxes": mailboxes,
            "ai_test": get_ai_test_result(),
            "ai_run": run,
            "ai_run_started_h": _fmt_iso(run.started_at),
            "ai_run_finished_h": _fmt_iso(run.finished_at),
            "queue_pending": queue_pending,
            "queue_started": queue_started,
            "rules": rules,
            "ai_base_url": settings.ai_base_url,
            "ai_model": settings.ai_model,
        },
    )


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
    q.enqueue(ai_run, 10, at_front=True)
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
    q.enqueue(ai_test_model)
    return RedirectResponse("/settings", status_code=303)


@app.post("/actions/ai-retry-failed")
def action_ai_retry_failed() -> RedirectResponse:
    q = get_queue()
    q.enqueue(ai_retry_failed, 100)
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
        mailboxes = list(s.scalars(select(Mailbox.id).where(Mailbox.provider == "imap", Mailbox.is_enabled == True)))  # noqa: E712
    for mailbox_id in mailboxes:
        q.enqueue(sync_imap_mailbox, mailbox_id)
    return RedirectResponse("/settings", status_code=303)


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
def connect_gmail() -> RedirectResponse:
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

    state = issue_state("gmail")
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
    if not code or not state or consume_state(state) != "gmail":
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
    flow.fetch_token(code=code)

    creds = flow.credentials

    with session_scope() as s:
        mb = s.scalars(select(Mailbox).where(Mailbox.provider == "gmail")).first()
        if not mb:
            mb = Mailbox(provider="gmail", name="Gmail", is_enabled=True)
            s.add(mb)
            s.flush()
        mb.gmail_credentials_enc = encrypt_str(creds.to_json())
        # Профиль (email/historyId) подтянем в sync job. На колбэке не падаем,
        # даже если Gmail API ещё не включён в проекте.
        mb.last_sync_status = "ok"
        mb.last_sync_error = None

    return RedirectResponse("/", status_code=303)

