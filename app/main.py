from __future__ import annotations

import datetime as dt
import logging

from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse
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
from app.models import EmailMessage, Mailbox
from app.queue import get_queue
from app.schema import ensure_schema
from app.settings import settings
from app.oauth_state import consume_state, issue_state
from app.app_state import AiRunStatus, get_ai_run_status, now_iso, set_ai_run_status, set_ai_stop_flag, get_ai_test_result

app = FastAPI(title="Email Assistant")
templates = Jinja2Templates(directory="templates")
log = logging.getLogger("email-assistant")


@app.on_event("startup")
def _startup() -> None:
    log.warning("startup: create_all begin")
    Base.metadata.create_all(bind=engine)
    log.warning("startup: create_all ok")
    log.warning("startup: ensure_schema begin")
    ensure_schema()
    log.warning("startup: ensure_schema ok")
    log.warning("startup: bootstrap imap begin")
    _bootstrap_imap_mailboxes()
    log.warning("startup: bootstrap imap ok")
    _bootstrap_gmail_mailbox()


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
) -> HTMLResponse:
    with session_scope() as s:
        mailboxes = list(s.scalars(select(Mailbox).order_by(Mailbox.id.asc())))
        qry = (
            select(EmailMessage)
            .where(EmailMessage.is_archived == False)  # noqa: E712
            .order_by(EmailMessage.date.desc().nullslast(), EmailMessage.id.desc())
            .limit(200)
        )
        if category:
            qry = qry.where(EmailMessage.category == category)
        mailbox_id_int: int | None = None
        if mailbox_id and mailbox_id.strip().isdigit():
            mailbox_id_int = int(mailbox_id.strip())
            qry = qry.where(EmailMessage.mailbox_id == mailbox_id_int)
        if q and q.strip():
            term = f"%{q.strip()}%"
            qry = qry.where(
                or_(
                    EmailMessage.subject.ilike(term),
                    EmailMessage.from_email.ilike(term),
                    EmailMessage.snippet.ilike(term),
                )
            )
        emails = list(s.scalars(qry))

    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={
            "emails": emails,
            "category": category or "",
            "mailboxes": mailboxes,
            "mailbox_id": mailbox_id_int or "",
            "search_q": q or "",
        },
    )


@app.post("/actions/email/{email_id}/archive")
def action_archive_email(email_id: int) -> RedirectResponse:
    with session_scope() as s:
        s.execute(update(EmailMessage).where(EmailMessage.id == email_id).values(is_archived=True))
    return RedirectResponse("/", status_code=303)


@app.post("/actions/email/{email_id}/set-category")
def action_set_category(email_id: int, category: str) -> RedirectResponse:
    if category not in {"important", "normal", "newsletter", "spam_candidate"}:
        return RedirectResponse("/", status_code=303)
    with session_scope() as s:
        s.execute(update(EmailMessage).where(EmailMessage.id == email_id).values(category=category))
    return RedirectResponse("/", status_code=303)


@app.get("/settings", response_class=HTMLResponse)
def settings_page(request: Request) -> HTMLResponse:
    with session_scope() as s:
        mailboxes = list(s.scalars(select(Mailbox).order_by(Mailbox.id.asc())))
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
            "ai_base_url": settings.ai_base_url,
            "ai_model": settings.ai_model,
        },
    )


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

