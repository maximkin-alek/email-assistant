from __future__ import annotations

import datetime as dt
import json
from dataclasses import asdict, dataclass

from app.queue import get_redis


@dataclass(frozen=True)
class AiTestResult:
    ok: bool
    configured_base_url: str = ""
    configured_model: str = ""
    used_model: str = ""
    message: str = ""
    tested_at: str = ""


@dataclass(frozen=True)
class AiTestStatus:
    running: bool
    started_at: str = ""
    finished_at: str = ""
    message: str = ""


AI_TEST_KEY = "ai_test:last"
AI_TEST_STATUS_KEY = "ai_test:status"
AI_RUN_KEY = "ai_run:last"
AI_STOP_KEY = "ai_run:stop"


def set_ai_test_result(result: AiTestResult) -> None:
    get_redis().set(AI_TEST_KEY, json.dumps(asdict(result), ensure_ascii=False))


def get_ai_test_result() -> AiTestResult | None:
    raw = get_redis().get(AI_TEST_KEY)
    if not raw:
        return None
    try:
        d = json.loads(raw.decode("utf-8", errors="ignore"))
        # Backward-compat: раньше было {ok, model, message}
        legacy_model = str(d.get("model") or "")
        used_model = str(d.get("used_model") or "") or legacy_model
        configured_model = str(d.get("configured_model") or "") or str(d.get("requested_model") or "")
        configured_base_url = str(d.get("configured_base_url") or "") or str(d.get("base_url") or "")
        tested_at = str(d.get("tested_at") or "")
        return AiTestResult(
            ok=bool(d.get("ok")),
            configured_base_url=configured_base_url,
            configured_model=configured_model,
            used_model=used_model,
            message=str(d.get("message") or ""),
            tested_at=tested_at,
        )
    except Exception:
        return None


def set_ai_test_status(status: AiTestStatus) -> None:
    get_redis().set(AI_TEST_STATUS_KEY, json.dumps(asdict(status), ensure_ascii=False))


def get_ai_test_status() -> AiTestStatus | None:
    raw = get_redis().get(AI_TEST_STATUS_KEY)
    if not raw:
        return None
    try:
        d = json.loads(raw.decode("utf-8", errors="ignore"))
        return AiTestStatus(
            running=bool(d.get("running")),
            started_at=str(d.get("started_at") or ""),
            finished_at=str(d.get("finished_at") or ""),
            message=str(d.get("message") or ""),
        )
    except Exception:
        return None


def now_iso() -> str:
    return dt.datetime.now(dt.UTC).isoformat()


@dataclass(frozen=True)
class AiRunStatus:
    running: bool
    started_at: str = ""
    updated_at: str = ""
    finished_at: str = ""
    total: int = 0
    processed: int = 0
    ok: int = 0
    failed: int = 0
    message: str = ""


def set_ai_run_status(status: AiRunStatus) -> None:
    get_redis().set(AI_RUN_KEY, json.dumps(asdict(status), ensure_ascii=False))


def get_ai_run_status() -> AiRunStatus | None:
    raw = get_redis().get(AI_RUN_KEY)
    if not raw:
        return None
    try:
        d = json.loads(raw.decode("utf-8", errors="ignore"))
        return AiRunStatus(
            running=bool(d.get("running")),
            started_at=str(d.get("started_at") or ""),
            updated_at=str(d.get("updated_at") or ""),
            finished_at=str(d.get("finished_at") or ""),
            total=int(d.get("total") or 0),
            processed=int(d.get("processed") or 0),
            ok=int(d.get("ok") or 0),
            failed=int(d.get("failed") or 0),
            message=str(d.get("message") or ""),
        )
    except Exception:
        return None


def set_ai_stop_flag(value: bool) -> None:
    get_redis().set(AI_STOP_KEY, "1" if value else "0")


def get_ai_stop_flag() -> bool:
    raw = get_redis().get(AI_STOP_KEY)
    if not raw:
        return False
    return raw.decode("utf-8", errors="ignore").strip() == "1"

