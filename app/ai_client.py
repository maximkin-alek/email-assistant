from __future__ import annotations

import json
import time
from dataclasses import dataclass

import httpx

from app.settings import settings


@dataclass(frozen=True)
class AiResult:
    category: str
    score: int
    summary: str
    explanation: str
    model: str


def _clamp_score(score: int) -> int:
    return max(0, min(100, int(score)))


def _truncate(s: str | None, limit: int) -> str | None:
    if s is None:
        return None
    s = str(s)
    if len(s) <= limit:
        return s
    return s[:limit] + "…"


def _to_payload_text(
    subject: str | None,
    from_email: str | None,
    snippet: str | None,
    body_text: str | None,
) -> str:
    """
    Текстовый формат заметно стабильнее, чем вложенный JSON, для некоторых прокси/моделей.
    """
    subj = (_truncate(subject, 500) or "").strip()
    frm = (_truncate(from_email, 500) or "").strip()
    snip = (_truncate(snippet, 1200) or "").strip()
    body = (_truncate(body_text, 8000) or "").strip()
    return (
        "SUBJECT: " + (subj or "(нет)") + "\n"
        "FROM: " + (frm or "(нет)") + "\n"
        "SNIPPET: " + (snip or "(нет)") + "\n"
        "BODY:\n" + (body or "(нет)")
    )


def classify_and_summarize(
    subject: str | None,
    from_email: str | None,
    snippet: str | None,
    body_text: str | None,
) -> AiResult:
    if not settings.ai_api_key:
        raise RuntimeError("AI_API_KEY не задан в .env")

    base = settings.ai_base_url.rstrip("/")
    url = f"{base}/chat/completions"
    model = settings.ai_model

    # Важно: некоторые прокси/провайдеры (в т.ч. через CLI) могут странно обрабатывать role=system.
    # Поэтому инструкцию даём в user-сообщении.
    instruction = (
        "Ответь ОДНОЙ строкой: только JSON (без markdown/текста вокруг).\n"
        "JSON-схема: {\"category\":\"normal\",\"score\":10,\"summary\":\"...\",\"explanation\":\"...\"}\n"
        "category ∈ [important, normal, newsletter, spam_candidate]. score 0..100.\n"
        "explanation должен быть непустой строкой и объяснять решение по этому письму.\n"
        "Письмо:\n"
    )
    instruction_retry = (
        "ПОВТОР: верни ТОЛЬКО JSON по схеме {\"category\":\"normal\",\"score\":10,\"summary\":\"...\",\"explanation\":\"...\"}.\n"
        "Никакого текста вокруг, никаких вопросов.\n"
        "Письмо:\n"
    )
    payload = _to_payload_text(subject=subject, from_email=from_email, snippet=snippet, body_text=body_text)
    user = instruction + payload

    start = time.monotonic()
    # Жёсткий бюджет времени на одно письмо: чтобы ai_run НИКОГДА не зависал на одном сообщении.
    deadline_s = 30.0

    def _time_left_s() -> float:
        return deadline_s - (time.monotonic() - start)

    def _post(messages: list[dict], temperature: float, model_override: str | None = None) -> dict:
        left = _time_left_s()
        if left <= 0:
            raise RuntimeError("AI timeout: превышен лимит времени на обработку письма.")
        timeout = httpx.Timeout(connect=5.0, read=min(20.0, left), write=10.0, pool=5.0)
        with httpx.Client(timeout=timeout) as client:
            resp = client.post(
                url,
                headers={"Authorization": f"Bearer {settings.ai_api_key}", "Content-Type": "application/json"},
                json={
                    "model": (model_override or model),
                    "messages": messages,
                    "temperature": temperature,
                    # Если провайдер поддерживает JSON mode — это резко повышает стабильность.
                    "response_format": {"type": "json_object"},
                },
            )
            try:
                resp.raise_for_status()
            except httpx.HTTPStatusError as e:
                code = e.response.status_code
                if code == 401:
                    raise RuntimeError(
                        "AI вернул 401 (Unauthorized): проверь AI_API_KEY и AI_BASE_URL. "
                        "Если используешь cursor-api-proxy и включал CURSOR_BRIDGE_API_KEY, "
                        "то AI_API_KEY должен совпадать с ним (заголовок Authorization: Bearer ...)."
                    ) from e
                if code == 402:
                    raise RuntimeError(
                        "AI вернул 402 (Payment Required): у провайдера нет доступного кредита/лимита или не активирован биллинг."
                    ) from e
                raise RuntimeError(f"AI: ошибка HTTP {code}. Детали: {e.response.text[:500]}") from e
            try:
                return resp.json()
            except Exception as e:
                snippet = (resp.text or "").strip()
                if not snippet:
                    raise RuntimeError("AI вернул пустой ответ (не удалось распарсить JSON).") from e
                raise RuntimeError(f"AI вернул неожиданный ответ (не JSON): {snippet[:500]}") from e

    def _extract_json_object(text: str) -> dict | None:
        txt = (text or "").strip()
        if not txt:
            return None
        try:
            obj = json.loads(txt)
            return obj if isinstance(obj, dict) else None
        except Exception:
            start = txt.find("{")
            end = txt.rfind("}")
            if start != -1 and end != -1 and end > start:
                try:
                    obj = json.loads(txt[start : end + 1])
                    return obj if isinstance(obj, dict) else None
                except Exception:
                    return None
            return None

    def _get_content(data: dict) -> str | None:
        try:
            choices = data.get("choices")
            if not isinstance(choices, list) or not choices:
                return None
            msg = choices[0].get("message") if isinstance(choices[0], dict) else None
            if not isinstance(msg, dict):
                return None
            content = msg.get("content")
            return content if isinstance(content, str) and content.strip() else None
        except Exception:
            return None

    def _fallback_result() -> AiResult:
        text = f"{subject or ''}\n{from_email or ''}\n{snippet or ''}\n{body_text or ''}".lower()
        category = "normal"
        score = 10
        if any(x in text for x in ["unsubscribe", "рассылка", "newsletter", "акция", "скидк", "промокод"]):
            category = "newsletter"
            score = 5
        if any(x in text for x in ["срочно", "urgent", "asap", "оплат", "счет", "invoice", "код", "подтверд"]):
            category = "important"
            score = 80
        summary = (snippet or body_text or "").strip()[:400]
        if not summary:
            summary = (subject or "").strip()[:200]
        category2, score2 = _heuristic_adjust(category, score)
        return AiResult(
            category=category2,
            score=score2,
            summary=summary or "(нет текста)",
            explanation=_heuristic_explanation(category2),
            model=model,
        )

    def _heuristic_adjust(category: str, score: int) -> tuple[str, int]:
        text = f"{subject or ''}\n{from_email or ''}\n{snippet or ''}\n{body_text or ''}".lower()
        # Если модель/прокси “залипли” на normal/10 — слегка подправим.
        if any(x in text for x in ["код", "confirmation code", "подтвержден", "verify", "авториз", "вход", "парол", "осаго", "полис", "оплат", "счет", "invoice"]):
            return ("important", max(score, 70))
        if any(x in text for x in ["unsubscribe", "рассылка", "newsletter", "акция", "скидк", "промокод", "бонус", "кэшбэк"]):
            return ("newsletter", min(score, 25))
        if any(x in text for x in ["вы выиграли", "подарок", "bitcoin", "крипто", "срочно подтвердите", "переходите по ссылке"]):
            return ("spam_candidate", min(score, 10))
        return (category, score)

    def _heuristic_explanation(category: str) -> str:
        subj = (subject or "").strip()
        frm = (from_email or "").strip()
        snip = (snippet or "").strip()
        text = f"{subj}\n{frm}\n{snip}".lower()
        if category == "important":
            return "Похоже на письмо, требующее действия (оплата/код/срочно). Проверьте детали и не игнорируйте сроки."
        if category == "newsletter":
            return "Похоже на рассылку/промо: можно прочитать выборочно или отписаться, если не актуально."
        if category == "spam_candidate":
            return "Есть признаки нежелательной рассылки/подозрительного письма. Открывайте осторожно, не вводите коды/пароли."
        # normal
        if any(x in text for x in ["сохранили", "товары", "корзин", "покупк", "акци", "скидк", "бонус"]):
            return "Похоже на маркетинговое письмо/напоминание о покупке. Важно только если вы планировали заказ/акцию."
        return "Похоже на обычное информационное письмо без срочного действия."

    def _looks_like_meta_explanation(text: str) -> bool:
        t = (text or "").lower()
        bad = [
            "зачем выполняется",
            "зачем выполняется этот запрос",
            "как результат поможет",
            "как это поможет",
            "поиск",
            "запрос",
            "следующему шагу",
            "я ищу актуальную информацию",
            "интернет",
            "в интернете",
            "web",
            "browser",
            "ссылка",
            "источник",
            "нет контекста",
            "нет задачи",
            "соответствие формату",
            "соответствия формату",
            "значения по умолчанию",
            "заданные значения",
            "запрос не содержит",
            "в сообщении отсутствует",
            "подтверждаю понимание",
            "если вы дадите",
            "если вы предоставите",
            "не указаны",
            "верну json",
            "нужными полями",
            "только json",
            "строго по схеме",
            "вне json",
            "по схеме",
            "json с полями",
            "json-объект",
            "json объект",
            "формат json",
            "в формате json",
            "формате {",
            "с полями",
            "без дополнительного текста",
            "без дополнительных комментариев",
            "форматирую ответ",
            "формирую ответ",
        ]
        return any(b in t for b in bad)

    try:
        data = _post([{"role": "user", "content": user}], temperature=0.0)
    except Exception as e:
        msg = str(e)
        # Ошибки конфигурации/биллинга лучше не скрывать.
        if "401" in msg or "Unauthorized" in msg or "402" in msg or "Payment Required" in msg:
            raise
        return _fallback_result()

    used_model = model
    if isinstance(data, dict):
        m = data.get("model")
        if isinstance(m, str) and m.strip():
            used_model = m.strip()

    content = _get_content(data) if isinstance(data, dict) else None
    if not content:
        return _fallback_result()

    parsed = _extract_json_object(content)
    if parsed is None:
        # 2-я попытка, но только если ещё есть время.
        if _time_left_s() > 3:
            try:
                data2 = _post([{"role": "user", "content": instruction_retry + "\n" + payload}], temperature=0.0)
            except Exception as e:
                msg = str(e)
                if "401" in msg or "Unauthorized" in msg or "402" in msg or "Payment Required" in msg:
                    raise
                return _fallback_result()
        else:
            return _fallback_result()
        content2 = _get_content(data2) if isinstance(data2, dict) else None
        parsed = _extract_json_object(content2) if isinstance(content2, str) else None
    if parsed is None:
        return _fallback_result()

    category = str(parsed.get("category") or "normal")
    if category not in {"important", "normal", "newsletter", "spam_candidate"}:
        category = "normal"
    score_val = parsed.get("score")
    if isinstance(score_val, (int, float)) and not isinstance(score_val, bool):
        score = _clamp_score(int(round(score_val)))
    else:
        score = 10
    summary = str(parsed.get("summary") or "").strip() or (snippet or body_text or "").strip()[:400]
    explanation = str(parsed.get("explanation") or "").strip()
    if not explanation:
        # Если модель вернула JSON, но без explanation — это по сути неудачный ответ.
        if _time_left_s() <= 3:
            explanation = "AI не вернул explanation (таймаут)."
        else:
            try:
                data3 = _post(
            [
                {
                    "role": "user",
                    "content": (
                        "ПОВТОР: в прошлом JSON ты не заполнил поле explanation.\n"
                        "Верни ТОЛЬКО JSON. Поле explanation ОБЯЗАТЕЛЬНО и должно быть непустой строкой.\n"
                        "{"
                        "\"category\": \"important|normal|newsletter|spam_candidate\", "
                        "\"score\": 0-100, "
                        "\"summary\": \"1-3 строки\", "
                        "\"explanation\": \"коротко почему\""
                        "}\n"
                        "Данные письма:\n"
                        + payload
                    ),
                }
            ],
            temperature=0.0,
                )
            except Exception as e:
                msg = str(e)
                if "401" in msg or "Unauthorized" in msg or "402" in msg or "Payment Required" in msg:
                    raise
                explanation = "AI не вернул explanation (ошибка/таймаут)."
        content3 = _get_content(data3) if isinstance(data3, dict) else None
        if isinstance(content3, str) and content3.strip():
            txt3 = content3.strip()
            try:
                parsed3 = json.loads(txt3)
            except Exception:
                start3 = txt3.find("{")
                end3 = txt3.rfind("}")
                if start3 != -1 and end3 != -1 and end3 > start3:
                    parsed3 = json.loads(txt3[start3 : end3 + 1])
                else:
                    parsed3 = {}
            explanation = str((parsed3 or {}).get("explanation") or "").strip()
            if parsed3:
                # если на ретрае пришёл валидный JSON — обновим и остальные поля
                category = str(parsed3.get("category") or category)
                if category not in {"important", "normal", "newsletter", "spam_candidate"}:
                    category = "normal"
                score3 = parsed3.get("score")
                if isinstance(score3, (int, float)) and not isinstance(score3, bool):
                    score = _clamp_score(int(round(score3)))
                summary2 = str(parsed3.get("summary") or "").strip()
                if summary2:
                    summary = summary2
        if not explanation:
            # Иногда конкретная модель стабильно отдаёт пустой explanation.
            # Пробуем альтернативную модель без изменения .env.
            alt_model = "gpt-5.4-nano-low"
            if _time_left_s() <= 3:
                explanation = "AI не вернул explanation (таймаут)."
            else:
                try:
                    data_alt = _post(
                [
                    {
                        "role": "user",
                        "content": (
                            "Верни ТОЛЬКО JSON по схеме "
                            "{\"category\":\"normal\",\"score\":10,\"summary\":\"...\",\"explanation\":\"...\"}.\n"
                            "explanation должен быть непустой строкой (минимум 20 символов) и про письмо.\n"
                            "Письмо:\n"
                            + payload
                        ),
                    }
                ],
                temperature=0.0,
                model_override=alt_model,
                    )
                except Exception as e:
                    msg = str(e)
                    if "401" in msg or "Unauthorized" in msg or "402" in msg or "Payment Required" in msg:
                        raise
                    data_alt = {}
                used_alt = data_alt.get("model") if isinstance(data_alt, dict) else None
                content_alt = _get_content(data_alt) if isinstance(data_alt, dict) else None
                parsed_alt = _extract_json_object(content_alt) if isinstance(content_alt, str) else None
                if isinstance(parsed_alt, dict):
                    exp_alt = str(parsed_alt.get("explanation") or "").strip()
                    if exp_alt:
                        explanation = exp_alt
                        if isinstance(used_alt, str) and used_alt.strip():
                            used_model = used_alt.strip()
            if not explanation:
                explanation = "AI не вернул explanation (переобработка не помогла)"

    # Защита от “мета-ответов” (модель иногда объясняет не письмо, а процесс).
    if _looks_like_meta_explanation(explanation):
        if _time_left_s() > 3:
            try:
                data4 = _post(
            [
                {
                    "role": "user",
                    "content": (
                        "ПОВТОР: ты написал мета-ответ (про формат/JSON/схему/запрос/поиск), это запрещено.\n"
                        "Нужно объяснение ТОЛЬКО про письмо: по содержанию письма (subject/snippet/body) и отправителю.\n"
                        "НЕ упоминай JSON, схему, формат, запрос, инструменты.\n"
                        "Верни ТОЛЬКО JSON по схеме.\n"
                        "{"
                        "\"category\": \"important|normal|newsletter|spam_candidate\", "
                        "\"score\": 0-100, "
                        "\"summary\": \"1-3 строки\", "
                        "\"explanation\": \"коротко почему\""
                        "}\n"
                        "Письмо:\n"
                        + payload
                    ),
                }
            ],
            temperature=0.0,
                )
                content4 = _get_content(data4) if isinstance(data4, dict) else None
                if isinstance(content4, str) and content4.strip():
                    txt4 = content4.strip()
                    try:
                        parsed4 = json.loads(txt4)
                    except Exception:
                        start4 = txt4.find("{")
                        end4 = txt4.rfind("}")
                        parsed4 = (
                            json.loads(txt4[start4 : end4 + 1])
                            if (start4 != -1 and end4 != -1 and end4 > start4)
                            else {}
                        )
                    if parsed4:
                        category = str(parsed4.get("category") or category)
                        if category not in {"important", "normal", "newsletter", "spam_candidate"}:
                            category = "normal"
                        score = _clamp_score(parsed4.get("score") if isinstance(parsed4.get("score"), int) else score)
                        summary2 = str(parsed4.get("summary") or "").strip()
                        if summary2:
                            summary = summary2
                        explanation2 = str(parsed4.get("explanation") or "").strip()
                        if explanation2:
                            explanation = explanation2
            except Exception as e:
                msg = str(e)
                if "401" in msg or "Unauthorized" in msg or "402" in msg or "Payment Required" in msg:
                    raise

    # Если даже после ретрая explanation остаётся мета-текстом — ставим понятное, но честное эвристическое объяснение.
    if _looks_like_meta_explanation(explanation):
        explanation = _heuristic_explanation(category)

    # Если AI не смог дать нормальный ответ (таймаут/не-JSON/ошибка) — тоже даём человеческое explanation.
    if explanation.startswith("AI не вернул") or explanation.startswith("Ошибка AI:"):
        explanation = _heuristic_explanation(category)

    category, score = _heuristic_adjust(category, score)

    return AiResult(category=category, score=score, summary=summary, explanation=explanation, model=used_model)

