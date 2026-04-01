# Email Assistant (MVP)

Личный проект “виртуальный секретарь” для почты: сбор писем, оценка важности, выжимка и уборка рассылок/спама.

⚠️ **Важно**: файл `.env` содержит секреты. Он **не должен** попадать в git.

## Запуск (Docker Compose)

1) Сгенерируй ключ шифрования:

```bash
python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
```

2) Создай `.env` (можно от `.env.example`) и укажи:

- `APP_ENCRYPTION_KEY=...`
- Gmail OAuth (если нужен Gmail):
  - `GMAIL_OAUTH_CLIENT_ID`
  - `GMAIL_OAUTH_CLIENT_SECRET`
  - `GMAIL_OAUTH_REDIRECT_URI` (по умолчанию подходит для локального запуска)
- AI-провайдер (RouterAI/OpenRouter/Cursor proxy) — см. разделы ниже

3) Запусти:

```bash
docker compose up --build
```

Сервис будет доступен на `http://localhost:8000`.

## AI через Cursor (ключ из Cloud Agents)

Ключ, который выдаётся в `https://cursor.com/dashboard/cloud-agents`, относится к **Cursor Cloud Agents API** (управление агентами на репозиториях) и **не является OpenAI-совместимым** `chat/completions` endpoint.

Чтобы этот проект (он использует OpenAI-совместимый API) работал с Cursor, нужен **локальный OpenAI‑compatible proxy** для Cursor CLI, например `cursor-api-proxy`.

Минимальная схема на Windows:

1) Установи Node.js 18+.
2) Установи Cursor agent CLI и залогинься (или используй переменную `CURSOR_API_KEY` для автоматизации, если поддерживается твоей установкой `agent`):

```bash
curl https://cursor.com/install -fsS | bash
agent login
agent --list-models
```

3) Запусти прокси:

```bash
npx cursor-api-proxy
```

4) В `.env` для нашего приложения укажи (для Docker):

- `AI_BASE_URL=http://host.docker.internal:8765/v1`
- `AI_API_KEY=unused` (или значение `CURSOR_BRIDGE_API_KEY`, если ты включил авторизацию на прокси)
- `AI_MODEL=gpt-5.4-nano-low` (рекомендуется фиксировать конкретную модель)

## AI через RouterAI (рекомендуется для “дёшево и стабильно”)

RouterAI даёт **OpenAI‑совместимый** `chat/completions` API и оплату в рублях.

- **base_url**: `https://routerai.ru/api/v1`
- **endpoint**: `POST /chat/completions`
- **ключ**: `Authorization: Bearer ...`
- **пример модели (дешёвая и сильная)**: `deepseek/deepseek-v3.2`

См. документацию RouterAI: `https://routerai.ru/docs/guides`

## IMAP (Яндекс / Mail.ru)

- Для Яндекс/Mail.ru обычно нужен **пароль приложения** (если включён 2FA) и включённый IMAP-доступ.
- IMAP‑ящики добавляются через UI: **Настройки → “Добавить IMAP ящик”**.
- После добавления можно нажать **“Синхронизировать всё”** (на главной странице тоже есть кнопка синхронизации).

## Полезные заметки

- AI-обработка защищена от “зависаний” на одном письме: на одно письмо стоит жёсткий лимит времени, после чего применяется безопасный fallback.
- Если модель иногда отдаёт “мета-текст” (про формат/JSON), приложение заменит explanation на понятное эвристическое объяснение.

## Тесты

```bash
pytest -q
```

