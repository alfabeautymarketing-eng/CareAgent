# Copilot / Agent quick guide for AgentCare

Этот файл поможет AI‑агентам быстро стать продуктивными в репозитории AgentCare.

## Что важно знать (Big picture) ✅

- **Назначение**: Python-сервис для автоматизации Google Sheets + AI (Gemini) — синхронизация листов, обработка прайсов, генерация инвойсов.
- **Основные слои**: `src/api/` (FastAPI + webhooks), `src/core/` (SyncEngine, RulesEngine, Transaction), `src/services/` (price_processor, sync_service и т.д.), `src/parsers/`, `src/integrations/` и `src/utils/`.
- **Конфиги**: секреты в `.env`, общие настройки `config/settings.yaml`, проектные настройки `config/projects/*.yaml`, правила — `config/rules/cascade_rules.yaml`.

## Быстрый старт (Run / Dev / Docker) 🔧

- Локально (dev): `make dev` (или `poetry run uvicorn src.main:app --reload`)
- Прод: `make run` (uvicorn без reload, `--workers 4`)
- Docker: `docker compose up -d` / `make docker-up`
- Тесты: `make test` (внутри использует `poetry run pytest -v --cov=src`)
- Линт/формат: `make lint` / `make format` (ruff, mypy, black)

## Важные API & формат взаимодействия 📡

- Webhook: `POST /webhook/sheets/{project}` — payload = `SheetEvent` (см. `src/api/webhooks.py`).
  - Проекты: **`mt`, `sk`, `ss`** — валидируются в эндпоинте.
  - Подпись: заголовок `X-Webhook-Signature` содержит HMAC `sha256=<hex>`; в dev режиме проверка пропускается, если `settings.webhook_secret` пуст.
- Manual sync: `POST /webhook/sync/{project}` — форсирует полный sync.
- Health: `GET /health` — отдаёт basic checks (stub на текущий момент).

## Кодовые конвенции и паттерны 🧭

- Асинхронность: используйте `async` в web-слое и при работе с IO.
- Pydantic / Settings: конфигурация через `src/utils/config.py` (pydantic-settings, `.env`).
- Логирование: структурированные логи через `src/utils/logger.py`; примеры: `logger.info("sync_completed", extra={...})`.
- Парсеры: `src/parsers/BaseParser` интерфейс (`parse`, `validate`, `transform`); создавайте парсер для каждого проекта (`mt_parser`, `sk_parser`, `ss_parser`).
- Транзакции и retry: архитектура ожидает `transaction` decorator и `retry` logic (см. docs/ARCHITECTURE.md). Ищите TODO/`@transaction` и `@retry` в коде.

## Integrations & секреты 🔐

- Gemini (AI): используется `google-generativeai` (конфиг через `GEMINI_API_KEY` / `settings.gemini_api_key`).
- Google APIs: `gspread`, `google-api-python-client`, credentials в `config/credentials.json` или base64 через env.
- Redis: очередь/кэш — `settings.redis_url` (по умолчанию `redis://localhost:6379/0`).

## Tests & CI expectations ✅

- Тесты асинхронные: `pytest-asyncio` активирован (см. `pyproject.toml` / pytest config).
- При добавлении фич: добавляйте unit/integration тесты в `tests/` и запускайте `make test`.

## Where to look for TODOs / hot spots 🔎

- `src/main.py` — lifecycle TODOs (init Redis, validate Google creds).
- `src/integrations/` — большинство внешних клиентов ещё не реализованы; implement clients here and wire them in lifespan.
- `src/core/` — SyncEngine/RulesEngine are central: changes here affect whole flow — prefer small, covered PRs.
- `docs/ARCHITECTURE.md` & `docs/MODULES.md` — canonical source for design decisions.

## PR & contributor guidelines 🧩

- Small focused PRs, 1 logical change per PR.
- Include/extend tests for behavior changes.
- Update `docs/` when public behavior or config changes.

---

## GAS / clasp / CI notes ⚙️

- GAS files may live in `gas/` or (in some forks) at the repo root; `scripts/deploy_all.sh` now auto-detects both locations.
- To deploy GAS locally:
  - Install and login: `npm i -g @google/clasp` (or use `nvm`) then `clasp login`.
  - Create `.clasp.json` in the GAS dir: `printf '{"scriptId":"<SCRIPT_ID>","rootDir":"."}' > .clasp.json` and run `clasp push -f`.
- For CI (recommended): store your clasp credentials in `CLASP_CREDENTIALS` (content of `~/.clasprc.json`) and an SSH key in `SSH_PRIVATE_KEY`, then use the provided GitHub Action workflow (`.github/workflows/deploy.yml`) which runs `./scripts/deploy_all.sh` on `push` to `main`.
- If you use a service account: `clasp login --creds ./sa.json` and put `sa.json` into a repo secret; ensure scopes and permissions are configured for script deployment.

---

Если что-то непонятно или нужно больше примеров (например, шаблон PR или подробности по `rules` YAML), скажите — добавлю уточнения. 💡
