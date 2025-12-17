# AgentCare - Зависимости

## Обзор

Проект использует минимальный набор хорошо поддерживаемых библиотек.

---

## Основные зависимости

### Web Framework

| Пакет | Версия | Назначение |
|-------|--------|------------|
| `fastapi` | ^0.109 | HTTP API, webhooks |
| `uvicorn` | ^0.27 | ASGI сервер |
| `pydantic` | ^2.5 | Валидация данных |
| `python-multipart` | ^0.0.6 | Form data parsing |

### Google APIs

| Пакет | Версия | Назначение |
|-------|--------|------------|
| `gspread` | ^6.0 | Google Sheets (высокоуровневый) |
| `google-auth` | ^2.27 | OAuth авторизация |
| `google-api-python-client` | ^2.116 | Google Drive API |
| `google-generativeai` | ^0.4 | Gemini AI API |

### Data Processing

| Пакет | Версия | Назначение |
|-------|--------|------------|
| `pandas` | ^2.2 | Работа с данными |
| `openpyxl` | ^3.1 | Чтение Excel файлов |
| `xlrd` | ^2.0 | Чтение старых .xls файлов |

### Async & Background Tasks

| Пакет | Версия | Назначение |
|-------|--------|------------|
| `asyncio` | stdlib | Асинхронность |
| `aiohttp` | ^3.9 | Async HTTP клиент |
| `celery` | ^5.3 | Background tasks (опционально) |
| `redis` | ^5.0 | Кэш и очередь задач |

### Logging & Monitoring

| Пакет | Версия | Назначение |
|-------|--------|------------|
| `structlog` | ^24.1 | Структурированное логирование |
| `sentry-sdk` | ^1.40 | Error tracking (опционально) |

### Configuration

| Пакет | Версия | Назначение |
|-------|--------|------------|
| `pyyaml` | ^6.0 | YAML конфиги |
| `python-dotenv` | ^1.0 | Переменные окружения |
| `pydantic-settings` | ^2.1 | Типизированные настройки |

### Notifications (опционально)

| Пакет | Версия | Назначение |
|-------|--------|------------|
| `python-telegram-bot` | ^21.0 | Telegram уведомления |
| `aiosmtplib` | ^3.0 | Email уведомления |

### Development

| Пакет | Версия | Назначение |
|-------|--------|------------|
| `pytest` | ^8.0 | Тестирование |
| `pytest-asyncio` | ^0.23 | Async тесты |
| `pytest-cov` | ^4.1 | Coverage |
| `ruff` | ^0.2 | Linting |
| `black` | ^24.1 | Formatting |
| `mypy` | ^1.8 | Type checking |

---

## pyproject.toml

```toml
[tool.poetry]
name = "agentcare"
version = "0.1.0"
description = "Google Sheets automation with AI"
authors = ["Your Name <your@email.com>"]
readme = "README.md"
python = "^3.11"

[tool.poetry.dependencies]
python = "^3.11"

# Web
fastapi = "^0.109"
uvicorn = {extras = ["standard"], version = "^0.27"}
pydantic = "^2.5"
pydantic-settings = "^2.1"
python-multipart = "^0.0.6"

# Google
gspread = "^6.0"
google-auth = "^2.27"
google-api-python-client = "^2.116"
google-generativeai = "^0.4"

# Data
pandas = "^2.2"
openpyxl = "^3.1"

# Async
aiohttp = "^3.9"
redis = "^5.0"

# Config & Logging
pyyaml = "^6.0"
python-dotenv = "^1.0"
structlog = "^24.1"

[tool.poetry.group.dev.dependencies]
pytest = "^8.0"
pytest-asyncio = "^0.23"
pytest-cov = "^4.1"
ruff = "^0.2"
black = "^24.1"
mypy = "^1.8"

[tool.poetry.group.optional.dependencies]
# Background tasks (если нужен Celery)
celery = "^5.3"

# Notifications
python-telegram-bot = "^21.0"
aiosmtplib = "^3.0"

# Monitoring
sentry-sdk = "^1.40"

[tool.ruff]
line-length = 100
target-version = "py311"
select = ["E", "F", "I", "N", "W", "UP"]

[tool.black]
line-length = 100
target-version = ["py311"]

[tool.mypy]
python_version = "3.11"
strict = true
ignore_missing_imports = true

[tool.pytest.ini_options]
asyncio_mode = "auto"
testpaths = ["tests"]

[build-system]
requires = ["poetry-core"]
build-backend = "poetry.core.masonry.api"
```

---

## requirements.txt (альтернатива Poetry)

```txt
# Core
fastapi>=0.109,<0.110
uvicorn[standard]>=0.27,<0.28
pydantic>=2.5,<3.0
pydantic-settings>=2.1,<3.0
python-multipart>=0.0.6,<0.1

# Google APIs
gspread>=6.0,<7.0
google-auth>=2.27,<3.0
google-api-python-client>=2.116,<3.0
google-generativeai>=0.4,<1.0

# Data processing
pandas>=2.2,<3.0
openpyxl>=3.1,<4.0

# Async
aiohttp>=3.9,<4.0
redis>=5.0,<6.0

# Config & Logging
pyyaml>=6.0,<7.0
python-dotenv>=1.0,<2.0
structlog>=24.1,<25.0
```

---

## Почему эти библиотеки?

### FastAPI vs Flask vs Django

| Критерий | FastAPI | Flask | Django |
|----------|---------|-------|--------|
| Async native | **Да** | Нет | Частично |
| Скорость | **Высокая** | Средняя | Средняя |
| Валидация | **Встроена** | Ручная | ORM-based |
| Документация | **Auto (OpenAPI)** | Ручная | Ручная |
| Размер | Минимальный | **Минимальный** | Большой |

**Выбор: FastAPI** — async из коробки, автовалидация, автодокументация.

### gspread vs google-api-python-client

| Критерий | gspread | google-api-python-client |
|----------|---------|--------------------------|
| API уровень | Высокий | Низкий |
| Простота | **Простой** | Сложный |
| Batch операции | **Поддержка** | Ручная реализация |
| Документация | **Отличная** | Техническая |

**Выбор: gspread** — проще использовать, меньше boilerplate.

### Pandas vs Polars

| Критерий | Pandas | Polars |
|----------|--------|--------|
| Экосистема | **Огромная** | Растущая |
| Скорость | Хорошая | **Отличная** |
| Memory | Больше | **Меньше** |
| Совместимость | **Везде** | Ограничена |

**Выбор: Pandas** — лучше совместимость с gspread и openpyxl.

### Redis vs In-memory cache

| Критерий | Redis | In-memory |
|----------|-------|-----------|
| Персистентность | **Да** | Нет |
| Масштабирование | **Да** | Нет |
| Скорость | Отличная | **Отличная** |
| Настройка | Требуется | **Не нужна** |

**Выбор: Redis** — для production, in-memory для development.

---

## Версии Python

**Минимум: Python 3.11**

Почему 3.11+:
- Значительно быстрее (10-60% vs 3.10)
- Лучшие error messages
- `asyncio.TaskGroup` из коробки
- Все зависимости поддерживают

---

## Обновление зависимостей

### Автоматическое (рекомендуется)

```bash
# С Poetry
poetry update

# Проверка уязвимостей
poetry audit
```

### Ручное

```bash
# Проверка outdated
poetry show --outdated

# Обновление конкретного пакета
poetry update fastapi
```

---

## Минимальная установка

Для быстрого старта без optional зависимостей:

```bash
# Только основные
poetry install --without optional

# Или с requirements
pip install -r requirements.txt
```

**Размер:** ~150 MB (с pandas и google libraries)

---

## Лицензии

Все используемые библиотеки имеют MIT/Apache/BSD лицензии — можно использовать в коммерческих проектах без ограничений.
