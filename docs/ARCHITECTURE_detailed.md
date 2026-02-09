# AgentCare - Архитектура системы

## Обзор

AgentCare — Python-сервис для автоматизации работы с Google Sheets, синхронизации данных между листами и AI-анализа продуктов (INCI, ТН ВЭД, классификация).

**Замена для:** MyGoogleScripts (Google Apps Script)

---

## Ключевые принципы

1. **Модульность** — каждый компонент в отдельном файле/папке
2. **Независимость** — модули связаны через интерфейсы, не напрямую
3. **Тестируемость** — каждый модуль можно тестировать изолированно
4. **Конфигурируемость** — все настройки в config, не в коде
5. **Отказоустойчивость** — retry, транзакции, откаты

---

## Структура проекта

```
AgentCare/
├── docs/                       # Документация
│   ├── ARCHITECTURE.md         # Этот файл
│   ├── MODULES.md              # Описание модулей
│   ├── API.md                  # API спецификация
│   ├── DEPENDENCIES.md         # Зависимости
│   ├── DEPLOYMENT.md           # Деплой и настройка
│   └── MIGRATION.md            # План миграции с GAS
│
├── config/                     # Конфигурация
│   ├── settings.yaml           # Основные настройки
│   ├── projects/               # Настройки проектов
│   │   ├── mt.yaml             # MT (CosmeticaBar)
│   │   ├── sk.yaml             # SK (Carmado)
│   │   └── ss.yaml             # SS (Сан)
│   └── rules/                  # Правила синхронизации
│       └── cascade_rules.yaml  # Каскадные правила
│
├── src/                        # Исходный код
│   ├── __init__.py
│   ├── main.py                 # Точка входа
│   │
│   ├── api/                    # HTTP API (webhooks)
│   │   ├── __init__.py
│   │   ├── router.py           # FastAPI роутер
│   │   ├── webhooks.py         # Обработчики webhook
│   │   └── middleware.py       # Middleware (auth, logging)
│   │
│   ├── core/                   # Ядро системы
│   │   ├── __init__.py
│   │   ├── sync_engine.py      # Движок синхронизации
│   │   ├── rules_engine.py     # Движок правил
│   │   ├── transaction.py      # Транзакции и откаты
│   │   └── scheduler.py        # Планировщик задач
│   │
│   ├── services/               # Бизнес-логика
│   │   ├── __init__.py
│   │   ├── price_processor.py  # Обработка прайсов
│   │   ├── sync_service.py     # Сервис синхронизации
│   │   ├── invoice_service.py  # Работа с инвойсами
│   │   └── order_service.py    # Автозаказ
│   │
│   ├── parsers/                # Парсеры форматов
│   │   ├── __init__.py
│   │   ├── base_parser.py      # Базовый класс парсера
│   │   ├── mt_parser.py        # Парсер MT формата
│   │   ├── sk_parser.py        # Парсер SK формата
│   │   └── ss_parser.py        # Парсер SS формата
│   │
│   ├── integrations/           # Внешние интеграции
│   │   ├── __init__.py
│   │   ├── google_sheets.py    # Google Sheets API
│   │   ├── google_drive.py     # Google Drive API
│   │   ├── gemini_client.py    # Gemini AI API
│   │   └── notification.py     # Уведомления (Telegram/Email)
│   │
│   └── utils/                  # Утилиты
│       ├── __init__.py
│       ├── cache.py            # Кэширование
│       ├── logger.py           # Логирование
│       ├── retry.py            # Retry логика
│       └── validators.py       # Валидаторы данных
│
├── tests/                      # Тесты
│   ├── __init__.py
│   ├── conftest.py             # Pytest fixtures
│   ├── unit/                   # Unit тесты
│   ├── integration/            # Интеграционные тесты
│   └── fixtures/               # Тестовые данные
│
├── scripts/                    # Скрипты
│   ├── setup_google_auth.py    # Настройка OAuth
│   ├── migrate_data.py         # Миграция данных
│   └── health_check.py         # Проверка здоровья
│
├── .env.example                # Пример переменных окружения
├── .gitignore
├── Dockerfile
├── docker-compose.yml
├── pyproject.toml              # Poetry конфиг
├── Makefile                    # Команды управления
└── README.md
```

---

## Компоненты системы

### 1. API Layer (`src/api/`)

**Назначение:** Приём webhooks от Google Sheets и внешних запросов

```
┌─────────────────┐     ┌─────────────────┐
│  Google Sheets  │────▶│   FastAPI       │
│  (onChange)     │     │   Webhooks      │
└─────────────────┘     └────────┬────────┘
                                 │
                                 ▼
                        ┌─────────────────┐
                        │  Task Queue     │
                        │  (Background)   │
                        └─────────────────┘
```

**Endpoints:**
- `POST /webhook/sheets/{project}` — событие из Google Sheets
- `POST /webhook/sync` — ручной запуск синхронизации
- `GET /health` — проверка здоровья
- `GET /status/{task_id}` — статус задачи

### 2. Core Layer (`src/core/`)

**Назначение:** Основная логика синхронизации и правил

#### SyncEngine
Движок синхронизации между листами:
- Получает изменения
- Определяет затронутые листы
- Применяет правила
- Выполняет обновления

#### RulesEngine
Движок каскадных правил:
- Загружает правила из YAML
- Применяет правила к данным
- Поддерживает условия и зависимости

#### Transaction
Управление транзакциями:
- Группирует операции
- Откат при ошибке
- Логирование изменений

### 3. Services Layer (`src/services/`)

**Назначение:** Бизнес-логика конкретных операций

| Сервис | Функция |
|--------|---------|
| `price_processor` | 9 фаз обработки прайса |
| `sync_service` | Синхронизация между листами |
| `invoice_service` | Генерация инвойсов |
| `order_service` | Автозаказ |

### 4. Parsers Layer (`src/parsers/`)

**Назначение:** Парсинг разных форматов Excel

```python
# Базовый интерфейс
class BaseParser:
    def parse(self, file_path: str) -> DataFrame
    def validate(self, data: DataFrame) -> ValidationResult
    def transform(self, data: DataFrame) -> DataFrame
```

Каждый проект (MT/SK/SS) имеет свой парсер с логикой извлечения полей.

### 5. Integrations Layer (`src/integrations/`)

**Назначение:** Работа с внешними API

| Интеграция | Библиотека | Функции |
|------------|------------|---------|
| Google Sheets | `gspread` | CRUD операции, batch updates |
| Google Drive | `google-api-python-client` | Файлы, папки, конвертация |
| Gemini AI | `google-generativeai` | INCI анализ, классификация |
| Notifications | `python-telegram-bot` | Алерты об ошибках |

---

## Потоки данных

### Поток 1: Webhook от Google Sheets

```
1. Google Sheets → onChange триггер
2. Apps Script → POST /webhook/sheets/mt
3. FastAPI → валидация, создание задачи
4. Background Worker → SyncEngine.process()
5. SyncEngine → RulesEngine.apply()
6. GoogleSheetsClient → batch update
7. Response → 200 OK + task_id
```

### Поток 2: Обработка прайса

```
1. Запрос: POST /api/price/process
2. PriceProcessor.run(project="mt")
   ├── Фаза 1: load_excel()
   ├── Фаза 2: clear_columns()
   ├── Фаза 3: sync_with_base()
   ├── Фаза 4: save_snapshot()
   ├── Фаза 5: copy_prices()
   ├── Фаза 6: fill_ids()
   ├── Фаза 7: process_new_articles()
   ├── Фаза 8: apply_formulas()
   └── Фаза 9: update_statuses()
3. Transaction.commit() или rollback()
4. Notification.send(result)
```

### Поток 3: AI анализ

```
1. Запрос: POST /api/ai/analyze
2. GeminiClient.analyze_inci(pdf_url)
   ├── Drive: конвертация PDF → текст
   ├── Gemini: извлечение данных
   └── Validation: проверка результата
3. SyncService.update_row(data)
4. Response: результат анализа
```

---

## Конфигурация

### Уровни конфигурации

```
1. .env                    # Секреты (API ключи)
2. config/settings.yaml    # Общие настройки
3. config/projects/*.yaml  # Настройки проектов
4. config/rules/*.yaml     # Правила синхронизации
```

### Пример settings.yaml

```yaml
app:
  name: AgentCare
  debug: false
  log_level: INFO

server:
  host: 0.0.0.0
  port: 8000
  workers: 4

cache:
  backend: redis  # или "memory"
  ttl: 3600
  redis_url: ${REDIS_URL}

google:
  credentials_file: ${GOOGLE_CREDENTIALS}
  scopes:
    - https://www.googleapis.com/auth/spreadsheets
    - https://www.googleapis.com/auth/drive

gemini:
  api_key: ${GEMINI_API_KEY}
  model: gemini-2.5-flash
  max_retries: 5
  retry_delays: [2, 5, 10, 20, 30]

sync:
  batch_size: 100
  timeout: 300
  lock_timeout: 30
```

---

## Обработка ошибок

### Уровни

1. **Retry** — автоматический повтор при временных ошибках
2. **Fallback** — альтернативное действие при постоянной ошибке
3. **Rollback** — откат транзакции при критической ошибке
4. **Alert** — уведомление при неразрешимой ошибке

### Пример

```python
@retry(max_attempts=3, delay=2, backoff=2)
@transaction
async def sync_sheets(source: str, target: str):
    try:
        data = await sheets.read(source)
        transformed = rules.apply(data)
        await sheets.write(target, transformed)
    except QuotaExceeded:
        await notify.alert("Quota exceeded, waiting...")
        raise RetryLater(delay=60)
    except ValidationError as e:
        raise RollbackError(f"Invalid data: {e}")
```

---

## Масштабирование

### Текущий дизайн (1 сервер)

```
┌─────────────────────────────────────┐
│           VPS (4 CPU, 8GB)          │
│  ┌─────────┐  ┌─────────┐           │
│  │ FastAPI │  │ Worker  │           │
│  └─────────┘  └─────────┘           │
│       │            │                │
│       └─────┬──────┘                │
│             │                       │
│       ┌─────▼─────┐                 │
│       │   Redis   │                 │
│       └───────────┘                 │
└─────────────────────────────────────┘
```

### Будущее (если нужно)

```
┌──────────────┐     ┌──────────────┐
│   FastAPI    │     │   FastAPI    │
│   Instance 1 │     │   Instance 2 │
└──────┬───────┘     └──────┬───────┘
       │                    │
       └────────┬───────────┘
                │
         ┌──────▼──────┐
         │ Redis Queue │
         └──────┬──────┘
                │
    ┌───────────┼───────────┐
    │           │           │
┌───▼───┐  ┌───▼───┐  ┌───▼───┐
│Worker1│  │Worker2│  │Worker3│
└───────┘  └───────┘  └───────┘
```

---

## Мониторинг

### Метрики

- Количество обработанных запросов
- Время выполнения синхронизации
- Количество ошибок по типам
- Использование квот Google API
- Размер очереди задач

### Логирование

```python
# Структурированные логи
logger.info("sync_completed", extra={
    "project": "mt",
    "sheets_updated": 5,
    "rows_changed": 150,
    "duration_ms": 2340
})
```

### Алерты

- Telegram/Email при критических ошибках
- Webhook при завершении длительных операций

---

## Безопасность

1. **Секреты** — только через переменные окружения
2. **OAuth** — Google Cloud проект с ограниченными scopes
3. **Webhook auth** — подпись запросов (HMAC)
4. **Rate limiting** — защита от злоупотреблений
5. **Audit log** — логирование всех изменений

---

## Следующие шаги

1. [MODULES.md](MODULES.md) — детальное описание каждого модуля
2. [API.md](API.md) — спецификация API endpoints
3. [DEPENDENCIES.md](DEPENDENCIES.md) — зависимости и их версии
4. [DEPLOYMENT.md](DEPLOYMENT.md) — инструкция по деплою
5. [MIGRATION.md](MIGRATION.md) — план миграции с Google Apps Script
