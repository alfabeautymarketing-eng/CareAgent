# 📜 Система логирования AgentCare - Полная реализация

**Дата**: 3 января 2026
**Статус**: ✅ Завершено

---

## 📋 Краткое содержание

Реализована полная трёхуровневая система логирования для AgentCare:

| Уровень | Компонент | Назначение |
|---------|-----------|-----------|
| **1. Журнал синхро** | Сервер-сторона (JSONL) | Запись каждого события синхронизации с полной информацией о передаче данных |
| **2. Логи функций** | Детальное отслеживание | Запись каждого шага выполнения функции с временем и статусом |
| **3. Сводные логи** | Google Sheets | Краткие записи в лист "Логи" для быстрого обзора в интерфейсе |

---

## 🎯 AgentCare-1zy.1: Журнал синхро (Сервер-сторона)

### ✅ Реализовано

**Сервис**: `src/services/sync_log_service.py`

- **Хранилище**: `/data/sync_logs/<spreadsheet_id>.jsonl`
- **Формат**: JSONL (одна запись JSON на строку)
- **Кодировка**: UTF-8
- **Блокировка**: FileLock для потокобезопасности

### Структура записи SyncLogEntry

```python
{
  "id": "uuid_hex",
  "timestamp": "2026-01-03T12:34:56Z",
  "spreadsheet_id": "13k...",
  "project": "mt",
  "row_key": "123",
  "source_info": "Главная!E5",
  "target_info": "Прайс!F10",
  "old_value": "100",
  "new_value": "150",
  "category": "PRICE",
  "status": "success",
  "event": "cell_changed",
  "tags": ["price_adjustment"],
  "rule_id": "rule_123",
  "extra": {}
}
```

### API Endpoints

#### Получение записей
```bash
GET /api/v1/sync-logs/{spreadsheet_id}
  ?limit=200
  &offset=0
  &order=desc
  &project=mt
  &category=PRICE
  &status=success
  &row_key=123
  &source=Главная
  &target=Прайс
  &rule_id=rule_123
  &start=2026-01-01T00:00:00Z
  &end=2026-01-03T23:59:59Z
```

**Ответ**:
```json
{
  "items": [...],
  "total": 1234,
  "limit": 200,
  "offset": 0
}
```

#### Статистика
```bash
GET /api/v1/sync-logs/{spreadsheet_id}/stats
```

**Ответ**:
```json
{
  "total": 1234,
  "by_category": {
    "PRICE": 800,
    "NAME": 434
  },
  "by_status": {
    "success": 1200,
    "error": 34
  },
  "by_rule_id": {
    "rule_123": 567,
    ...
  },
  "latest_timestamp": "2026-01-03T14:30:00Z",
  "oldest_timestamp": "2025-10-06T10:00:00Z"
}
```

#### Экспорт
```bash
# JSON
GET /api/v1/sync-logs/{spreadsheet_id}/export?format=json&limit=10000

# CSV
GET /api/v1/sync-logs/{spreadsheet_id}/export?format=csv&limit=10000
```

#### Очистка старых логов
```bash
POST /api/v1/sync-logs/{spreadsheet_id}/truncate
  ?keep_days=0  # 0=удалить всё, 7/30/90=оставить последние N дней
```

**Ответ**:
```json
{
  "deleted": 50,
  "remaining": 1184
}
```

---

## 🎯 AgentCare-1zy.2: UI журнала синхро

### ✅ Реализовано

**Файл**: `config/logs_manager.html`

### Функциональность

- 📊 **Статистика панель** - быстрый обзор
  - Общее количество записей
  - Успешных транзакций
  - Ошибок
  - Уникальных категорий

- 🔍 **Фильтрация**
  - Поиск по полю/ID
  - Фильтр по категориям
  - Фильтр по статусу (Успех/Ошибка/Ожидание)
  - Контроль количества результатов

- 📋 **Таблица с пагинацией**
  - Время события (дата + время)
  - Объект (row_key)
  - Категория (badge)
  - Поле источника
  - Было → Стало (old_value → new_value)
  - Статус (color-coded)

- ⬇️ **Экспорт**
  - JSON (полные данные)
  - CSV (табличный формат)
  - Контроль количества экспортируемых записей

- 🗑️ **Очистка**
  - Удалить все
  - Оставить последние 7/30/90 дней
  - Требует подтверждения

### Использование

```html
<!-- Откройте в браузере -->
file:///path/to/config/logs_manager.html?id=13k...

<!-- Или через меню AgentCare -->
Логи > Журнал синхро
```

---

## 🎯 AgentCare-1zy.3: Детальные логи функций

### ✅ Реализовано

**Сервис**: `src/services/function_log_service.py`

### Концепция

Каждый вызов функции получает уникальный `execution_id`, и внутри функции логируются отдельные шаги:

```
execution_id: "abc123def456"
├── step 0: "start" - Начало функции
├── step 1: "validate_event" - Валидация входа
├── step 2: "acquire_lock" - Получение блокировки
├── step 3: "read_source" - Чтение источника
├── step 4: "transform_data" - Трансформация
├── step 5: "write_target" - Запись в цель
└── step 6: "end" - Завершение (success/failed)
```

### Структура FunctionStep

```python
{
  "id": "uuid_hex",
  "execution_id": "abc123def456",
  "timestamp": "2026-01-03T12:34:56.789Z",

  # Function context
  "module": "src.services.sync",
  "function_name": "sync_event",

  # Step details
  "step_number": 3,
  "step_name": "read_source",
  "level": "info",  # info, warning, error, debug
  "message": "Reading from Главная!E5",
  "status": "completed",  # started, completed, failed, skipped
  "duration_ms": 45.2,

  # Context
  "row_key": "123",
  "spreadsheet_id": "13k...",
  "rule_id": "rule_123",

  # Data
  "data": {
    "cell_value": "150",
    "rows_affected": 1
  },
  "error": null
}
```

### API Endpoints

#### Список выполненных функций
```bash
GET /api/v1/function-logs/executions
  ?module=src.services.sync
  ?function_name=sync_event
  ?limit=100
  ?offset=0
```

**Ответ**:
```json
{
  "executions": [
    {
      "execution_id": "abc123",
      "status": "completed",
      "total_duration_ms": 234.5,
      "step_count": 7,
      "timestamp": "2026-01-03T12:34:56Z",
      "steps": [...]
    }
  ],
  "total": 500,
  "limit": 100,
  "offset": 0
}
```

### Использование в коде

#### Способ 1: Context Manager (рекомендуется)

```python
from src.services.function_log_service import FunctionLogService, FunctionLogContext

service = FunctionLogService()

with FunctionLogContext(service, "src.services.sync", "sync_event") as ctx:
    # Автоматически логирует "start"

    # Шаг 1
    try:
        ctx.log(
            step_name="validate_event",
            message="Validating incoming event",
            status="completed",
            data={"event_type": "cell_changed"}
        )
    except Exception as e:
        ctx.log(
            step_name="validate_event",
            message="Validation failed",
            status="failed",
            error=str(e),
            level="error"
        )

    # Шаг 2
    ctx.log(
        step_name="read_source",
        message="Reading from Главная!E5",
        status="completed",
        duration_ms=45.2,
        data={"cell_value": "150"}
    )

# Автоматически логирует "end" с полным временем выполнения
```

#### Способ 2: Ручной контроль

```python
service = FunctionLogService()
exec_id = service.start_execution()

# ... код функции ...

service.log_step(
    execution_id=exec_id,
    module="src.services.sync",
    function_name="sync_event",
    step_number=1,
    step_name="validate_event",
    message="Validating incoming event",
    status="completed"
)

# ... ещё код ...

service.end_execution(exec_id, total_duration_ms=234.5, status="completed")
```

### Хранилище

- **Шаги**: `data/function_logs/steps.jsonl` - все шаги всех функций
- **Выполнения**: `data/function_logs/executions.jsonl` - итоговые записи функций
- **Кодировка**: UTF-8
- **Блокировка**: FileLock
- **Удержание**: 30 дней

---

## 🎯 AgentCare-1zy.4: Сводные логи в "Логи" лист

### ✅ Реализовано

**Сервис**: `src/services/logging.py` (расширен)

### Структура листа "Логи"

| Столбец | Содержимое |
|---------|-----------|
| A | 🕒 Время (дата + время) |
| B | 🏷️ Категория (СИНХРО, ФУНКЦИЯ, СИСТЕМА, ...) |
| C | 💬 Действие (описание) |
| D | 📝 Детали (дополнительная информация) |
| E | 🔘 Статус (✅ OK, ⚠️ ОШИБКА, ⏳ ОЖИДАНИЕ) |

### Пример записей

```
Время                   | Категория | Действие                    | Детали                      | Статус
20.12.2025 10:30:45    | СИСТЕМА   | Сессия открыта             | Пользователь открыл        | ✅ OK
20.12.2025 10:30:50    | СИНХРО    | Синхро: Главная → Прайс   | Шаги: 3 | Время: 234ms    | ✅ OK
20.12.2025 10:31:00    | ФУНКЦИЯ   | sync_event                 | Шаги: 7 | Время: 456ms    | ✅ OK
20.12.2025 10:31:10    | СИНХРО    | Синхро: Прайс → Архив      | Ошибка: API timeout        | ⚠️ ОШИБКА
```

### Методы записи сводных логов

#### Запись вручную
```python
logging_service.add_summary_log(
    spreadsheet_id="13k...",
    category="СИНХРО",
    action="Синхро: Главная → Прайс",
    details="Обновлено 5 строк",
    status="✅ OK"
)
```

#### Логирование синхронизации
```python
logging_service.log_sync_summary(
    spreadsheet_id="13k...",
    source_sheet="Главная",
    target_sheet="Прайс",
    status="success",
    details="Обновлено 5 строк",
    rows_affected=5
)
```

#### Логирование функции
```python
logging_service.log_function_summary(
    spreadsheet_id="13k...",
    function_name="sync_event",
    status="completed",
    step_count=7,
    duration_ms=456.2
)
```

### Логирование при ошибке
```python
logging_service.log_function_summary(
    spreadsheet_id="13k...",
    function_name="sync_event",
    status="failed",
    step_count=5,
    duration_ms=150.0,
    error="API timeout after 10s"
)
```

---

## 📚 Практические примеры

### Пример 1: Логирование синхронизации в sync_service.py

```python
from src.services.sync_log_service import SyncLogService
from src.services.logging import LoggingService

class SyncService:
    def __init__(self, logging_service, sync_log_service):
        self.logging_service = logging_service
        self.sync_log_service = sync_log_service

    def sync_event(self, spreadsheet_id: str, event_data: dict):
        exec_id = f"sync_{uuid4().hex}"

        # Запись в серверный журнал синхро
        self.sync_log_service.add_entry(
            spreadsheet_id=spreadsheet_id,
            source_info=f"{event_data['sheet']}!{event_data['cell']}",
            target_info=f"{target_sheet}!{target_cell}",
            old_value=event_data.get('old_value'),
            new_value=event_data.get('new_value'),
            category="PRICE",
            status="success",
            row_key=event_data.get('row'),
            rule_id=rule_id
        )

        # Запись в сводный лист "Логи"
        self.logging_service.log_sync_summary(
            spreadsheet_id=spreadsheet_id,
            source_sheet=event_data['sheet'],
            target_sheet=target_sheet,
            status="success",
            details=f"Обновлено значение с {old_value} на {new_value}",
            rows_affected=1
        )
```

### Пример 2: Детальное логирование функции

```python
from src.services.function_log_service import FunctionLogContext

async def process_price(spreadsheet_id: str, rules: List[Rule]):
    with FunctionLogContext(function_log_service, "src.services", "process_price") as ctx:
        # Шаг 1: Загрузка данных
        ctx.log(
            step_name="load_data",
            message="Loading product data from Sheets",
            status="completed",
            duration_ms=120.5,
            data={"rows_loaded": 500}
        )

        # Шаг 2: Валидация
        invalid_rows = validate_prices(products)
        if invalid_rows:
            ctx.log(
                step_name="validate_prices",
                message=f"Found {len(invalid_rows)} invalid prices",
                status="completed",
                level="warning",
                data={"invalid_count": len(invalid_rows)}
            )

        # Шаг 3: Обработка
        processed = await apply_rules(products, rules)
        ctx.log(
            step_name="apply_rules",
            message="Applied pricing rules",
            status="completed",
            duration_ms=850.2,
            data={"rules_applied": len(rules), "items_processed": len(processed)}
        )

        # Шаг 4: Загрузка результатов
        ctx.log(
            step_name="write_results",
            message="Writing results to Sheets",
            status="completed",
            duration_ms=340.1
        )

        # Логирование в сводный лист
        logging_service.log_function_summary(
            spreadsheet_id=spreadsheet_id,
            function_name="process_price",
            status="completed",
            step_count=4,
            duration_ms=1310.8
        )
```

---

## 🔍 Запросы к логам

### Получить все ошибки синхронизации за сегодня

```python
from datetime import datetime, timedelta

today_start = datetime.utcnow().replace(hour=0, minute=0, second=0, microsecond=0)
today_end = datetime.utcnow()

items, total = sync_log_service.list_entries(
    spreadsheet_id="13k...",
    status="error",
    start=today_start,
    end=today_end,
    limit=1000
)

for item in items:
    print(f"{item['timestamp']}: {item['source_info']} → {item['target_info']}: {item['error']}")
```

### Получить статистику по категориям

```python
stats = sync_log_service.list_entries(
    spreadsheet_id="13k...",
    limit=10000
)[0]  # Получаем items

from collections import Counter
categories = Counter(item.get('category') for item in stats)
print(f"Категории: {dict(categories)}")
```

### Найти все синхро для конкретного ряда

```python
items, total = sync_log_service.list_entries(
    spreadsheet_id="13k...",
    row_key="123",
    limit=100
)

print(f"Найдено {total} операций для строки 123")
for item in items:
    print(f"  {item['timestamp']}: {item['old_value']} → {item['new_value']} ({item['status']})")
```

---

## 📊 Архитектура системы

```
┌─────────────────────────────────────────────────────────────┐
│                   Google Sheets (UI)                         │
│  ┌──────────────────────────────────────────────────────┐   │
│  │ Лист "Логи"                                          │   │
│  │ 🕒 Время | 🏷️ Категория | 💬 Действие | 📝 Детали  │   │
│  │ (Сводные записи - краткий обзор)                     │   │
│  └──────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────┘
                            ↑
                     (write_to_sheet)
                            │
┌─────────────────────────────────────────────────────────────┐
│              Python Backend (API)                            │
│  ┌────────────────────┐  ┌────────────────────────────┐    │
│  │ LoggingService     │  │ FunctionLogService         │    │
│  │ - add_log()        │  │ - log_step()               │    │
│  │ - log_sync_summary │  │ - start_execution()        │    │
│  │ - log_function_sum │  │ - end_execution()          │    │
│  └────────────────────┘  └────────────────────────────┘    │
│  /api/v1/logs/write (/api/v1/function-logs/executions)    │
└─────────────────────────────────────────────────────────────┘
                            ↑
                   (append rows)
                            │
┌─────────────────────────────────────────────────────────────┐
│           Storage Layer (Files + JSONL)                      │
│  ┌──────────────────┐  ┌──────────────────────────┐         │
│  │ Sync Logs        │  │ Function Logs            │         │
│  │ (JSONL)          │  │ (JSONL)                  │         │
│  │                  │  │                          │         │
│  │ /data/sync_logs/ │  │ /data/function_logs/     │         │
│  │ <sheet_id>.jsonl │  │ steps.jsonl              │         │
│  │                  │  │ executions.jsonl         │         │
│  └──────────────────┘  └──────────────────────────┘         │
└─────────────────────────────────────────────────────────────┘
                            ↑
                  (query/export)
                            │
┌─────────────────────────────────────────────────────────────┐
│            UI Layer (Web Dashboard)                          │
│  ┌────────────────────────────────────────────────────┐    │
│  │ logs_manager.html                                  │    │
│  │ - Statistics panel                                 │    │
│  │ - Advanced filters                                 │    │
│  │ - Paginated table view                            │    │
│  │ - Export (JSON/CSV)                               │    │
│  │ - Truncate tools                                  │    │
│  └────────────────────────────────────────────────────┘    │
│  /api/v1/sync-logs/{id}                                    │
│  /api/v1/sync-logs/{id}/stats                              │
│  /api/v1/sync-logs/{id}/export                             │
│  /api/v1/sync-logs/{id}/truncate                           │
└─────────────────────────────────────────────────────────────┘
```

---

## 🚀 Следующие шаги

### Интеграция в синхро-сервис

```python
# В sync.py добавить вызовы логирования:

from src.services.function_log_service import FunctionLogContext

def sync_event(self, ...):
    with FunctionLogContext(self.function_log_service, ...) as ctx:
        ctx.log(step_name="validate", ...)
        ctx.log(step_name="sync", ...)
        self.logging_service.log_sync_summary(...)
```

### Интеграция в UI меню

```javascript
// В 08StatusManager.js добавить кнопку:
{
  name: "📜 Логи",
  functionName: "openLogsManager",
  url: "/config/logs_manager.html"
}
```

### Мониторинг и оповещения

Можно добавить:
- WebSocket для real-time обновлений логов
- Email уведомления при критических ошибках
- Telegram alerts для SLA нарушений

---

## 📞 Поддержка

### Возможные проблемы

**P1: Логи не пишутся в "Логи" лист**
- Проверьте сеть и права на таблицу
- Убедитесь, что лист "Логи" создан
- Проверьте серверные логи: `grep "log_append_failed" logs/`

**P2: Файлы JSONL становятся очень большими**
- Ручная очистка: `POST /api/v1/sync-logs/{id}/truncate?keep_days=90`
- Добавить в cron: `0 0 * * * curl -X POST /api/v1/sync-logs/*/truncate?keep_days=90`

**P3: UI не загружается**
- Проверьте URL: `file:///path/to/config/logs_manager.html?id=<SPREADSHEET_ID>`
- Очистите кэш браузера

---

**Версия**: 1.0
**Последнее обновление**: 2026-01-03
**Автор**: Claude Code
