# 📋 Краткий справочник по логированию

## 🚀 Быстрый старт

### Запись в "Логи" лист (самый простой способ)

```python
from src.services.logging import LoggingService
from src.services.sheets import SheetsService

sheets_service = SheetsService()
logging_service = LoggingService(sheets_service)

# Одна строка - одна запись в "Логи"
logging_service.add_log(
    spreadsheet_id="13k...",
    category="СИНХРО",
    action="Синхронизация Главная→Прайс",
    details="Обновлено 5 рядов",
    status="✅ OK"
)
```

---

## 🔍 Три типа логирования

### 1️⃣ Логирование синхронизации (Сервер)

**Что?** Каждое движение данных между листами
**Где?** `/data/sync_logs/<sheet_id>.jsonl`
**Как долго?** 90 дней

```python
sync_log_service.add_entry(
    spreadsheet_id="13k...",
    source_info="Главная!E5",
    target_info="Прайс!F10",
    old_value="100",
    new_value="150",
    category="PRICE",
    status="success",
    row_key="123"
)
```

**Поиск в веб-UI**: http://localhost:8000/config/logs_manager.html?id=13k...

---

### 2️⃣ Логирование функций (Детальное)

**Что?** Каждый шаг выполнения функции
**Где?** `/data/function_logs/steps.jsonl` + `/data/function_logs/executions.jsonl`
**Как долго?** 30 дней

```python
from src.services.function_log_service import FunctionLogContext

with FunctionLogContext(service, "module_name", "function_name") as ctx:
    # Автоматически логирует start и end

    ctx.log(step_name="validate", message="Checking input")
    ctx.log(step_name="process", message="Processing data")
    ctx.log(step_name="save", message="Saving results", duration_ms=123)
```

---

### 3️⃣ Сводное логирование ("Логи" лист)

**Что?** Краткие записи для быстрого обзора в Sheets
**Где?** Google Sheets, лист "Логи"
**Как долго?** До очистки (стирается в полночь)

```python
# Синхронизация
logging_service.log_sync_summary(
    spreadsheet_id="13k...",
    source_sheet="Главная",
    target_sheet="Прайс",
    status="success"
)

# Функция
logging_service.log_function_summary(
    spreadsheet_id="13k...",
    function_name="sync_event",
    status="completed",
    step_count=7,
    duration_ms=450
)
```

---

## 🎯 Распространённые сценарии

### Сценарий 1: Записать результат синхронизации

```python
# После успешной синхронизации
logging_service.log_sync_summary(
    spreadsheet_id=ss_id,
    source_sheet=source_sheet,
    target_sheet=target_sheet,
    status="success",
    details=f"Скопировано {count} значений"
)

# При ошибке
logging_service.log_sync_summary(
    spreadsheet_id=ss_id,
    source_sheet=source_sheet,
    target_sheet=target_sheet,
    status="error",
    details=f"API error: {error_msg}"
)
```

### Сценарий 2: Отследить все шаги функции

```python
def my_complex_function(data):
    with FunctionLogContext(log_service, "my_module", "my_complex_function") as ctx:
        # Шаг 1
        ctx.log(
            step_name="load_data",
            message=f"Loaded {len(data)} items",
            status="completed"
        )

        # Шаг 2
        try:
            result = process(data)
            ctx.log(
                step_name="process",
                message="Data processed",
                status="completed",
                duration_ms=elapsed
            )
        except Exception as e:
            ctx.log(
                step_name="process",
                message="Failed to process",
                status="failed",
                error=str(e),
                level="error"
            )
            raise

        # Шаг 3
        ctx.log(
            step_name="return",
            message="Returning results",
            status="completed"
        )

        return result
```

### Сценарий 3: Запросить логи за период

```python
from datetime import datetime, timedelta

# Логи за последний час
end = datetime.utcnow()
start = end - timedelta(hours=1)

items, total = sync_log_service.list_entries(
    spreadsheet_id="13k...",
    start=start,
    end=end,
    limit=1000
)

print(f"Найдено {total} событий")
for item in items:
    print(f"  {item['timestamp']}: {item['status']}")
```

---

## 📊 API endpoints быстрого доступа

| Задача | Endpoint |
|--------|----------|
| **Получить логи** | `GET /api/v1/sync-logs/{id}?limit=100` |
| **Статистика** | `GET /api/v1/sync-logs/{id}/stats` |
| **Экспорт JSON** | `GET /api/v1/sync-logs/{id}/export?format=json` |
| **Экспорт CSV** | `GET /api/v1/sync-logs/{id}/export?format=csv` |
| **Очистить старые** | `POST /api/v1/sync-logs/{id}/truncate?keep_days=90` |
| **Истории функций** | `GET /api/v1/function-logs/executions?limit=50` |

---

## ⚠️ Типичные ошибки

❌ **НЕПРАВИЛЬНО:**
```python
# Не забудьте про контекст
function_log_service.log_step(...)  # ❌ Неправильно - нет execution_id
```

✅ **ПРАВИЛЬНО:**
```python
with FunctionLogContext(service, "module", "function") as ctx:
    ctx.log(step_name="step1", message="msg")  # ✅ Правильно
```

---

❌ **НЕПРАВИЛЬНО:**
```python
# Не пишите в "Логи" после каждого события
for item in items:
    logging_service.add_log(...)  # ❌ Слишком много записей
```

✅ **ПРАВИЛЬНО:**
```python
# Пишите одну сводку за операцию
logging_service.log_sync_summary(
    spreadsheet_id=ss_id,
    source_sheet=src,
    target_sheet=tgt,
    status="success"
)  # ✅ Одна запись, несколько событий
```

---

## 🔗 Связанные документы

- [Полная реализация](LOGS_IMPLEMENTATION.md)
- [API документация](API.md)
- [Архитектура](ARCHITECTURE.md)

---

## 💡 Советы

1. **Используйте `with` для функций** - автоматически запишет start/end
2. **Используйте `log_sync_summary()` для синхро** - проще чем добавлять поля вручную
3. **Не пишите в логи для каждого элемента** - используйте сводки
4. **Проверяйте логи через web-UI** - удобнее чем JSONL файлы
5. **Экспортируйте периодически** - для архива и анализа

---

**Версия**: 1.0
**Последнее обновление**: 2026-01-03
