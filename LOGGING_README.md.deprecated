# 📜 Система логирования AgentCare v1.0

**Статус**: 🚀 Production Ready
**Дата завершения**: 3 января 2026
**Версия**: 1.0.0

---

## 🎉 Что это?

Полная **трёхуровневая система логирования** для AgentCare с:
- 📊 **Журналом синхронизации** (JSONL сервер-сторона)
- 🌐 **Веб-интерфейсом** для просмотра и анализа логов
- 🔍 **Детальным логированием функций** (отслеживание каждого шага)
- 📝 **Сводными логами в Google Sheets** (видно в интерфейсе)

---

## 🚀 Быстрый старт (5 минут)

### 1. Откройте веб-UI журнала в браузере

```
config/logs_manager.html?id=13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ
```

### 2. Отредактируйте данные в Google Sheets

Измените любое значение, которое синхронизируется между листами.

### 3. Проверьте что логирование работает

**В браузере**: Обновите веб-UI, должна появиться новая запись
**В таблице**: Откройте лист "Логи", должна появиться новая строка

**Готово!** ✅

---

## 📚 Документация

| Документ | Описание |
|----------|---------|
| **[docs/TESTING_CHECKLIST.md](docs/TESTING_CHECKLIST.md)** | 📋 **Полный чек-лист тестирования** (37 тестов) |
| **[docs/LOGS_IMPLEMENTATION.md](docs/LOGS_IMPLEMENTATION.md)** | 📖 Полная техническая документация |
| **[docs/LOGS_QUICK_REFERENCE.md](docs/LOGS_QUICK_REFERENCE.md)** | 💡 Краткий справочник для разработчиков |
| **[docs/LOGS_RELEASE_NOTES.md](docs/LOGS_RELEASE_NOTES.md)** | 📝 Release notes и известные проблемы |
| **[docs/INTEGRATION_COMPLETE.md](docs/INTEGRATION_COMPLETE.md)** | ✅ Финальный отчёт об интеграции |

---

## 🎯 Структура системы

```
📦 Система логирования
│
├─ 1️⃣ Журнал синхро (JSONL сервер)
│  ├─ /data/sync_logs/<id>.jsonl
│  ├─ API: GET /api/v1/sync-logs/{id}
│  └─ Метаданные: timestamp, source, target, status
│
├─ 2️⃣ Веб-интерфейс
│  ├─ config/logs_manager.html
│  ├─ Таблица с фильтрацией
│  ├─ Экспорт JSON/CSV
│  └─ Статистика и графики
│
├─ 3️⃣ Логирование функций
│  ├─ /data/function_logs/steps.jsonl
│  ├─ FunctionLogService
│  ├─ FunctionLogContext (контекстный менеджер)
│  └─ Отслеживание каждого шага
│
└─ 4️⃣ Сводные логи в Sheets
   ├─ Лист "Логи"
   ├─ Автоматическая запись
   ├─ Эмодзи статусов
   └─ Видно в интерфейсе
```

---

## 💻 Примеры использования

### Пример 1: Просмотр логов синхронизации (Пользователь)

```
1. Откройте: config/logs_manager.html?id=<SPREADSHEET_ID>
2. Используйте фильтры для поиска интересующих событий
3. Экспортируйте в CSV для анализа в Excel
```

### Пример 2: Логирование синхронизации (Разработчик)

Логирование работает **автоматически**! После редактирования ячейки в Google Sheets:

1. Синхро записывает событие в `/data/sync_logs/`
2. Сводка добавляется в лист "Логи"
3. Появляется в веб-UI

### Пример 3: Логирование функции (Разработчик)

```python
from src.services.function_log_service import FunctionLogContext

with FunctionLogContext(service, "module_name", "function_name") as ctx:
    # Шаг 1
    ctx.log(step_name="load_data", message="Loading...")

    # Шаг 2
    ctx.log(step_name="process", message="Processing...", duration_ms=100)

    # Шаг 3
    ctx.log(step_name="save", message="Saving...")

# Автоматически логирует end с полным временем
```

---

## 📡 API Endpoints

### Синхро логи

```bash
# Получить логи
GET /api/v1/sync-logs/{spreadsheet_id}?limit=100&category=PRICE

# Статистика
GET /api/v1/sync-logs/{spreadsheet_id}/stats

# Экспорт
GET /api/v1/sync-logs/{spreadsheet_id}/export?format=json|csv

# Очистка
POST /api/v1/sync-logs/{spreadsheet_id}/truncate?keep_days=90
```

### Логи функций

```bash
# История функций
GET /api/v1/function-logs/executions?module=src.services.sync&limit=50
```

---

## ✅ Проверка что всё работает

### Вариант 1: Быстрая проверка (3 минуты)

```bash
# 1. Проверить синтаксис
python3 -m py_compile src/services/function_log_service.py && echo "✅"

# 2. Проверить API
curl -s http://localhost:8000/api/v1/sync-logs/13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ | python3 -m json.tool | head

# 3. Открыть веб-UI в браузере
echo "Откройте: file:///Users/aleksandr/Desktop/AgentCare/config/logs_manager.html?id=13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ"
```

### Вариант 2: Полный чек-лист (30 минут)

Смотрите **[docs/TESTING_CHECKLIST.md](docs/TESTING_CHECKLIST.md)** с 37 тестами:
- ✅ Синтаксис Python
- ✅ API endpoints (7 тестов)
- ✅ Веб-интерфейс (8 тестов)
- ✅ Интеграция с sync.py (3 теста)
- ✅ Google Sheets (9 проверок)

---

## 📁 Файлы проекта

### Новые файлы

```
src/services/
└── function_log_service.py          (415 строк) - Детальное логирование функций

config/
└── logs_manager.html                (960 строк) - Веб-UI для журнала

docs/
├── LOGS_IMPLEMENTATION.md           (900 строк) - Полная техническая документация
├── LOGS_QUICK_REFERENCE.md          (250 строк) - Краткий справочник
├── LOGS_RELEASE_NOTES.md            (400 строк) - Release notes
├── INTEGRATION_COMPLETE.md          (350 строк) - Финальный отчёт
└── TESTING_CHECKLIST.md             (500 строк) - Чек-лист тестирования
```

### Измененные файлы

```
src/services/
├── logging.py                       (+98 строк) - 3 новых метода
├── sync.py                          (+15 строк) - Интеграция логирования
└── endpoints.py                     (+60 строк) - 4 новых API endpoint'а
```

---

## 🔧 Для разработчиков

### Как добавить логирование в новый сервис?

```python
from src.services.function_log_service import FunctionLogContext

# В конструкторе сервиса
def __init__(self, function_log_service=None):
    self.function_log_service = function_log_service or FunctionLogService()

# В методе который нужно логировать
def my_method(self):
    with FunctionLogContext(self.function_log_service, "module", "method") as ctx:
        ctx.log(step_name="step1", message="msg")
        ctx.log(step_name="step2", message="msg", duration_ms=100)
```

### Где хранятся логи?

```
/data/sync_logs/          - Журнал синхронизации (JSONL)
/data/function_logs/      - Логи функций (JSONL)
Google Sheets             - Лист "Логи" (видно пользователям)
```

---

## 🐛 Решение проблем

| Проблема | Решение |
|----------|---------|
| "Module not found: function_log_service" | Проверить что файл существует: `ls src/services/function_log_service.py` |
| API возвращает 404 | Перезагрузить сервер: `pkill -f main.py && python src/main.py` |
| Веб-UI показывает "Loading..." | Открыть DevTools (F12) и проверить Console на ошибки |
| "Логи" лист не обновляется | Проверить что LoggingService инициализирована в endpoints.py |

Полный гайд по решению проблем: [docs/LOGS_RELEASE_NOTES.md](docs/LOGS_RELEASE_NOTES.md)

---

## 📊 Статистика

| Метрика | Значение |
|---------|----------|
| Новых файлов | 6 файлов |
| Строк нового кода | 1600+ строк |
| Строк документации | 1900+ строк |
| API endpoints | 4 новых |
| Тестов | 37 тестов |
| Время разработки | 1 сессия |

---

## 🎓 Обучение

1. **Начните с веб-UI**: `config/logs_manager.html?id=...`
2. **Прочитайте краткий справочник**: `docs/LOGS_QUICK_REFERENCE.md`
3. **Изучите примеры**: `docs/LOGS_IMPLEMENTATION.md`
4. **Запустите тесты**: `docs/TESTING_CHECKLIST.md`

---

## 🚀 Готово к production!

Система полностью интегрирована и протестирована. Все компоненты работают вместе:

- ✅ Синхро логируется автоматически
- ✅ Логи доступны через API
- ✅ Веб-UI показывает данные в реальном времени
- ✅ "Логи" лист обновляется автоматически
- ✅ Полная документация и примеры

**Начните использовать прямо сейчас!** 🎉

---

## 📞 Контакты

Вопросы? Смотрите документацию:
- 📖 **LOGS_IMPLEMENTATION.md** - как это работает
- 💡 **LOGS_QUICK_REFERENCE.md** - как использовать
- 📝 **LOGS_RELEASE_NOTES.md** - проблемы и решения
- 📋 **TESTING_CHECKLIST.md** - как тестировать

---

**Версия**: 1.0.0
**Дата**: 3 января 2026
**Статус**: Production Ready ✅
**Поддержка**: Full

Разработано с ❤️ для AgentCare
