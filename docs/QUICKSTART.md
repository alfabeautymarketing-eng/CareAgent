# ⚡ Быстрый старт - 5 минут

**Последовать этому руководству**: ~5 минут
**Статус**: ✅ Production Ready

---

## 🚀 Шаг 1: Откройте веб-интерфейс (1 минута)

Скопируйте URL в браузер:

```
file:///Users/aleksandr/Desktop/AgentCare/config/logs_manager.html?id=13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ
```

**Что вы увидите:**
- 📊 Статистика: всего записей, успешных, ошибок
- 📋 Таблица с логами синхронизации
- 🔍 Фильтры и поиск
- ⬇️ Кнопки для экспорта

---

## 🔄 Шаг 2: Отредактируйте данные (2 минуты)

1. Откройте таблицу Google Sheets: https://docs.google.com/spreadsheets/d/13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ

2. **Отредактируйте любую ячейку** (например, цену в листе "Главная")

3. Синхронизация произойдёт автоматически

---

## 📜 Шаг 3: Проверьте логи (2 минуты)

### В браузере (веб-UI)

1. Обновите страницу (F5)
2. Должна появиться новая запись в таблице
3. Попробуйте фильтровать по категориям

### В Google Sheets

1. Откройте таблицу
2. Перейдите на лист **"Логи"**
3. Должна появиться новая строка с деталями синхронизации

**Пример записи:**
```
Время              | Категория | Действие              | Детали        | Статус
03.01.2026 12:45  | СИНХРО    | Синхро: Главная→Прайс | Обработано... | ✅ OK
```

---

## ✅ Готово!

Вы успешно использовали систему логирования! 🎉

---

## 📚 Дальше - для разработчиков

### Если нужно добавить логирование в код:

```python
# Автоматическое логирование синхро
logging_service.log_sync_summary(
    spreadsheet_id="13kB...",
    source_sheet="Главная",
    target_sheet="Прайс",
    status="success"
)

# Логирование функции
from src.services.function_log_service import FunctionLogContext

with FunctionLogContext(service, "module", "function") as ctx:
    ctx.log(step_name="step", message="msg")
```

### API запрос:

```bash
curl "http://localhost:8000/api/v1/sync-logs/13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ?limit=10"
```

---

## 🆘 Если что-то не работает

### "Веб-UI показывает Loading..."

- Откройте DevTools (F12)
- Проверьте Console на ошибки
- Убедитесь что сервер запущен: `python src/main.py`

### "В Google Sheets нет новых логов"

- Проверьте что лист "Логи" существует
- Убедитесь что синхро срабатывает
- Смотрите документацию: `docs/LOGS_RELEASE_NOTES.md`

---

## 📖 Дополнительно

- **LOGGING_README.md** - главное описание
- **docs/LOGS_QUICK_REFERENCE.md** - справочник для разработчиков
- **docs/TESTING_CHECKLIST.md** - полный чек-лист тестирования
- **FINAL_REPORT.md** - финальный отчёт

---

**Готово!** ✅
