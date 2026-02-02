# Правила именования файлов роутеров

> [!IMPORTANT]
> Эта конвенция является обязательной для всех файлов в `src/api/routes/`

## Структура именования

### Файлы для функций меню: `<номер>m-<название>.py`

**Формат:** `{порядковый_номер}m-{название_латиницей}.py`

- **Порядковый номер** — соответствует порядку группы в меню (см. `src/config/project_menus.py`)
- **Префикс `m`** — означает "menu" (функция из меню)
- **Название** — латиницей, snake_case, отражает название группы меню

**Примеры:**
```
01m-zakaz.py              # 🧾 Заказ
02m-stadii_zakaza.py      # 📊 Стадии по заказ
03m-price_list.py         # 🏷️ Прайс-лист
04m-documents.py          # 🛒 Заказ & Документация
05m-export.py             # 📦 Экспорт & Выгрузка
06m-ai_agent.py           # 🤖 AI Агент
07m-settings.py           # ⚙️ Настройка
08m-dev_tools.py          # 🛠️ Инструменты разработчика
```

### Файлы для серверных функций: `s_<название>.py`

**Формат:** `s_{название_латиницей}.py`

- **Префикс `s_`** — означает "server" (внутренняя серверная функция, не связанная с меню)
- **Название** — латиницей, snake_case, отражает назначение

**Примеры:**
```
s_sync.py                 # Синхронизация (не в меню, вызывается триггерами)
s_rules.py                # Управление правилами синхронизации
s_logs.py                 # Журналы синхронизации
s_meta.py                 # Метаданные и discovery
```

---

## Содержимое файла роутера

Каждый файл роутера должен содержать:

1. **Описание модуля** (docstring в начале файла)
2. **Импорты**
3. **Создание роутера**: `router = APIRouter()`
4. **Pydantic модели** (если специфичны только для этого роутера)
5. **Эндпоинты** с декораторами `@router.get/post/patch/delete`
6. **Подробная документация** к каждой функции:
   - Описание что делает функция
   - Какие параметры принимает
   - Что возвращает
   - Примеры использования (опционально)

### Пример структуры файла:

```python
"""
Роутер для функций группы меню "🧾 Заказ".

Содержит:
- Обработка данных основного листа
- Загрузка остатков со склада
"""

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel, Field

from src.services.price_processor import get_price_processor
from src.services.stock_processor import StockProcessor
from src.utils.logger import logger

router = APIRouter()

# ============== Models ==============

class ProcessDataRequest(BaseModel):
    """Запрос на обработку данных."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    mode: str = Field("main", description="Режим обработки")


# ============== Endpoints ==============

@router.post("/process-data", summary="Обработка данных")
async def process_primary_data(request: ProcessDataRequest):
    """
    Обрабатывает данные основного листа "Заказ".
    
    Функция выполняет:
    1. Валидацию данных
    2. Обновление цен из прайс-листа
    3. Пересчёт формул
    
    Args:
        request: Параметры обработки
        
    Returns:
        dict: Статус обработки и количество обработанных строк
        
    Raises:
        HTTPException: При ошибке обработки данных
        
    Соответствует кнопке меню:
        🧾 Заказ → 📥 Обработка (serverProcessPrimaryData)
    """
    try:
        processor = get_price_processor()
        result = await processor.process_data(
            spreadsheet_id=request.spreadsheet_id,
            mode=request.mode
        )
        return {"status": "success", "result": result}
    except Exception as e:
        logger.error("process_data_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))
```

---

## Маппинг меню → роутеры

| Группа меню | Файл роутера | Функции |
|------------|--------------|---------|
| 🧾 Заказ | `01m-zakaz.py` | Обработка, Загрузка остатков |
| 📊 Стадии по заказ | `02m-stadii_zakaza.py` | Сортировка, Фильтры стадий |
| 🏷️ Прайс-лист | `03m-price_list.py` | Анализ цен, Формирование прайса |
| 🛒 Заказ & Документация | `04m-documents.py` | Заказы, Сертификация, Документы |
| 📦 Экспорт & Выгрузка | `05m-export.py` | Выгрузка акций/наборов, Инвойсы |
| 🤖 AI Агент | `06m-ai_agent.py` | AI анализ, Smart Match |
| ⚙️ Настройка | `07m-settings.py` | Настройка правил, Триггеры |
| 🛠️ Инструменты разработчика | `08m-dev_tools.py` | Туннели, Обновление меню |

### Серверные функции (не в меню)

| Назначение | Файл роутера | Эндпоинты |
|-----------|--------------|-----------|
| Синхронизация | `s_sync.py` | `/sync/event`, `/sync/batch-event`, `/sync/row` |
| Правила синхронизации | `s_rules.py` | `/rules/{id}`, CRUD операции |
| Журналы | `s_logs.py` | `/sync-logs/`, `/function-logs/` |
| Метаданные | `s_meta.py` | `/meta/`, discovery endpoints |

---

## Подключение роутеров

Все роутеры подключаются в `src/api/router.py`:

```python
from fastapi import APIRouter
from src.api.routes import (
    # Menu routes
    m01_zakaz,
    m02_stadii_zakaza,
    m03_price_list,
    m04_documents,
    m05_export,
    m06_ai_agent,
    m07_settings,
    m08_dev_tools,
    # Server routes
    s_sync,
    s_rules,
    s_logs,
    s_meta,
)

router = APIRouter()

# Menu routes
router.include_router(m01_zakaz.router, prefix="/order", tags=["🧾 Заказ"])
router.include_router(m02_stadii_zakaza.router, prefix="/stages", tags=["📊 Стадии"])
# ... и т.д.

# Server routes
router.include_router(s_sync.router, prefix="/sync", tags=["Sync"])
router.include_router(s_rules.router, prefix="/rules", tags=["Rules"])
# ... и т.д.
```

---

## Преимущества такого подхода

1. **Соответствие UI и кода** — структура файлов точно отражает структуру меню
2. **Простая навигация** — номера в именах файлов соответствуют порядку в меню
3. **Быстрый поиск** — по названию кнопки в меню сразу понятно, где искать код
4. **Модульность** — все связанные функции в одном файле
5. **Документированность** — каждая функция имеет подробное описание

---

## Обязательные требования

> [!CAUTION]
> 1. **НЕ изменять** пути эндпоинтов `/api/v1/...` — обратная совместимость с GAS критична
> 2. **НЕ изменять** форматы JSON ответов
> 3. **Каждая функция** должна иметь docstring с описанием и соответствием кнопке меню
> 4. **Импорты сервисов** — использовать глобальные singleton экземпляры из `endpoints.py`
---

## 🚀 План реализации (Roadmap)

### Этап 1: Подготовка (1-2 часа)
1. Создать директорию `src/api/models/`
2. Создать пустые файлы роутеров в `src/api/routes/` согласно маппингу выше.

### Этап 2: Миграция моделей (1 час)
1. Перенести Pydantic модели из `endpoints.py` в `src/api/models/`
2. Настроить импорты в каждом файле моделей.
3. Обновить `src/api/models/__init__.py`.

### Этап 3: Миграция эндпоинтов (3-4 часа)
Перенос кода из `endpoints.py` в соответствующие файлы роутеров в следующем порядке:
1. `s_meta.py`
2. `07m-settings.py`
3. `s_logs.py`
4. `s_rules.py`
5. `s_sync.py` (критично!)
6. `01m-zakaz.py`
7. `02m-stadii_zakaza.py`
8. `06m-ai_agent.py`
9. `05m-export.py`
10. `04m-documents.py`

### Этап 4: Интеграция и верификация (2 часа)
1. Обновить `src/api/router.py` для подключения всех новых роутеров.
2. Удалить/архивировать старый `endpoints.py`.
3. Запустить тесты `pytest tests/`.
4. Проверить работоспособность кнопок в Google Sheets.

---

## 📝 Критерии успеха
- [ ] Сервер успешно запускается.
- [ ] Все пути API (`/api/v1/...`) остались неизменными.
- [ ] GAS функции работают без ошибок.
- [ ] Каждая функция задокументирована с указанием соответствия кнопке меню.
