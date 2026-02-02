# 📋 Руководство по Логированию - AgentCare System
**Версия**: 1.0
**Статус**: ✅ Активно
**Дата создания**: 14 января 2026 г.
**Последнее обновление**: 14 января 2026 г.

---

## 📑 Оглавление

1. [Обзор системы логирования](#обзор-системы-логирования)
2. [Структура логов](#структура-логов)
3. [Категории логирования](#категории-логирования)
4. [Статусы операций](#статусы-операций)
5. [Уровни логирования](#уровни-логирования)
6. [Правила составления логов](#правила-составления-логов)
7. [Примеры использования](#примеры-использования)
8. [Хранение и архивирование](#хранение-и-архивирование)
9. [Лучшие практики](#лучшие-практики)
10. [Часто задаваемые вопросы](#часто-задаваемые-вопросы)

---

## Обзор системы логирования

### Архитектура

Система логирования состоит из **двух основных компонентов**:

| Компонент | Назначение | Лист | Хранение | Архивирование |
|-----------|-----------|------|----------|---------------|
| **LOG** | Журнал синхронизации | "Журнал синхро" | 24 часа | Каждые 12 часов |
| **LOG_DEBUG** | Технический лог (всё) | "Журнал логов" | 24 часа | Каждые 12 часов |

### Принцип разделения

```
LOG              → Синхронизация данных (logSynchronization функция)
LOG_DEBUG        → Все события (logWithEmoji функция + API вызовы)
```

**Критично**: LOG_DEBUG содержит ВСЕ события, включая DEBUG, INFO, WARN, ERROR.

---

## Структура логов

### LOG_DEBUG (Журнал логов) - 8 колонок

```
A: Время                    | Дата/время события | "2026-01-14 12:34:56"
B: Категория                | 15 категорий       | "Sync", "Article", "Price"
C: Статус                   | 5 статусов         | "START", "PROGRESS", "SUCCESS", "WARNING", "ERROR"
D: Сообщение                | Текст события      | "🔵 Начало синхронизации"
E: Функция                  | Имя функции        | "processPriceSheet_MT"
F: Детали                   | Доп. информация    | "Строка 42, ошибка валидации"
G: Параметры (JSON)         | Входные данные     | {"row": 42, "article": "ABC123"}
H: Результаты (JSON)        | Результаты         | {"success": true, "updated": 5}
```

### LOG (Журнал синхро) - 9 колонок

```
A: Дата/время               | Событие синхро     | "2026-01-14 12:34:56"
B: ID                       | Идентификатор      | "SYNC-123456"
C: Источник                 | Откуда данные      | "Динамика цены"
D: Цель                     | Куда данные        | "Основная таблица"
E: Старое значение          | Было               | "1500"
F: Новое значение           | Стало              | "1750"
G: Категория                | Тип данных         | "Price", "Article", "Stock"
H: Хэштеги                  | Теги для фильтра   | "#urgent #warning"
I: Событие                  | Описание            | "Цена обновлена"
```

---

## Категории логирования

### 15 Основных Категорий

| # | Категория | Эмодзи | Цвет | Примеры операций |
|---|-----------|--------|------|-----------------|
| 1 | **Article** | 📦 | Оранжевый | addArticle, deleteArticle, updateArticle |
| 2 | **Sync** | 🔄 | Зелёный | syncRow, runFullSync, handleOnChange |
| 3 | **Delete** | 🗑️ | Красный | deleteSelectedRows, purgeData |
| 4 | **Network** | 🌐 | Синий | callServer, fetchData, uploadFile |
| 5 | **Price** | 💰 | Жёлтый | processPriceSheet, updatePrice, calculateCost |
| 6 | **Certificate** | 📄 | Фиолетовый | createNewsSheet, generateProtocols |
| 7 | **Invoice** | 📋 | Голубой | createFullInvoice, collectDocuments |
| 8 | **Order** | 📦 | Коричневый | showOrderStage, formatOrderSheet |
| 9 | **Label** | 🏷️ | Розовый | generateLabel, printLabel |
| 10 | **Spirit** | 🥃 | Фиолетовый | calculateSpiritNumbers, generateProtocols |
| 11 | **Dashboard** | 📊 | Синий | loadLogsIntoDashboard, refreshDashboard |
| 12 | **Menu** | 📋 | Серый | createAgentMenu, buildRegistry |
| 13 | **Cache** | 💾 | Серебристый | saveToDrive, loadFromCache |
| 14 | **Lock** | 🔒 | Красный | acquireLock, releaseLock |
| 15 | **Settings** | ⚙️ | Серый | showSettings, setupGemini, checkService |

---

## Статусы операций

### 5 Основных Статусов

| Статус | Эмодзи | Цвет | Значение | Когда использовать |
|--------|--------|------|---------|-------------------|
| **START** | 🔵 | Синий | Операция началась | В начале функции, перед основной логикой |
| **PROGRESS** | 🟡 | Жёлтый | Выполнение в процессе | На критических точках, перед долгими операциями |
| **SUCCESS** | 🟢 | Зелёный | Успешное завершение | В конце функции, если без ошибок |
| **WARNING** | 🟠 | Оранжевый | Предупреждение | Некритичные проблемы, странное поведение |
| **ERROR** | 🔴 | Красный | Ошибка | В catch блоке, критичные сбои |

### Рекомендуемый паттерн

```javascript
function myOperation(input) {
  // 1. START - начало с параметрами
  Lib.logWithEmoji("Начало операции", "INFO", "", "MyCategory", JSON.stringify(input));

  try {
    // 2. PROGRESS - критичные точки
    Lib.logWithEmoji("Загрузка данных", "INFO", "", "MyCategory", "");
    const data = loadData();

    // 3. PROGRESS - перед долгой операцией
    Lib.logWithEmoji("Обработка данных", "INFO", "", "MyCategory", "");
    const result = processData(data);

    // 4. SUCCESS - успешное завершение
    Lib.logWithEmoji("Операция завершена успешно", "INFO", "", "MyCategory", JSON.stringify(result));
    return result;

  } catch (error) {
    // 5. ERROR - обработка ошибок
    Lib.logWithEmoji("Ошибка: " + error.message, "ERROR", "", "MyCategory", error.stack);
    throw error;
  }
}
```

---

## Уровни логирования

### 4 Уровня с приоритетами

| Уровень | Эмодзи | Приоритет | Описание | Примеры |
|---------|--------|-----------|---------|---------|
| **DEBUG** | 🐛 | 1 (низкий) | Детальная отладочная информация | Значения переменных, промежуточные расчёты |
| **INFO** | ℹ️ | 2 | Информационные сообщения | Начало операции, прогресс, успех |
| **WARN** | ⚠️ | 3 | Предупреждения | Некритичные ошибки, странные данные |
| **ERROR** | ❌ | 4 (высокий) | Ошибки | Критичные сбои, исключения |

### Текущий уровень логирования

Установлен в `01Config.js` → `SETTINGS.CURRENT_LOG_LEVEL`

```javascript
CURRENT_LOG_LEVEL: SETTINGS.LOG_LEVELS["INFO"],  // Только INFO и выше
```

---

## Правила составления логов

### Правило 1: Структурированная информация

✅ **Хорошо**:
```javascript
Lib.logWithEmoji(
  "Синхронизация строки завершена",
  "INFO",
  "",
  "Sync",
  JSON.stringify({row: 42, status: "success", time: "2.5s"})
);
```

❌ **Плохо**:
```javascript
Lib.logWithEmoji("sync done row42", "INFO");
```

### Правило 2: Категория должна быть из списка 15

✅ **Хорошо**: `"Article"`, `"Sync"`, `"Price"`
❌ **Плохо**: `"MyCustom"`, `"foo"`, `""` (используй "Dashboard" или другую близкую)

### Правило 3: Сообщение должно быть понятным

✅ **Хорошо**: `"Цена обновлена для артикула ABC123 с 1500 на 1750"`
❌ **Плохо**: `"error"`, `"ok"`, `"done"`

### Правило 4: JSON в колонках G и H

- **Параметры (G)**: Входные данные, параметры функции
  ```json
  {"article_id": "ABC123", "new_price": 1750, "user_id": "u123"}
  ```

- **Результаты (H)**: Результаты работы, что было сделано
  ```json
  {"rows_updated": 5, "success": true, "duration_ms": 250}
  ```

### Правило 5: Логирование в начале и конце

```javascript
function processData(data) {
  // В начале - START
  logWithEmoji("Начало обработки данных", "INFO", "", "Article",
    JSON.stringify({count: data.length}));

  try {
    // ... логика функции ...

    // В конце - SUCCESS
    logWithEmoji("Обработка завершена успешно", "INFO", "", "Article",
      JSON.stringify({processed: data.length, errors: 0}));

    return result;
  } catch (error) {
    // В случае ошибки - ERROR
    logWithEmoji("Ошибка обработки: " + error.message, "ERROR", "", "Article",
      error.stack);
    throw error;
  }
}
```

---

## Примеры использования

### Пример 1: Простая операция

```javascript
function addArticle(articleName) {
  const category = "Article";

  // START - начало операции
  logWithEmoji(`Добавление артикула "${articleName}"`, "INFO", "", category, "");

  try {
    // PROGRESS - промежуточный этап
    logWithEmoji("Проверка наличия артикула", "INFO", "", category, "");
    if (articleExists(articleName)) {
      throw new Error("Артикул уже существует");
    }

    // SUCCESS - успешно добавлен
    const id = createArticle(articleName);
    logWithEmoji(`Артикул добавлен (ID: ${id})`, "INFO", "", category,
      JSON.stringify({article_id: id, name: articleName}));

    return id;

  } catch (error) {
    // ERROR - ошибка при добавлении
    logWithEmoji(`Ошибка добавления: ${error.message}`, "ERROR", "", category, error.stack);
    throw error;
  }
}
```

### Пример 2: Синхронизация с параметрами

```javascript
function syncSelectedRow(rowIndex, options) {
  const category = "Sync";

  // START с параметрами
  logWithEmoji("Синхронизация строки", "INFO", "", category,
    JSON.stringify({row: rowIndex, ...options}));

  try {
    // PROGRESS - загрузка
    logWithEmoji("Загрузка данных из строки", "INFO", "", category, "");
    const data = loadRowData(rowIndex);

    // PROGRESS - отправка на сервер
    logWithEmoji("Отправка данных на сервер", "INFO", "", category, "");
    const result = callServer("/sync/row", {row_index: rowIndex, data: data});

    // SUCCESS с результатами
    logWithEmoji("Синхронизация успешна", "INFO", "", category,
      JSON.stringify({
        row: rowIndex,
        status: "success",
        updated_fields: result.updated_fields,
        time_ms: result.time_ms
      }));

    return result;

  } catch (error) {
    // ERROR с полной информацией
    logWithEmoji(`Ошибка синхронизации: ${error.message}`, "ERROR", "", category,
      JSON.stringify({
        row: rowIndex,
        error: error.message,
        stack: error.stack
      }));
    throw error;
  }
}
```

### Пример 3: Сложная операция с несколькими этапами

```javascript
function processPriceSheet(sheetName) {
  const category = "Price";

  logWithEmoji(`Обработка листа "${sheetName}"`, "INFO", "", category, "");

  try {
    // Этап 1: Загрузка
    logWithEmoji("Загрузка данных о ценах", "INFO", "", category, "");
    const prices = loadPrices(sheetName);

    // Этап 2: Валидация
    logWithEmoji("Валидация цен", "INFO", "", category, "");
    const validated = validatePrices(prices);

    if (validated.errors.length > 0) {
      logWithEmoji(`Найдены ошибки валидации: ${validated.errors.length}`,
        "WARN", "", category, JSON.stringify(validated.errors));
    }

    // Этап 3: Обновление
    logWithEmoji("Обновление цен на сервере", "INFO", "", category, "");
    const updateResult = updatePricesOnServer(validated.valid);

    // SUCCESS
    logWithEmoji("Обработка завершена", "INFO", "", category,
      JSON.stringify({
        total: prices.length,
        valid: validated.valid.length,
        errors: validated.errors.length,
        updated: updateResult.count,
        time_ms: updateResult.time_ms
      }));

    return updateResult;

  } catch (error) {
    logWithEmoji(`Критичная ошибка: ${error.message}`, "ERROR", "", category, error.stack);
    throw error;
  }
}
```

---

## Хранение и архивирование

### Хранение логов

| Параметр | Значение | Описание |
|----------|----------|---------|
| **Активное хранилище** | 24 часа | Последние 24 часа в LOG_DEBUG |
| **Автоматическое архивирование** | Каждые 12 часов | Копирование в месячный архив |
| **Расположение архива** | Google Drive папка | `/Логи/{проект}/{месяц}` |
| **Сохранение архивов** | Неограниченно | Архивы хранятся вечно |

### Процесс архивирования

```
1. Каждые 12 часов запускается midnightLogRotation()
   ↓
2. Все логи из LOG_DEBUG копируются в месячный архив
   Имя: MT-January-2026.xlsx (пример)
   ↓
3. Лист LOG_DEBUG очищается (остаются только заголовки)
   ↓
4. Архив сохраняется в Google Drive
```

### Очистка логов

**Автоматическая очистка**:
```javascript
quickCleanLogSheet()  // Оставляет 100 последних записей
```

**Вручную**:
```javascript
// Из меню: ⚙️ ЭКОСИСТЕМА → 📊 Логи → 🧹 Очистить журнал (быстро)
```

---

## Лучшие практики

### 1. ✅ Всегда логируйте критичные операции

```javascript
// ДА - логируем начало и конец
function updateInventory(articles) {
  logWithEmoji("Обновление инвентаря", "INFO", "", "Article", "");
  try {
    // ... логика ...
    logWithEmoji("Инвентарь обновлён", "INFO", "", "Article", "");
  } catch (e) {
    logWithEmoji("Ошибка: " + e.message, "ERROR", "", "Article", "");
  }
}

// НЕТ - нет логирования
function updateInventory(articles) {
  // ... логика без логов ...
}
```

### 2. ✅ Используйте правильные категории

```javascript
// ДА - категория соответствует операции
logWithEmoji("Цена обновлена", "INFO", "", "Price", "");

// НЕТ - неправильная категория
logWithEmoji("Цена обновлена", "INFO", "", "General", "");
```

### 3. ✅ Логируйте параметры и результаты

```javascript
// ДА - структурированная информация
logWithEmoji("Синхронизация", "INFO", "", "Sync",
  JSON.stringify({row: 42, status: "success", time_ms: 250}));

// НЕТ - просто текст без деталей
logWithEmoji("Синхронизирована строка 42", "INFO", "", "Sync", "");
```

### 4. ✅ Логируйте ошибки с полной информацией

```javascript
// ДА - полная информация об ошибке
catch (error) {
  logWithEmoji("Ошибка: " + error.message, "ERROR", "", "Sync",
    JSON.stringify({error: error.message, stack: error.stack}));
}

// НЕТ - минимум информации
catch (error) {
  logWithEmoji("Ошибка!", "ERROR", "", "Sync", "");
}
```

### 5. ✅ Используйте JSON для структурированных данных

```javascript
// ДА - JSON структура
const details = {
  article_id: "ABC123",
  old_price: 1500,
  new_price: 1750,
  reason: "market_update"
};
logWithEmoji("Цена изменена", "INFO", "", "Price", JSON.stringify(details));

// НЕТ - просто текст
logWithEmoji("ABC123 1500->1750 market", "INFO", "", "Price", "");
```

### 6. ✅ Избегайте дублирования логов

```javascript
// ДА - один лог с полной информацией
logWithEmoji("Обновлено 5 артикулов", "INFO", "", "Article",
  JSON.stringify({updated: 5, failed: 0, total: 5}));

// НЕТ - пять отдельных логов
for (let i = 0; i < 5; i++) {
  logWithEmoji("Артикул обновлён", "INFO", "", "Article", "");
}
```

---

## Часто задаваемые вопросы

### Q1: Как логировать длинные строки?

**A**: Используйте JSON для структурирования:
```javascript
const longData = {
  description: "Очень длинное описание...",
  details: {...}
};
logWithEmoji("Данные обновлены", "INFO", "", "Article", JSON.stringify(longData));
```

### Q2: Как найти лог определённого события?

**A**: Используйте фильтры в Google Sheets или Дашборде:
1. Откройте 📊 Дашборд логов из меню
2. Используйте фильтры по Категории, Статусу, Времени
3. Или используйте поиск по тексту сообщения

### Q3: Как экспортировать логи?

**A**:
```javascript
// Из меню
⚙️ ЭКОСИСТЕМА → 📊 Логи → 📦 Архивировать логи

// Или вручную
Lib.manualArchiveLogs()
```

### Q4: Можно ли изменить категорию лога?

**A**: Нет, категория выбирается при создании лога и не может быть изменена.
Если нужно добавить новую категорию - обновите список в LOGGING_GUIDE.md и
`Lib.__LOG_CATEGORIES__` в z-server-overrides.js

### Q5: Почему мой лог не появляется в Дашборде?

**A**: Проверьте:
1. Уровень логирования в `SETTINGS.CURRENT_LOG_LEVEL` (должен быть INFO или ниже)
2. Существует ли лист "Журнал логов" (LOG_DEBUG)
3. Правильно ли составлена категория (из списка 15)
4. Нет ли ошибок в JSON параметрах

### Q6: Как установить уровень логирования?

**A**: В 01Config.js найдите `CURRENT_LOG_LEVEL`:
```javascript
// Только INFO и выше (скрывает DEBUG)
CURRENT_LOG_LEVEL: SETTINGS.LOG_LEVELS["INFO"],

// Все логи включая DEBUG
CURRENT_LOG_LEVEL: SETTINGS.LOG_LEVELS["DEBUG"],

// Только WARN и ERROR
CURRENT_LOG_LEVEL: SETTINGS.LOG_LEVELS["WARN"],
```

### Q7: Сколько логов я могу записать?

**A**:
- В час: ~5000 логов
- За 24 часа: ~120 000 логов (практически неограниченно)
- Архив: неограниченно (в Google Drive)

### Q8: Как удалить все логи?

**A**:
```javascript
// Из меню
⚙️ ЭКОСИСТЕМА → 📊 Логи → 🧹 Очистить журнал (быстро)

// Или вручную
Lib.quickCleanLogSheet()  // Оставляет 100 последних записей
```

---

## Интеграция с планом разработки

### Phase 1: Инфраструктура ✅

- ✅ Создана структура LOG_DEBUG с 8 колонками
- ✅ Определены 15 категорий логирования
- ✅ Определены 5 статусов операций
- ✅ Реализованы функции для логирования

### Phase 2: Реструктурирование меню

- 🔲 Добавить опции для просмотра и архивирования логов в меню
- 🔲 Создать интерфейс управления логированием

### Phase 3: Логирование функций

- 🔲 Добавить логирование ко всем 50+ функциям
- 🔲 Использовать категории в соответствии с типом операции
- 🔲 Логировать параметры и результаты для каждой функции

### Phase 4: Тестирование

- 🔲 Проверить что все логи записываются корректно
- 🔲 Проверить фильтры в Дашборде
- 🔲 Проверить архивирование логов

---

## Чеклист для разработчика

Перед началом работы над функцией:

- [ ] Выбрал правильную категорию из 15 основных
- [ ] Добавил лог START в начале функции
- [ ] Добавил логи PROGRESS в критичных точках
- [ ] Добавил лог SUCCESS в конце успешного выполнения
- [ ] Добавил лог ERROR в catch блоке
- [ ] Структурировал параметры в JSON
- [ ] Структурировал результаты в JSON
- [ ] Проверил что сообщение понятно и информативно
- [ ] Протестировал что логи появляются в LOG_DEBUG

---

**Конец документа**

Версия: 1.0
Статус: ✅ АКТИВНО
Последнее обновление: 14 января 2026 г.

Для обновлений обращайтесь к TASKS.md и PLAN_TESTING.md
