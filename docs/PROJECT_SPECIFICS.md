# AgentCare - Специфика проектов MT/SK/SS

## Обзор различий

| Аспект | MT (CosmeticaBar) | SK (Carmado) | SS (ООО Сан) |
|--------|-------------------|--------------|--------------|
| Режимы обработки | 3 (Б/З, Тестер, Пробник) | 2 (Основной, Пробники) | 1 (Основной) |
| Множитель цены | Пробники ×10 | Нет | Нет |
| Скидка | Нет | 20% на Динамика цены | Нет |
| ID-G (группы) | Да | Да | Нет |
| ID-L (линии) | Да | Да (очищается отдельно) | Нет |
| RRP | Нет | Да | Нет |
| Парсинг по цвету | Нет | Да (groupColor) | Нет |
| Комбинирование групп | Нет | "Линия - Группа" | Нет |
| Маркер "-ПРОФ" | Нет | Нет | Да |

---

## MT (CosmeticaBar) — Детали

### Три режима обработки

#### 1. Основной прайс (`processMtMainPrice`)

**Функция:** Обработка Б/З поставщик

**Входные колонки Excel:**
| Индекс | Название | Маппинг |
|--------|----------|---------|
| 0 | CODE | Арт. произв. |
| 1 | CATEGORY | (определяет группу) |
| 2 | DESCRIPTION | Название ENG |
| 3 | FORMAT | Объём |
| 4 | UNITS | шт./уп. |
| 5 | PRICE | Цена EXW |
| 6 | EAN | BAR CODE |

**Логика определения группы:**
```python
# Псевдокод
if code_value and not category_value:
    # Это строка группы
    current_group = description_value
elif code_value and description_value:
    # Это артикул
    article.group = current_group
```

**Выходные колонки:**
- ID-P, Арт. произв., Название ENG, Объём, BAR CODE, шт./уп., Цена, Группа

#### 2. Тестеры (`processMtTesterPrice`)

**Особенности:**
- Префикс к объёму: `"-Тестер " + volume`
- Суффикс к группе: `group + " - Тестер"`
- **НЕ очищает ID-P** — продолжает нумерацию от максимума
- Копирует цены без множителей

**Пример трансформации:**
```
Вход:  Объём = "50ml", Группа = "Face Care"
Выход: Объём = "-Тестер 50ml", Группа = "Face Care - Тестер"
```

#### 3. Пробники (`processMtSamplesPrice`)

**Особенности:**
- Суффикс к группе: `group + " - пробник"`
- **Множитель цены = 10** (передаётся в `copyPriceFromPrimaryToSheets`)
- После обработки вызывает `updateStatusesAfterProcessing()`

**Вызов копирования цены:**
```javascript
copyPriceFromPrimaryToSheets(processed, null, 10)  // multiplier = 10
```

### Порядок синхронизации MT

```
1. clearIdpColumnOnSheets()        — очистка ID-P
2. _syncIdWithMain()               — синхро ID между листами
3. fillIdpOnSheetsByIdFromPrimary() — заполнение ID-P
4. copyPriceFromPrimaryToSheets()  — копирование цены
5. recalculatePriceDynamicsFormulas() — формулы Динамика цены
6. (для пробников) updateStatusesAfterProcessing()
```

---

## SK (Carmado) — Детали

### Два режима обработки

#### 1. Основной прайс (`processSkPriceSheet`)

**Входные колонки Excel:**
| Название | Маппинг |
|----------|---------|
| CODE | Арт. произв. |
| PRODUCT | Название ENG |
| UNITS | шт./уп. |
| PRICE | Цена EXW |
| TOTAL | всего |
| RRP | RRP |

**Особенность: Парсинг по цвету фона**
```javascript
// Определение группы по цвету ячейки
function _detectGroupRow(row, bgRow, groupColor) {
    // groupColor берётся из config.GROUP_HEADER_COLOR
    // Если цвет фона совпадает — это строка группы
}
```

**Комбинирование "Линия - Группа":**
```javascript
combinedGroup = lineValue + " - " + currentGroup
// Пример: "Professional - Face Care"
```

**Специальная функция скидки:**
```javascript
_fillDiscountForPriceDynamics()
// Заполняет колонку "СКИДКА ОТ EXW [YEAR], %" = 20%
```

**Очистка ID-L (только SK):**
```javascript
clearIdlColumnOnSheets()  // Вызывается отдельно от ID-P
```

#### 2. Пробники (`processSkPriceProbes`)

**Входные колонки Excel (точный блок заголовков):**
```
REFERENC./PRODUCT/PRODUCTO/TYPE/AREA/UNITS/PRICE/TOTAL
```

**Маркер остановки:**
```
"CATÁLOGOS Y MATERIALES IMPRESOS / CATALOGUES & PRINTED MATERIALS"
```

**Логика определения артикула:**
```javascript
// Если артикул НЕ начинается с "00" — это группа
if (!/^0{2}/.test(referenceValue)) {
    // Это группа
}
```

**Последовательная нумерация ID-P:**
```javascript
_assignSequentialIdpForProbes_()
// Отдельная логика вместо стандартной синхронизации
// Опция: overrideIdpFromSource: true
```

### Порядок синхронизации SK

```
1. clearIdpColumnOnSheets()         — очистка ID-P
2. clearIdlColumnOnSheets()         — очистка ID-L (SK only!)
3. clearPriceExwColumnOnSheets()    — очистка цены
4. _syncIdWithMain()                — синхро ID
5. fillIdpOnSheetsByIdFromPrimary() — заполнение ID-P
6. copyPriceFromPrimaryToSheets()   — копирование цены
7. _fillDiscountForPriceDynamics()  — заполнение скидки 20% (SK only!)
8. recalculatePriceDynamicsFormulas() — формулы
9. updatePriceCalculationFormulas() — INDEX/MATCH
10. applyCalculationFormulas()      — расчётные формулы
```

---

## SS (ООО Сан) — Детали

### Один режим обработки

#### Основной прайс (`processSsPriceSheet`)

**Входные колонки Excel:**
| Индекс | Название | Маппинг |
|--------|----------|---------|
| 1 | CODE | Арт. произв. |
| 2 | PRODUCT NAME | Название ENG |
| 3 | SIZE | Объём |
| 4 | PACK | Форма выпуска |
| 5 | BAR CODE/ACL | BAR CODE |
| 6 | QTY/BOX | шт./уп. |
| 7 | EX WORKS CARROS (06) | Цена EXW |

**Выходные колонки:**
- ID-P, Арт. произв., Название ENG, Объём, Форма выпуска, BAR CODE, шт./уп., Цена, Группа

**Маркер режима PROFESSIONAL:**
```javascript
if (row_contains("PROFESSIONAL PRODUCTS")) {
    // Добавляет "-ПРОФ" к группе
    group = group + "-ПРОФ"
}
```

**Маркер остановки:**
```javascript
if (row_contains("PROMOTIONAL MATERIALS")) {
    // Прекращает обработку
    break;
}
```

**Логика SAMPLES:**
```javascript
// Для группы SAMPLES: извлекает число из "Форма выпуска" в "шт./уп."
if (group === "SAMPLES") {
    const match = formValue.match(/\d+/);
    if (match) {
        unitsPerPack = parseInt(match[0]);
    }
}
```

### Порядок синхронизации SS

```
1. clearIdpColumnOnSheets()         — очистка ID-P
2. clearPriceExwColumnOnSheets()    — очистка цены
3. _syncIdWithMain()                — синхро ID
4. fillIdpOnSheetsByIdFromPrimary() — заполнение ID-P
5. copyPriceFromPrimaryToSheets()   — копирование цены
6. recalculatePriceDynamicsFormulas() — формулы
7. updatePriceCalculationFormulas() — INDEX/MATCH
8. applyCalculationFormulas()       — расчётные формулы
9. updateStatusesAfterProcessing()  — обновление статусов
```

---

## Каскадные правила (CASCADE_RULES)

### Триггер: Лист "Сертификация"

**Триггерные колонки (SOURCE_HEADERS):**
- "Наименования рус по ДС"
- "Наименования англ по ДС"
- "Объём"
- "Код ТН ВЭД"

**Целевые поля (TARGET_FIELDS):**

| Код | Колонка | Формула |
|-----|---------|---------|
| RUS | Наименования рус по ДС | (источник) |
| ENG | Наименования англ по ДС | (источник) |
| VOL | Объём | (источник) |
| TNVED | Код ТН ВЭД | (источник) |
| VOL_EN | Объём англ. | Замены: мл→ml, гр→g |
| DS_NAME | Наименование ДС | `{RUS} / {ENG}` или `{RUS}, {ENG}` |
| INV_RU | Наименование для инвойса | `{DS_NAME} {VOL}\nКод ТН ВЭД: {TNVED}` |
| INV_EN | Наименование для инвойса Англ | `{ENG} {VOL_EN}\nCode: {TNVED}` |

**Логика формирования DS_NAME:**
```python
if rus_name.endswith(","):
    ds_name = f"{rus_name} {eng_name}"      # Без разделителя
else:
    ds_name = f"{rus_name} / {eng_name}"    # С разделителем "/"
```

**Замены для VOL_EN (VOLUME_EN_REPLACEMENTS):**
```yaml
replacements:
  - from: "мл"
    to: "ml"
  - from: "гр"
    to: "g"
  - from: "шт"
    to: "pcs"
```

**Условие применения:**
```python
# Применяется только если:
# 1. Изменённый заголовок в SOURCE_HEADERS
# 2. Все обязательные колонки присутствуют (RUS, ENG, VOL)
# 3. Новое значение отличается от текущего (areValuesEqual_)
```

---

## Формулы INDEX/MATCH

### Лист "Расчет цены"

**Применяется для SK и SS:**
```javascript
updatePriceCalculationFormulas(silent=true)
```

**Формула (примерная):**
```
=INDEX('Динамика цены'!$E:$E, MATCH($B2, 'Динамика цены'!$B:$B, 0))
```

**Подтягиваемые данные:**
- EXW предыдущая
- EXW текущая
- EXW ALFASPA текущая
- Закупочная цена
- DDP-МОСКВА

### Лист "Динамика цены"

**Расчётные формулы:**
```
EXW ALFASPA = ЦЕНА EXW из Б/З * (1 - СКИДКА/100)
Закупочная цена = EXW ALFASPA * курс_евро * коэффициент
DDP-МОСКВА = Закупочная цена + логистика + таможня
Прирост EXW = EXW текущая - EXW предыдущая
Прирост DDP = DDP текущая - DDP предыдущая
```

---

## Триггеры и события

### onEdit триггер

**Защита от гонки:**
```javascript
if (Lib._isSyncRunning) {
    Lib.logWarn("onEdit: занято другим запуском, пропускаем.");
    return;
}
```

**Специфичные триггеры для листа "Заказ":**

| Колонка | Действие |
|---------|----------|
| АКЦИИ | Заполняет "условие#", "коментарий#", "Срок#" |
| Набор | Заполняет "условие", "коментарий", "Срок" |
| СГ 1/2/3 | Автозаполнение "Срок#" и "Срок" |

### onChange триггер

**Функция:** `handleOnChange`

**Действия:**
1. Проверка проекта (MT/SK/SS)
2. Очистка кэша (если нужна)
3. Вызов `updateFieldCascade()` — каскадное обновление
4. Вызов `handleAvtoZakazUpdate()` — пересчёт заказов
5. Логирование в "Журнал синхро"

---

## ID документов по проектам

```yaml
# Google Sheets ID → Проект
document_project_map:
  "12yIL1CuESZxeUUd-oKK2brtN1FnXE9q95N7SqzNc7vk": "SS"
  "1CpYYLvRYslsyCkuLzL9EbbjsvbNpWCEZcmhKqMoX5zw": "SK"
  "1zSu0PzKKa5wvwMZCicwLN8N7Rwhs8XlJVrTrt2LMzQs": "SK"
  "1fMOjUE7oZV96fCY5j5rPxnhWGJkDqg-GfwPZ8jUVgPw": "MT"
  "1BW8Gk5_X2EZVjbnaa2yDm-bPzzlggwQrHepeNCcPCc0": "MT"
```

---

## Проверки проекта в коде

**В каждом файле обработки:**
```javascript
var TARGET_PROJECT_KEY = "MT";  // или "SK", "SS"

function _isActiveProject_() {
    return (
        global.CONFIG &&
        global.CONFIG.ACTIVE_PROJECT_KEY === TARGET_PROJECT_KEY
    );
}

// Перед обработкой:
if (!_isActiveProject_()) {
    ui.alert("Эта функция доступна только в проекте " + TARGET_PROJECT_KEY);
    return;
}
```

**Маппинг функций в Config.js:**
```yaml
PROCESSING:
  MAIN:
    SK: processSkPriceSheet
    SS: processSsPriceSheet
    MT: processMtMainPrice
  TESTER:
    MT: processMtTesterPrice
  SAMPLES:
    MT: processMtSamplesPrice
  PROBES:
    SK: processSkPriceProbes
  STOCKS:
    SK: loadSkStockData
    SS: loadSsStockData
    MT: loadMtStockData
  SORT_MANUFACTURER:
    SK: sortSkOrderByManufacturer
    SS: sortSsOrderByManufacturer
    MT: sortMtOrderByManufacturer
```

---

## Важные детали для реализации

### MT
1. **Три парсера** — разные режимы, разные трансформации
2. **Множитель ×10** для пробников — не забыть в `copyPriceFromPrimaryToSheets`
3. **Продолжение нумерации** для тестеров — не очищать ID-P

### SK
1. **Парсинг по цвету** — нужен доступ к форматированию ячеек
2. **Комбинирование "Линия - Группа"** — специфичная трансформация
3. **Скидка 20%** — отдельный шаг после копирования цен
4. **Очистка ID-L** — дополнительный шаг

### SS
1. **Маркеры режимов** — "PROFESSIONAL PRODUCTS" и "PROMOTIONAL MATERIALS"
2. **Суффикс "-ПРОФ"** — для профессиональных продуктов
3. **Извлечение числа** для SAMPLES из "Форма выпуска"
