# AgentCare - План миграции с Google Apps Script

## Обзор

Этот документ описывает пошаговый план переноса системы MyGoogleScripts (Google Apps Script) на AgentCare (Python).

---

## Исходная система (MyGoogleScripts)

### Статистика

| Метрика | Значение |
|---------|----------|
| Строк кода | 25,000+ |
| Модулей | 26 файлов |
| Главный модуль | 55,000 строк (03Синхронизация.js) |
| Листов в синхронизации | 20+ |
| Проектов | 3 (MT, SK, SS) |
| Триггеров | 4 типа |

### Модули для миграции

| Модуль GAS | Файл | Строк | Python модуль | Приоритет |
|------------|------|-------|---------------|-----------|
| Конфиг | 01Config.js | 1,400 | config/*.yaml | P1 |
| Парсинг MT | 02Обработка_Mt.js | 2,000 | parsers/mt_parser.py | P1 |
| Парсинг SK | 02Обработка_Sk.js | 1,800 | parsers/sk_parser.py | P1 |
| Парсинг SS | 02Обработка_Ss.js | 1,500 | parsers/ss_parser.py | P1 |
| Синхронизация | 03Синхронизация.js | 55,000 | core/sync_engine.py | P1 |
| Инвойсы | 04Инвойсы.js | 3,000 | services/invoice_service.py | P2 |
| Цены | 05Цены.js | 2,500 | services/price_processor.py | P1 |
| Статусы | 06Статусы.js | 1,200 | core/rules_engine.py | P2 |
| Автозаказ | 07AvtoZakaz.js | 2,000 | services/order_service.py | P3 |
| Сертификация | 08Сертификация.js | 4,000 | services/certification.py | P3 |
| Gemini API | Automation/01GeminiAPI.js | 550 | integrations/gemini_client.py | P1 |
| Sheets интеграция | Automation/02SheetsIntegration.js | 400 | integrations/google_sheets.py | P1 |

---

## Фазы миграции

### Фаза 0: Подготовка (1 неделя)

**Цель:** Настроить инфраструктуру и подготовить окружение.

#### Задачи

- [ ] Арендовать VPS (Hetzner CX31 рекомендуется)
- [ ] Настроить Docker и docker-compose
- [ ] Создать Google Cloud проект
- [ ] Получить Service Account credentials
- [ ] Получить Gemini API key
- [ ] Дать Service Account доступ к таблицам
- [ ] Настроить домен и SSL (Let's Encrypt)
- [ ] Развернуть базовый скелет AgentCare

#### Результат

- Сервер доступен по HTTPS
- `/health` endpoint работает
- Google API авторизация проверена

---

### Фаза 1: Интеграции (1-2 недели)

**Цель:** Реализовать базовые интеграции с Google APIs.

#### 1.1 Google Sheets Client

```python
# src/integrations/google_sheets.py
class GoogleSheetsClient:
    async def read_sheet(spreadsheet_id, sheet_name, range) -> DataFrame
    async def write_sheet(spreadsheet_id, sheet_name, data, range)
    async def batch_update(spreadsheet_id, updates: List[SheetUpdate])
    async def find_row_by_value(spreadsheet_id, sheet, column, value) -> int
    async def append_row(spreadsheet_id, sheet, data)
    async def delete_row(spreadsheet_id, sheet, row_number)
```

**Тестирование:**
- [ ] Чтение листа "Главная" из MT
- [ ] Запись в тестовый лист
- [ ] Batch update 100 ячеек
- [ ] Поиск строки по артикулу

#### 1.2 Google Drive Client

```python
# src/integrations/google_drive.py
class GoogleDriveClient:
    async def create_folder(name, parent_id) -> str
    async def copy_file(file_id, new_name, folder_id) -> str
    async def get_file_content(file_id) -> bytes
    async def pdf_to_text(file_id) -> str
    async def list_files(folder_id) -> List[DriveFile]
```

**Тестирование:**
- [ ] Создание папки
- [ ] Копирование документа
- [ ] Конвертация PDF в текст

#### 1.3 Gemini Client

```python
# src/integrations/gemini_client.py
class GeminiClient:
    async def analyze_inci(text, product_name) -> INCIAnalysis
    async def classify_product(name, description) -> ProductCategory
    async def extract_from_pdf(pdf_text) -> Dict
```

**Тестирование:**
- [ ] Анализ тестового INCI
- [ ] Retry при rate limit
- [ ] Обработка ошибок

---

### Фаза 2: Парсеры (1 неделя)

**Цель:** Реализовать парсеры Excel для каждого проекта.

#### 2.1 Базовый парсер

```python
# src/parsers/base_parser.py
class BaseParser(ABC):
    COLUMN_MAPPING: Dict[str, str]
    REQUIRED_FIELDS: List[str]

    def parse(file_path) -> DataFrame
    def validate(data) -> ValidationResult
    def transform(data) -> DataFrame
```

#### 2.2 Парсеры проектов

| Проект | Особенности | Файл |
|--------|-------------|------|
| MT | Тестеры, пробники, ID-G | parsers/mt_parser.py |
| SK | RRP, скидки, пробники | parsers/sk_parser.py |
| SS | Базовый формат | parsers/ss_parser.py |

**Тестирование:**
- [ ] Парсинг реального файла MT
- [ ] Парсинг реального файла SK
- [ ] Парсинг реального файла SS
- [ ] Валидация обязательных полей
- [ ] Обработка дубликатов

---

### Фаза 3: Ядро синхронизации (2-3 недели)

**Цель:** Реализовать движок синхронизации.

#### 3.1 Sync Engine

```python
# src/core/sync_engine.py
class SyncEngine:
    async def process_change(source_sheet, changed_rows, changed_columns)
    async def full_sync(source_sheet)
    async def sync_row(article)
    def get_dependent_sheets(source) -> List[str]
```

#### 3.2 Rules Engine

```python
# src/core/rules_engine.py
class RulesEngine:
    def load_rules(config_path)
    def apply(data, context) -> Tuple[DataFrame, List[RuleResult]]
    def validate(data) -> ValidationResult
```

#### 3.3 Transaction Manager

```python
# src/core/transaction.py
class Transaction:
    async def __aenter__()  # Создание снимков
    async def __aexit__()   # Commit или rollback
    async def update(sheet, changes)
    async def rollback()
```

**Тестирование:**
- [ ] Синхронизация одной строки
- [ ] Каскадное обновление 5 листов
- [ ] Откат при ошибке
- [ ] Обработка конфликтов

---

### Фаза 4: Сервисы (2 недели)

**Цель:** Реализовать бизнес-логику.

#### 4.1 Price Processor (9 фаз)

```python
# src/services/price_processor.py
class PriceProcessor:
    async def run(file_path, start_phase=1) -> ProcessingResult

    # Фазы
    async def _phase_load_excel()
    async def _phase_clear_columns()
    async def _phase_sync_with_base()
    async def _phase_save_snapshot()
    async def _phase_copy_prices()
    async def _phase_fill_ids()
    async def _phase_process_new()
    async def _phase_apply_formulas()
    async def _phase_update_statuses()
```

#### 4.2 Sync Service

```python
# src/services/sync_service.py
class SyncService:
    async def sync_row(project, article)
    async def sync_range(project, sheet, range)
    async def full_sync(project, sheet)
```

#### 4.3 Invoice Service

```python
# src/services/invoice_service.py
class InvoiceService:
    async def create_invoice(project, rows)
    async def generate_document(invoice)
    async def calculate_totals(items)
```

**Тестирование:**
- [ ] Обработка прайса от фазы 1 до 9
- [ ] Сохранение снимка
- [ ] Генерация инвойса

---

### Фаза 5: Webhooks и триггеры (1 неделя)

**Цель:** Настроить real-time синхронизацию.

#### 5.1 Apps Script (минимальный)

```javascript
// В каждой таблице MT/SK/SS
function onEdit(e) { sendWebhook("onEdit", e); }
function onChange(e) { sendWebhook("onChange", e); }
```

#### 5.2 Webhook Handler

```python
# src/api/webhooks.py
@router.post("/webhook/sheets/{project}")
async def handle_sheets_webhook(project, event):
    # Валидация подписи
    # Создание задачи в очереди
    # Возврат task_id
```

**Тестирование:**
- [ ] Webhook от MT таблицы
- [ ] Проверка подписи
- [ ] Асинхронная обработка
- [ ] Логирование событий

---

### Фаза 6: Тестирование и миграция (2 недели)

**Цель:** Полное тестирование и переключение.

#### 6.1 Тестовый режим

1. Создать копии таблиц (MT-test, SK-test, SS-test)
2. Направить webhooks на Python сервер
3. Параллельно оставить GAS для сравнения
4. Логировать все расхождения

#### 6.2 План переключения

```
День 1-5:   Тестирование на копиях таблиц
День 6-7:   Исправление найденных проблем
День 8:     Переключение SS (самый простой)
День 9-10:  Мониторинг SS
День 11:    Переключение SK
День 12-13: Мониторинг SK
День 14:    Переключение MT
День 15+:   Полный мониторинг
```

#### 6.3 Rollback план

Если критические проблемы:
1. Отключить webhooks в Apps Script
2. Включить обратно GAS триггеры
3. Проанализировать логи
4. Исправить и повторить

---

## Маппинг функций

### GAS → Python

| GAS функция | Python метод |
|-------------|--------------|
| `handleOnChange(e)` | `WebhookHandler.handle_sheets_event()` |
| `syncSelectedRow()` | `SyncService.sync_row()` |
| `fullSync()` | `SyncService.full_sync()` |
| `processPrice_Mt()` | `PriceProcessor.run(project="mt")` |
| `processPrice_Sk()` | `PriceProcessor.run(project="sk")` |
| `processPrice_Ss()` | `PriceProcessor.run(project="ss")` |
| `analyzeINCI()` | `GeminiClient.analyze_inci()` |
| `createFullInvoice()` | `InvoiceService.create_invoice()` |
| `collectAndCopyDocuments()` | `InvoiceService.collect_documents()` |
| `calculateAndAssignSpiritNumbers()` | `CertificationService.calculate_spirits()` |

### Листы → Конфиг

| GAS константа | YAML путь |
|---------------|-----------|
| `SHEETS.PRIMARY` | `sheets.main` |
| `SHEETS.CERTIFICATION` | `sheets.certification` |
| `SHEETS.LABELS` | `sheets.labels` |
| `SHEETS.INVOICE_FULL` | `sheets.invoice_full` |
| `SHEETS.ORDER_FORM` | `sheets.order` |
| `SHEETS.PRICE` | `sheets.price` |
| `SHEETS.PRICE_DYNAMICS` | `sheets.dynamics` |
| `SHEETS.PRICE_CALCULATION` | `sheets.calculation` |
| `SHEETS.ABC_ANALYSIS` | `sheets.abc` |
| `SHEETS.LOG` | `sheets.journal` |

---

## Риски и митигация

| Риск | Вероятность | Влияние | Митигация |
|------|-------------|---------|-----------|
| Расхождение данных | Высокая | Критическое | Параллельный запуск, сравнение |
| Rate limits Google API | Средняя | Среднее | Batch операции, кэширование |
| Сервер недоступен | Низкая | Критическое | Мониторинг, автоперезапуск |
| Ошибки парсинга | Средняя | Среднее | Unit тесты на реальных файлах |
| Потеря webhook | Низкая | Среднее | Логирование в GAS, retry |

---

## Чеклист готовности к запуску

### Инфраструктура
- [ ] VPS работает стабильно 7+ дней
- [ ] SSL сертификат валидный
- [ ] Мониторинг настроен
- [ ] Backup конфигурации

### Интеграции
- [ ] Google Sheets API работает
- [ ] Google Drive API работает
- [ ] Gemini API работает
- [ ] Webhooks доходят

### Функциональность
- [ ] Парсинг всех форматов (MT/SK/SS)
- [ ] Синхронизация между листами
- [ ] Обработка прайса (все 9 фаз)
- [ ] AI анализ INCI
- [ ] Генерация инвойсов

### Тестирование
- [ ] Unit тесты > 80% coverage
- [ ] Интеграционные тесты на копиях таблиц
- [ ] Нагрузочное тестирование (100 строк/мин)
- [ ] Rollback протестирован

---

## Timeline

```
Неделя 1:    Подготовка инфраструктуры
Неделя 2-3:  Интеграции (Sheets, Drive, Gemini)
Неделя 4:    Парсеры (MT, SK, SS)
Неделя 5-7:  Ядро синхронизации
Неделя 8-9:  Бизнес-сервисы
Неделя 10:   Webhooks
Неделя 11-12: Тестирование и миграция
─────────────────────────────────────────
Итого:       ~3 месяца
```

---

## Контакты и ответственные

| Роль | Ответственность |
|------|-----------------|
| Разработчик | Реализация кода |
| Владелец продукта | Приёмка функций |
| DevOps | Инфраструктура |
