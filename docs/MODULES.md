# AgentCare - Спецификация модулей

## Обзор модулей

```
src/
├── api/          # HTTP интерфейс
├── core/         # Ядро системы
├── services/     # Бизнес-логика
├── parsers/      # Парсеры форматов
├── integrations/ # Внешние API
└── utils/        # Утилиты
```

---

## 1. API Layer (`src/api/`)

### router.py

**Назначение:** Главный роутер FastAPI

```python
from fastapi import APIRouter, Depends
from .webhooks import webhook_router
from .endpoints import api_router

router = APIRouter()
router.include_router(webhook_router, prefix="/webhook", tags=["webhooks"])
router.include_router(api_router, prefix="/api", tags=["api"])
```

### webhooks.py

**Назначение:** Обработка webhook запросов от Google Sheets

```python
class WebhookHandler:
    """
    Обрабатывает события из Google Sheets.

    Endpoints:
        POST /webhook/sheets/{project} - событие onChange/onEdit
        POST /webhook/sync/{project}   - ручной запуск синхро

    Пример payload:
        {
            "event": "onChange",
            "project": "mt",
            "sheet": "Главная",
            "range": "A1:Z100",
            "timestamp": "2024-01-15T10:30:00Z",
            "user": "user@example.com"
        }
    """

    async def handle_sheets_event(
        self,
        project: str,
        payload: SheetEvent
    ) -> TaskResponse:
        """
        1. Валидирует payload
        2. Создает задачу в очереди
        3. Возвращает task_id для отслеживания
        """
        pass
```

### middleware.py

**Назначение:** Middleware для авторизации и логирования

```python
class WebhookAuthMiddleware:
    """Проверка подписи webhook запросов (HMAC-SHA256)"""

class RequestLoggingMiddleware:
    """Логирование всех входящих запросов"""

class RateLimitMiddleware:
    """Ограничение частоты запросов"""
```

---

## 2. Core Layer (`src/core/`)

### sync_engine.py

**Назначение:** Главный движок синхронизации

```python
class SyncEngine:
    """
    Движок синхронизации между листами Google Sheets.

    Отвечает за:
    - Определение затронутых листов при изменении
    - Координацию обновлений между листами
    - Управление порядком операций
    - Обработку конфликтов

    Пример использования:
        engine = SyncEngine(project="mt")
        result = await engine.process_change(
            source_sheet="Главная",
            changed_rows=[5, 10, 15],
            changed_columns=["Цена EXW", "Статус"]
        )
    """

    def __init__(self, project: str, config: SyncConfig):
        self.project = project
        self.config = config
        self.rules_engine = RulesEngine(config.rules)
        self.sheets_client = GoogleSheetsClient(config.credentials)

    async def process_change(
        self,
        source_sheet: str,
        changed_rows: List[int],
        changed_columns: List[str]
    ) -> SyncResult:
        """
        Обрабатывает изменение в листе.

        1. Загружает затронутые данные
        2. Определяет целевые листы
        3. Применяет правила трансформации
        4. Выполняет batch update
        5. Логирует результат
        """
        pass

    async def full_sync(self, source_sheet: str) -> SyncResult:
        """Полная синхронизация листа (для первичной загрузки)"""
        pass

    def get_dependent_sheets(self, source: str) -> List[str]:
        """Возвращает список листов, зависящих от source"""
        pass
```

### rules_engine.py

**Назначение:** Движок каскадных правил

```python
class RulesEngine:
    """
    Применяет правила трансформации к данным.

    Типы правил:
    - copy: копирование значения в другой столбец
    - transform: преобразование значения (формулы, замены)
    - validate: проверка значения на соответствие
    - cascade: цепочка зависимых обновлений

    Формат правила (YAML):
        - name: "copy_price_to_invoice"
          type: copy
          source: "Главная.Цена EXW"
          target: "Инвойс.Цена"
          condition: "row.Статус == 'Активный'"
    """

    def __init__(self, rules_config: RulesConfig):
        self.rules = self._load_rules(rules_config)

    def apply(
        self,
        data: DataFrame,
        context: RuleContext
    ) -> Tuple[DataFrame, List[RuleResult]]:
        """
        Применяет все подходящие правила к данным.

        Returns:
            - Трансформированные данные
            - Список примененных правил с результатами
        """
        pass

    def validate(self, data: DataFrame) -> ValidationResult:
        """Проверяет данные на соответствие правилам валидации"""
        pass
```

### transaction.py

**Назначение:** Управление транзакциями

```python
class Transaction:
    """
    Группирует операции для атомарного выполнения.

    Возможности:
    - Группировка нескольких операций
    - Откат при ошибке
    - Логирование всех изменений
    - Восстановление после сбоя

    Пример:
        async with Transaction(sheets_client) as tx:
            await tx.update("Главная", changes1)
            await tx.update("Инвойс", changes2)
            # Если ошибка - откат обоих изменений
    """

    def __init__(self, client: GoogleSheetsClient):
        self.client = client
        self.operations: List[Operation] = []
        self.snapshots: Dict[str, DataFrame] = {}

    async def __aenter__(self):
        """Начало транзакции - создание снимков"""
        pass

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        """
        Завершение транзакции:
        - При успехе: commit
        - При ошибке: rollback
        """
        pass

    async def update(self, sheet: str, changes: DataFrame):
        """Добавляет операцию обновления в транзакцию"""
        pass

    async def rollback(self):
        """Откатывает все операции используя снимки"""
        pass
```

### scheduler.py

**Назначение:** Планировщик фоновых задач

```python
class Scheduler:
    """
    Управляет фоновыми задачами.

    Задачи:
    - Периодическая синхронизация
    - Очистка кэша
    - Очистка логов
    - Проверка здоровья интеграций

    Использует: APScheduler или Celery Beat
    """

    def schedule_sync(
        self,
        project: str,
        interval_minutes: int = 15
    ):
        """Планирует периодическую синхронизацию"""
        pass

    def schedule_cleanup(self, interval_hours: int = 24):
        """Планирует очистку старых данных"""
        pass
```

---

## 3. Services Layer (`src/services/`)

### price_processor.py

**Назначение:** Обработка прайсов (9 фаз)

```python
class PriceProcessor:
    """
    Обрабатывает прайс поставщика.

    9 фаз обработки:
    1. load_excel       - Загрузка из внешнего Excel
    2. clear_columns    - Очистка столбцов (ID-P, цены, ID-G)
    3. sync_with_base   - Синхронизация с базой
    4. save_snapshot    - Сохранение снимка (лист с датой)
    5. copy_prices      - Копирование цен EXW
    6. fill_ids         - Заполнение ID-P на все листы
    7. process_new      - Обработка новых артикулов
    8. apply_formulas   - Применение формул INDEX/MATCH
    9. update_statuses  - Обновление статусов

    Каждая фаза может быть перезапущена отдельно.
    """

    def __init__(self, project: str, config: ProjectConfig):
        self.project = project
        self.parser = self._get_parser(project)
        self.sync_engine = SyncEngine(project, config.sync)

    async def run(
        self,
        file_path: str,
        start_phase: int = 1
    ) -> ProcessingResult:
        """
        Запускает обработку с указанной фазы.

        При ошибке сохраняет состояние для продолжения.
        """
        phases = [
            self._phase_load_excel,
            self._phase_clear_columns,
            self._phase_sync_with_base,
            self._phase_save_snapshot,
            self._phase_copy_prices,
            self._phase_fill_ids,
            self._phase_process_new,
            self._phase_apply_formulas,
            self._phase_update_statuses,
        ]

        for i, phase in enumerate(phases[start_phase - 1:], start_phase):
            try:
                await phase()
                self._save_progress(i)
            except Exception as e:
                return ProcessingResult(
                    success=False,
                    failed_phase=i,
                    error=str(e)
                )

        return ProcessingResult(success=True)
```

### sync_service.py

**Назначение:** Сервис синхронизации

```python
class SyncService:
    """
    Высокоуровневый сервис синхронизации.

    Методы:
    - sync_row: синхронизация одной строки между листами
    - sync_range: синхронизация диапазона
    - sync_column: синхронизация столбца
    - full_sync: полная синхронизация листа
    """

    async def sync_row(
        self,
        project: str,
        article: str,
        source_sheet: str,
        target_sheets: List[str] = None
    ) -> SyncResult:
        """
        Синхронизирует строку по артикулу.

        1. Находит строку в source_sheet
        2. Определяет target листы (или использует переданные)
        3. Находит/создает строку в каждом target
        4. Применяет правила трансформации
        5. Обновляет данные
        """
        pass

    async def sync_range(
        self,
        project: str,
        source_sheet: str,
        range_str: str
    ) -> SyncResult:
        """Синхронизирует диапазон ячеек"""
        pass
```

### invoice_service.py

**Назначение:** Работа с инвойсами

```python
class InvoiceService:
    """
    Сервис работы с инвойсами.

    Функции:
    - Создание инвойса из выбранных строк
    - Генерация документа из шаблона
    - Расчет сумм и налогов
    - Сохранение в Google Drive
    """

    async def create_invoice(
        self,
        project: str,
        rows: List[int],
        template: str = "default"
    ) -> Invoice:
        """Создает инвойс из выбранных строк"""
        pass

    async def generate_document(
        self,
        invoice: Invoice
    ) -> str:
        """Генерирует документ и возвращает URL"""
        pass
```

### order_service.py

**Назначение:** Автозаказ

```python
class OrderService:
    """
    Сервис автозаказа.

    Логика:
    - Анализ остатков и продаж
    - Расчет оптимального заказа
    - Формирование заявки поставщику
    """

    async def calculate_order(
        self,
        project: str,
        period_days: int = 30
    ) -> OrderRecommendation:
        """Рассчитывает рекомендуемый заказ"""
        pass

    async def create_order(
        self,
        recommendation: OrderRecommendation,
        adjustments: Dict[str, int] = None
    ) -> Order:
        """Создает заказ с возможными корректировками"""
        pass
```

---

## 4. Parsers Layer (`src/parsers/`)

### base_parser.py

**Назначение:** Базовый класс парсера

```python
from abc import ABC, abstractmethod
from pandas import DataFrame

class BaseParser(ABC):
    """
    Базовый класс для парсеров Excel.

    Каждый проект (MT/SK/SS) наследует этот класс
    и реализует свою логику извлечения данных.
    """

    # Маппинг колонок Excel -> внутренние имена
    COLUMN_MAPPING: Dict[str, str] = {}

    # Обязательные поля
    REQUIRED_FIELDS: List[str] = ["article", "name", "price"]

    @abstractmethod
    def parse(self, file_path: str) -> DataFrame:
        """
        Парсит Excel файл и возвращает DataFrame.

        Должен:
        1. Загрузить файл
        2. Найти нужный лист
        3. Извлечь данные
        4. Применить маппинг колонок
        """
        pass

    def validate(self, data: DataFrame) -> ValidationResult:
        """
        Проверяет данные на корректность.

        - Наличие обязательных полей
        - Типы данных
        - Дубликаты артикулов
        """
        errors = []

        for field in self.REQUIRED_FIELDS:
            if field not in data.columns:
                errors.append(f"Missing required field: {field}")

        duplicates = data[data.duplicated("article", keep=False)]
        if not duplicates.empty:
            errors.append(f"Duplicate articles: {duplicates['article'].tolist()}")

        return ValidationResult(
            valid=len(errors) == 0,
            errors=errors
        )

    def transform(self, data: DataFrame) -> DataFrame:
        """
        Трансформирует данные в стандартный формат.

        - Нормализация названий
        - Конвертация типов
        - Очистка данных
        """
        pass
```

### mt_parser.py

**Назначение:** Парсер формата MT (CosmeticaBar)

```python
class MTParser(BaseParser):
    """
    Парсер для формата MT.

    Особенности MT:
    - Есть тестеры и пробники
    - Специфичные колонки: BAR CODE, Форма выпуска
    - ID-G для группировки
    """

    COLUMN_MAPPING = {
        "Артикул": "article",
        "Наименование": "name",
        "Цена EXW": "price_exw",
        "BAR CODE": "barcode",
        "Форма выпуска": "form",
        "Объем": "volume",
        "ID-G": "group_id",
    }

    def parse(self, file_path: str) -> DataFrame:
        """Парсит MT формат"""
        df = pd.read_excel(file_path, sheet_name="Прайс")
        df = df.rename(columns=self.COLUMN_MAPPING)
        return self.transform(df)

    def _extract_testers(self, data: DataFrame) -> DataFrame:
        """Извлекает тестеры в отдельный DataFrame"""
        pass
```

### sk_parser.py

**Назначение:** Парсер формата SK (Carmado)

```python
class SKParser(BaseParser):
    """
    Парсер для формата SK.

    Особенности SK:
    - Пробники с RRP
    - Скидки по категориям
    - Специфичная структура файла
    """

    COLUMN_MAPPING = {
        "Код": "article",
        "Название": "name",
        "Цена": "price_exw",
        "RRP": "rrp",
        "Скидка %": "discount",
    }
```

### ss_parser.py

**Назначение:** Парсер формата SS (Сан)

```python
class SSParser(BaseParser):
    """
    Парсер для формата SS.

    Особенности SS:
    - Базовый прайс без сложной логики
    - Минимальный набор полей
    """

    COLUMN_MAPPING = {
        "Артикул": "article",
        "Наименование": "name",
        "Цена": "price_exw",
    }
```

---

## 5. Integrations Layer (`src/integrations/`)

### google_sheets.py

**Назначение:** Работа с Google Sheets API

```python
class GoogleSheetsClient:
    """
    Клиент для Google Sheets API.

    Оптимизации:
    - Batch операции (до 100 запросов за раз)
    - Кэширование метаданных листов
    - Retry при ошибках квот

    Использует: gspread + google-auth
    """

    def __init__(self, credentials_path: str):
        self.gc = gspread.service_account(filename=credentials_path)
        self._cache = {}

    async def read_sheet(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        range_str: str = None
    ) -> DataFrame:
        """
        Читает данные из листа.

        Args:
            spreadsheet_id: ID таблицы
            sheet_name: Имя листа
            range_str: Диапазон (A1:Z100) или None для всего листа
        """
        pass

    async def write_sheet(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        data: DataFrame,
        range_str: str = None
    ):
        """Записывает данные в лист"""
        pass

    async def batch_update(
        self,
        spreadsheet_id: str,
        updates: List[SheetUpdate]
    ):
        """
        Batch обновление нескольких диапазонов.

        Эффективнее чем отдельные запросы.
        """
        pass

    async def find_row_by_value(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        column: str,
        value: str
    ) -> Optional[int]:
        """Находит номер строки по значению в столбце"""
        pass
```

### google_drive.py

**Назначение:** Работа с Google Drive API

```python
class GoogleDriveClient:
    """
    Клиент для Google Drive API.

    Функции:
    - Создание папок
    - Копирование документов из шаблонов
    - Конвертация PDF в текст
    - Управление правами доступа
    """

    async def create_folder(
        self,
        name: str,
        parent_id: str = None
    ) -> str:
        """Создает папку и возвращает её ID"""
        pass

    async def copy_template(
        self,
        template_id: str,
        new_name: str,
        folder_id: str
    ) -> str:
        """Копирует шаблон документа"""
        pass

    async def pdf_to_text(self, file_id: str) -> str:
        """
        Конвертирует PDF в текст.

        Использует Google Drive для OCR.
        """
        pass
```

### gemini_client.py

**Назначение:** Интеграция с Gemini AI

```python
class GeminiClient:
    """
    Клиент для Google Gemini API.

    Функции:
    - Анализ INCI состава
    - Определение кода ТН ВЭД
    - Классификация продуктов
    - Извлечение данных из PDF

    Retry логика:
    - 5 попыток
    - Экспоненциальная задержка: 2, 5, 10, 20, 30 сек
    - Обработка 429 (rate limit) и 503 (overload)
    """

    def __init__(self, api_key: str, model: str = "gemini-2.5-flash"):
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel(model)
        self.retry_delays = [2, 5, 10, 20, 30]

    @retry_on_rate_limit
    async def analyze_inci(
        self,
        text: str,
        product_name: str = None
    ) -> INCIAnalysis:
        """
        Анализирует INCI состав.

        Возвращает:
        - tn_ved_code: код ТН ВЭД (3304, 3305, 3307)
        - product_type: тип продукта
        - active_ingredients: активные ингредиенты
        - regulatory_doc: СГР или Декларация
        """
        prompt = self._build_inci_prompt(text, product_name)
        response = await self.model.generate_content_async(prompt)
        return self._parse_inci_response(response.text)

    async def extract_from_pdf(
        self,
        pdf_text: str
    ) -> Dict[str, Any]:
        """Извлекает структурированные данные из текста PDF"""
        pass
```

### notification.py

**Назначение:** Уведомления

```python
class NotificationService:
    """
    Сервис уведомлений.

    Каналы:
    - Telegram
    - Email
    - Webhook
    """

    async def send_alert(
        self,
        level: str,  # "info", "warning", "error", "critical"
        message: str,
        details: Dict = None
    ):
        """Отправляет алерт во все активные каналы"""
        pass

    async def send_telegram(self, chat_id: str, message: str):
        """Отправляет сообщение в Telegram"""
        pass

    async def send_email(self, to: str, subject: str, body: str):
        """Отправляет email"""
        pass
```

---

## 6. Utils Layer (`src/utils/`)

### cache.py

**Назначение:** Кэширование

```python
class CacheService:
    """
    Сервис кэширования.

    Backends:
    - Redis (для production)
    - In-memory (для development/testing)

    Используется для:
    - Кэширование номеров строк по артикулам
    - Кэширование метаданных листов
    - Кэширование результатов AI анализа
    """

    def __init__(self, backend: str = "redis", ttl: int = 3600):
        self.backend = self._create_backend(backend)
        self.ttl = ttl

    async def get(self, key: str) -> Optional[Any]:
        """Получает значение из кэша"""
        pass

    async def set(self, key: str, value: Any, ttl: int = None):
        """Сохраняет значение в кэш"""
        pass

    async def delete(self, key: str):
        """Удаляет значение из кэша"""
        pass

    async def clear_pattern(self, pattern: str):
        """Удаляет все ключи по паттерну"""
        pass
```

### logger.py

**Назначение:** Логирование

```python
import structlog

def setup_logging(level: str = "INFO"):
    """
    Настраивает структурированное логирование.

    Формат логов:
    {
        "timestamp": "2024-01-15T10:30:00Z",
        "level": "info",
        "event": "sync_completed",
        "project": "mt",
        "duration_ms": 1234,
        ...
    }
    """
    structlog.configure(
        processors=[
            structlog.processors.TimeStamper(fmt="iso"),
            structlog.processors.JSONRenderer()
        ],
        wrapper_class=structlog.make_filtering_bound_logger(level),
    )

logger = structlog.get_logger()

# Использование:
# logger.info("sync_completed", project="mt", rows=150)
# logger.error("sync_failed", error=str(e), traceback=traceback)
```

### retry.py

**Назначение:** Retry логика

```python
from functools import wraps

def retry(
    max_attempts: int = 3,
    delay: float = 1.0,
    backoff: float = 2.0,
    exceptions: tuple = (Exception,)
):
    """
    Декоратор для автоматического retry.

    Args:
        max_attempts: максимум попыток
        delay: начальная задержка (сек)
        backoff: множитель задержки
        exceptions: какие исключения перехватывать

    Пример:
        @retry(max_attempts=5, delay=2, backoff=2)
        async def call_api():
            ...
    """
    def decorator(func):
        @wraps(func)
        async def wrapper(*args, **kwargs):
            last_exception = None
            current_delay = delay

            for attempt in range(max_attempts):
                try:
                    return await func(*args, **kwargs)
                except exceptions as e:
                    last_exception = e
                    if attempt < max_attempts - 1:
                        await asyncio.sleep(current_delay)
                        current_delay *= backoff

            raise last_exception

        return wrapper
    return decorator
```

### validators.py

**Назначение:** Валидаторы данных

```python
from pydantic import BaseModel, validator

class ArticleData(BaseModel):
    """Валидация данных артикула"""

    article: str
    name: str
    price_exw: float
    barcode: Optional[str]

    @validator("article")
    def validate_article(cls, v):
        if not v or len(v) < 3:
            raise ValueError("Article must be at least 3 characters")
        return v.strip().upper()

    @validator("price_exw")
    def validate_price(cls, v):
        if v < 0:
            raise ValueError("Price cannot be negative")
        return round(v, 2)

class SyncRuleValidator:
    """Валидация правил синхронизации"""

    @staticmethod
    def validate_rule(rule: dict) -> ValidationResult:
        required = ["name", "type", "source", "target"]
        ...
```

---

## Зависимости между модулями

```
                    ┌─────────┐
                    │   API   │
                    └────┬────┘
                         │
              ┌──────────┼──────────┐
              │          │          │
              ▼          ▼          ▼
         ┌────────┐ ┌────────┐ ┌────────┐
         │Services│ │  Core  │ │Parsers │
         └───┬────┘ └───┬────┘ └───┬────┘
             │          │          │
             └──────────┼──────────┘
                        │
              ┌─────────┼─────────┐
              │         │         │
              ▼         ▼         ▼
         ┌────────────────────────────┐
         │       Integrations         │
         │  (Sheets, Drive, Gemini)   │
         └─────────────┬──────────────┘
                       │
                       ▼
                  ┌─────────┐
                  │  Utils  │
                  │(cache,  │
                  │ logger, │
                  │ retry)  │
                  └─────────┘
```

**Правило:** Зависимости только сверху вниз. Utils не зависит ни от чего. Integrations зависит только от Utils.
