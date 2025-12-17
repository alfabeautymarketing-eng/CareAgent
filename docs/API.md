# AgentCare - API Specification

## Обзор

REST API для управления синхронизацией Google Sheets и AI-анализом.

**Base URL:** `https://your-server.com/api/v1`

---

## Аутентификация

### API Key (для внешних запросов)

```http
Authorization: Bearer <api_key>
```

### Webhook Signature (для Google Sheets)

```http
X-Webhook-Signature: sha256=<hmac_signature>
X-Webhook-Timestamp: <unix_timestamp>
```

---

## Endpoints

### Health & Status

#### `GET /health`

Проверка здоровья сервиса.

**Response:**
```json
{
  "status": "healthy",
  "version": "0.1.0",
  "uptime_seconds": 86400,
  "checks": {
    "redis": "ok",
    "google_sheets": "ok",
    "gemini": "ok"
  }
}
```

#### `GET /status/{task_id}`

Статус задачи.

**Response:**
```json
{
  "task_id": "abc123",
  "status": "completed",  // pending, running, completed, failed
  "created_at": "2024-01-15T10:00:00Z",
  "completed_at": "2024-01-15T10:00:05Z",
  "result": {
    "rows_processed": 150,
    "sheets_updated": 5
  }
}
```

---

### Webhooks (от Google Sheets)

#### `POST /webhook/sheets/{project}`

Обработка события из Google Sheets.

**Path Parameters:**
- `project` — код проекта (`mt`, `sk`, `ss`)

**Request Body:**
```json
{
  "event": "onChange",
  "sheet": "Главная",
  "range": "A5:Z5",
  "changed_columns": ["Цена EXW", "Статус"],
  "user": "user@example.com",
  "timestamp": "2024-01-15T10:30:00Z"
}
```

**Response:**
```json
{
  "task_id": "abc123",
  "status": "accepted",
  "message": "Sync task queued"
}
```

**Errors:**
- `400` — Invalid payload
- `401` — Invalid signature
- `404` — Unknown project
- `429` — Rate limit exceeded

---

### Синхронизация

#### `POST /sync/row`

Синхронизация одной строки.

**Request:**
```json
{
  "project": "mt",
  "article": "ABC123",
  "source_sheet": "Главная",
  "target_sheets": ["Инвойс", "Прайс"]  // опционально, по умолчанию все
}
```

**Response:**
```json
{
  "task_id": "abc123",
  "status": "completed",
  "result": {
    "article": "ABC123",
    "sheets_updated": ["Инвойс", "Прайс"],
    "fields_changed": 12
  }
}
```

#### `POST /sync/range`

Синхронизация диапазона.

**Request:**
```json
{
  "project": "mt",
  "source_sheet": "Главная",
  "range": "A1:Z100"
}
```

#### `POST /sync/full`

Полная синхронизация листа.

**Request:**
```json
{
  "project": "mt",
  "source_sheet": "Главная",
  "async": true  // запустить в фоне
}
```

**Response (async=true):**
```json
{
  "task_id": "abc123",
  "status": "accepted",
  "message": "Full sync started in background"
}
```

---

### Обработка прайсов

#### `POST /price/process`

Запуск обработки прайса (9 фаз).

**Request:**
```json
{
  "project": "mt",
  "file_url": "https://drive.google.com/...",  // или file_id
  "start_phase": 1,  // с какой фазы начать (для retry)
  "options": {
    "skip_snapshot": false,
    "dry_run": false
  }
}
```

**Response:**
```json
{
  "task_id": "abc123",
  "status": "accepted",
  "estimated_duration_seconds": 120
}
```

#### `GET /price/status/{task_id}`

Статус обработки прайса.

**Response:**
```json
{
  "task_id": "abc123",
  "status": "running",
  "current_phase": 5,
  "total_phases": 9,
  "phase_name": "copy_prices",
  "progress_percent": 55,
  "rows_processed": 500,
  "total_rows": 1000
}
```

#### `POST /price/cancel/{task_id}`

Отмена обработки.

**Response:**
```json
{
  "task_id": "abc123",
  "status": "cancelled",
  "last_completed_phase": 4
}
```

---

### AI Анализ

#### `POST /ai/analyze/inci`

Анализ INCI состава.

**Request:**
```json
{
  "project": "mt",
  "row_number": 42,  // или article
  "pdf_url": "https://drive.google.com/...",
  "product_name": "Крем для лица"  // опционально, улучшает точность
}
```

**Response:**
```json
{
  "task_id": "abc123",
  "status": "completed",
  "result": {
    "tn_ved_code": "3304",
    "product_type": "Крем косметический",
    "active_ingredients": [
      {"name": "Hyaluronic Acid", "percentage": "2%"},
      {"name": "Niacinamide", "percentage": null}
    ],
    "regulatory_doc": "СГР",
    "confidence": 0.95
  }
}
```

#### `POST /ai/analyze/batch`

Batch анализ нескольких строк.

**Request:**
```json
{
  "project": "mt",
  "rows": [42, 43, 44, 45],
  "options": {
    "skip_existing": true,  // пропустить уже проанализированные
    "parallel": 3  // параллельных запросов
  }
}
```

**Response:**
```json
{
  "task_id": "abc123",
  "status": "accepted",
  "total_rows": 4,
  "estimated_duration_seconds": 30
}
```

---

### Инвойсы

#### `POST /invoice/create`

Создание инвойса.

**Request:**
```json
{
  "project": "mt",
  "rows": [10, 15, 20, 25],
  "template": "default",
  "options": {
    "include_samples": true,
    "currency": "EUR"
  }
}
```

**Response:**
```json
{
  "invoice_id": "INV-2024-001",
  "document_url": "https://docs.google.com/...",
  "total_amount": 1500.00,
  "currency": "EUR",
  "items_count": 4
}
```

#### `GET /invoice/{invoice_id}`

Получение инвойса.

**Response:**
```json
{
  "invoice_id": "INV-2024-001",
  "project": "mt",
  "created_at": "2024-01-15T10:00:00Z",
  "status": "draft",
  "document_url": "https://docs.google.com/...",
  "items": [
    {
      "article": "ABC123",
      "name": "Product Name",
      "quantity": 10,
      "price": 15.00,
      "total": 150.00
    }
  ],
  "subtotal": 1500.00,
  "tax": 0,
  "total": 1500.00
}
```

---

### Автозаказ

#### `POST /order/calculate`

Расчет рекомендуемого заказа.

**Request:**
```json
{
  "project": "mt",
  "period_days": 30,
  "min_stock_days": 14,
  "options": {
    "include_new": false,
    "category_filter": ["Уход", "Макияж"]
  }
}
```

**Response:**
```json
{
  "recommendation_id": "REC-001",
  "total_items": 45,
  "total_amount": 5000.00,
  "items": [
    {
      "article": "ABC123",
      "name": "Product",
      "current_stock": 5,
      "avg_sales_per_day": 2,
      "recommended_qty": 50,
      "reason": "Low stock, high demand"
    }
  ]
}
```

#### `POST /order/create`

Создание заказа.

**Request:**
```json
{
  "recommendation_id": "REC-001",
  "adjustments": {
    "ABC123": 40,  // изменить кол-во
    "DEF456": 0    // исключить
  }
}
```

---

### Кэш

#### `POST /cache/clear`

Очистка кэша.

**Request:**
```json
{
  "project": "mt",  // опционально, если не указан - весь кэш
  "pattern": "row_*"  // опционально, паттерн ключей
}
```

**Response:**
```json
{
  "cleared_keys": 150,
  "message": "Cache cleared successfully"
}
```

---

## Error Responses

Все ошибки возвращаются в едином формате:

```json
{
  "error": {
    "code": "VALIDATION_ERROR",
    "message": "Invalid project code",
    "details": {
      "field": "project",
      "value": "unknown",
      "allowed": ["mt", "sk", "ss"]
    }
  },
  "request_id": "req-123"
}
```

### Коды ошибок

| HTTP | Code | Описание |
|------|------|----------|
| 400 | `VALIDATION_ERROR` | Невалидные данные |
| 401 | `UNAUTHORIZED` | Неверная авторизация |
| 403 | `FORBIDDEN` | Нет доступа |
| 404 | `NOT_FOUND` | Ресурс не найден |
| 409 | `CONFLICT` | Конфликт (напр. дубликат) |
| 429 | `RATE_LIMITED` | Превышен лимит запросов |
| 500 | `INTERNAL_ERROR` | Внутренняя ошибка |
| 502 | `GOOGLE_API_ERROR` | Ошибка Google API |
| 503 | `SERVICE_UNAVAILABLE` | Сервис недоступен |

---

## Rate Limits

| Endpoint | Лимит |
|----------|-------|
| `/webhook/*` | 100 req/min |
| `/sync/*` | 20 req/min |
| `/price/*` | 5 req/min |
| `/ai/*` | 30 req/min |
| `/invoice/*` | 10 req/min |

**Headers:**
```http
X-RateLimit-Limit: 100
X-RateLimit-Remaining: 95
X-RateLimit-Reset: 1705312800
```

---

## WebSocket (опционально)

### `WS /ws/tasks`

Realtime обновления статуса задач.

**Subscribe:**
```json
{
  "action": "subscribe",
  "task_ids": ["abc123", "def456"]
}
```

**Updates:**
```json
{
  "type": "task_update",
  "task_id": "abc123",
  "status": "running",
  "progress_percent": 75
}
```

---

## OpenAPI

Автоматическая документация доступна:
- Swagger UI: `/docs`
- ReDoc: `/redoc`
- OpenAPI JSON: `/openapi.json`
