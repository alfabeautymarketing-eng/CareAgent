# API Reference - AgentCare

**Base URL:** `http://46.226.167.153:8000`  
**API Version:** v1  
**Last Updated:** 2026-01-07

---

## Authentication

Все запросы от Google Apps Script проходят без дополнительной аутентификации.  
Server использует Service Account для доступа к Google Sheets/Drive API.

---

## Common Response Format

**Success Response:**
```json
{
  "status": "success",
  "message": "Операция выполнена",
  "data": { }
}
```

**Error Response:**
```json
{
  "detail": "Описание ошибки"
}
```

---

## Sync Endpoints

### POST /api/v1/sync/event

Обработка onEdit события из Google Sheets.

**Request Body:**
```json
{
  "spreadsheet_id": "1ABC...",
  "sheet_name": "Главная",
  "row": 5,
  "col": 3,
  "value": "Новое значение",
  "old_value": "Старое значение",
  "user_email": "user@example.com",
  "header_name": "Наименование",
  "row_key": "MT-001",
  "sync_origin": "user",
  "transaction_id": "tx_123"
}
```

**Response:**
```json
{
  "status": "queued",
  "message": "Sync queued"
}
```

---

### POST /api/v1/sync/add-article

Добавить новый артикул.

**Request:**
```json
{
  "project": "MT",
  "article": "",
  "spreadsheet_id": "1ABC..."
}
```

**Response:**
```json
{
  "status": "success",
  "message": "Артикул обработан",
  "details": {
    "article": "MT-157",
    "added_to_sheets": ["Главная", "Информация", "Сертификация"]
  }
}
```

---

## Rules Management

### GET /api/v1/rules/{spreadsheet_id}

Получить все правила синхронизации.

**Response:**
```json
{
  "rules": [
    {
      "id": "rule-001",
      "enabled": true,
      "mode": "unidirectional",
      "category": "Основные данные",
      "source_sheet": "Главная",
      "source_header": "Наименование",
      "target_sheet": "Информация",
      "target_header": "Название продукта"
    }
  ]
}
```

---

### POST /api/v1/rules/{spreadsheet_id}/create

Создать новое правило.

**Request:**
```json
{
  "mode": "bidirectional",
  "enabled": true,
  "category": "Цены",
  "sheet_a": "Главная",
  "header_a": "Цена закупки",
  "sheet_b": "Расчет цены",
  "header_b": "Закупка"
}
```

---

## Logs Management

### POST /api/v1/logs/archive

Архивировать логи в отдельный spreadsheet.

**Request:**
```json
{
  "spreadsheet_id": "1ABC..."
}
```

**Response:**
```json
{
  "status": "success",
  "message": "Архивирование завершено!",
  "data": {
    "archive_name": "mt-январь-2026",
    "archive_id": "1XYZ...",
    "folder_url": "https://drive.google.com/...",
    "archived_sheets": ["Логи", "Журнал синхро", "Журнал логов"],
    "total_rows": 1500
  }
}
```

---

### GET /api/v1/logs/archive/status

Проверить статус архива за текущий месяц.

**Response:**
```json
{
  "status": "success",
  "data": {
    "archive_name": "mt-январь-2026",
    "archive_id": "1XYZ...",
    "folder_id": "1oWoxwuOMlriwgHuua6MpLheROEIT4SIP",
    "exists": true,
    "last_modified": "2026-01-07T10:30:00Z"
  }
}
```

---

## Documents

### POST /api/v1/documents/collect

Собрать документы для инвойса.

**Request:**
```json
{
  "spreadsheet_id": "1ABC...",
  "target_sheet": "Для инвойса"
}
```

**Response:**
```json
{
  "status": "success",
  "data": {
    "folder_id": "1DjrtDKLUMCypFqlC38kjgOESupTjjUpK",
    "folder_url": "https://drive.google.com/...",
    "files_copied": 15,
    "details": [
      {"article": "MT-001", "file": "DS_MT-001.pdf", "status": "copied"},
      {"article": "MT-002", "file": "Spirit_MT-002.pdf", "status": "copied"}
    ]
  }
}
```

---

### POST /api/v1/documents/structure-353pp

Структурировать документы для заявки 353пп.

**Response:**
```json
{
  "status": "success",
  "data": {
    "processed": 5,
    "folder_url": "https://drive.google.com/folders/...",
    "details": []
  }
}
```

---

## Price Processing

### POST /api/v1/price/process

Обработать прайс-лист.

**Request:**
```json
{
  "spreadsheet_id": "1ABC...",
  "mode": "main",
  "source_doc_id": null,
  "dry_run": false
}
```

**Modes:**
- `main` - основной прайс
- `tester` - тестеры
- `samples` - пробники  
- `probes` - пробы

---

## AI Analysis

### POST /api/v1/ai/analyze

Анализ продукта с помощью Gemini AI.

**Request:**
```json
{
  "spreadsheet_id": "1ABC...",
  "sheet_name": "Информация",
  "row_number": 5,
  "pdf_url": "https://drive.google.com/...",
  "purpose": null,
  "application": null
}
```

**Response:**
```json
{
  "status": "success",
  "analysis": {
    "purpose": "Увлажнение кожи",
    "application": "Крема, лосьоны",
    "inci_analysis": "Гиалуроновая кислота..."
  }
}
```

---

## Utility

### GET /health

Health check endpoint.

**Response:**
```json
{
  "status": "healthy",
  "version": "2.0.0",
  "timestamp": "2026-01-07T14:30:00Z"
}
```

---

## Error Codes

| Code | Meaning | Description |
|------|---------|-------------|
| 200 | OK | Успешно |
| 400 | Bad Request | Неверные параметры |
| 404 | Not Found | Ресурс не найден |
| 500 | Internal Server Error | Ошибка сервера |

---

## Rate Limits

Нет жёстких ограничений, но рекомендуется:
- Batch операции вместо множества одиночных
- Delay между AI запросами (2+ секунды)

---

## Examples

### Full Sync Flow (GAS → Server)

```javascript
// GAS Client.gs
function callServerSyncEvent(eventData) {
  const url = SERVER_URL + '/api/v1/sync/event';
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(eventData),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}
```

---

## References

- [Architecture](./ARCHITECTURE.md)
- [Function Map](./FUNCTION_MAP.md)
