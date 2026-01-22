# AgentCare System Architecture

**Version:** 3.0 (Cleanup Phase)
**Last Updated:** 2026-01-15
**Handbook:** [AI_HANDBOOK.md](./AI_HANDBOOK.md)

---

## 🏗 System Overview

AgentCare — это гибридная система автоматизации для бизнеса, построенная на принципе **Server-First**.

### Core Pillars
1.  **Python Server (FastAPI)**: "Мозг" системы. Вся бизнес-логика здесь.
2.  **Supabase (PostgreSQL)**: "Память" системы. Хранит логи, настройки и кэш данных.
3.  **Google Sheets (GAS)**: "Лицо" системы. Только отображение и сбор событий.
4.  **AI (Gemini)**: "Интеллект". Анализирует данные по запросу.

### 🔄 The "Journal" Workflow (Debug Philosophy)
Главный инструмент поддержки — **Журнал (Logs)**.
- **Philosophy**: Логирование должно быть максимально подробным для каждой функции (даже самой мелкой).
- **Problem**: "Кнопка не работает" или "Ошибка в расчетах".
- **Solution**: Агент/Пользователь идет в **Журнал UI** (серверный интерфейс) или смотрит логи сервера, находит причину, исправляет.
- **Rule**: Никаких "слепых" правок. Сначала лог, потом фикс.

---

## 🧩 Components

### 1. Server Layer (`src/`)
- **FastAPI**: REST API для GAS и внешних хуков.
- **Services**:
  - `SyncService`: Синхронизация данных (Sheets <-> Supabase <-> Sheets).
  - `LoggingService`: Запись событий в БД.
  - `SupabaseService`: Работа с базой данных.
  - `AIService`: Интеграция с Gemini.

### 2. Client Layer (`gas/`)
- **Google Apps Script**: Тонкий клиент.
- **Задача**:
  1. Поймать `onEdit` или клик по меню.
  2. Отправить `POST` на сервер.
  3. Показать результат (Toast/Alert) или обновить ячейки по ответу сервера.
- **ЗАПРЕЩЕНО**: Писать сложную логику обработки данных внутри GAS.

### 3. Data Layer
- **Supabase**: Хранит конфигурацию, очередь задач, логи.
- **Google Sheets**: Пользовательский интерфейс.

---

## 🔄 Data Flow

### Типичная Синхронизация (onEdit)
1.  **User** меняет ячейку в Google Sheet.
2.  **GAS** (`onEdit` trigger) собирает данные: `{sheet, row, col, val}`.
3.  **GAS** отправляет `POST /api/v1/sync/event` на сервер.
4.  **Server**:
    - Валидирует запрос.
    - Смотрит правила синхронизации (из БД/Config).
    - Обновляет данные в других листах (через Google Sheets API).
    - Пишет лог в **Supabase**.
5.  **Server** возвращает статус `200 OK`.
6.  **GAS** показывает всплывающее уведомление (Toast).

---

## 📍 Key Directories

| Directory | Purpose |
|-----------|---------|
| `src/` | Исходный код Python сервера |
| `gas/` | Исходный код Google Apps Script |
| `docs/` | Вся документация (API, Деплой) |
| `.beads/` | База данных задач (Task Tracking) |

---

## 🔗 Useful Links
- [AI Handbook (Правила)](./AI_HANDBOOK.md)
- [API Documentation](./docs/API.md)
- [Legacy Docs](./docs/archive/)
