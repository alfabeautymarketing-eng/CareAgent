# MCP Servers Configuration & Setup

**Дата обновления:** 2026-01-14

## 📋 Overview

Этот документ описывает настройку MCP (Model Context Protocol) серверов для AgentCare проекта. MCP серверы позволяют Claude AI взаимодействовать с различными внешними сервисами и инструментами.

---

## 🔗 Подключенные MCP Servers

### 1. **Supabase** ✨ NEW
- **Тип:** Database Management
- **Статус:** ✅ Активен
- **Назначение:** Управление PostgreSQL базой данных в Supabase
- **Подключение:** npx @supabase/mcp-server
- **Переменные:**
  - `SUPABASE_URL`: https://kxxdxnsvbvdpyfxccvjb.supabase.co
  - `SUPABASE_KEY`: [API Key]

**Функции меню:**
- 🔗 Открыть Supabase Console
- 📊 Просмотр данных (таблицы)
- 🔍 Выполнить SQL запрос
- 📥 Импортировать данные
- 📤 Экспортировать данные
- 🔐 Управление правами доступа
- ⚙️ Настройки подключения

---

### 2. **GitHub MCP Server**
- **Тип:** Version Control & Issues
- **Статус:** ✅ Активен
- **Назначение:** Управление GitHub репозиториями, issues, PRs
- **Подключение:** Docker (ghcr.io/github/github-mcp-server)
- **Переменные:**
  - `GITHUB_PERSONAL_ACCESS_TOKEN`: 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

---

### 3. **Google Drive**
- **Тип:** Cloud Storage
- **Статус:** ✅ Активен
- **Назначение:** Управление файлами Google Drive
- **Подключение:** Docker с Google credentials
- **Файл credentials:** `/Users/aleksandr/Desktop/AgentCare/config/credentials.json`

---

### 4. **Google Sheets**
- **Тип:** Spreadsheet Management
- **Статус:** ✅ Активен
- **Назначение:** Чтение и запись в Google Sheets
- **Подключение:** Docker с Google credentials

---

### 5. **Google Apps Script**
- **Тип:** Cloud Function Management
- **Статус:** ✅ Активен
- **Назначение:** Деплой и управление Apps Script функциями
- **Подключение:** Docker с Google credentials

---

### 6. **Fetch Server**
- **Тип:** HTTP Client
- **Статус:** ✅ Активен
- **Назначение:** HTTP запросы к внешним API
- **Подключение:** Docker (modelcontextprotocol/server-fetch)

---

### 7. **Brave Search**
- **Тип:** Web Search
- **Статус:** ✅ Активен
- **Назначение:** Поиск в интернете
- **Подключение:** Docker
- **Переменные:**
  - `BRAVE_API_KEY`: BSA7I0LGblTYM_CqzP8j18EE7uQC1qm

---

### 8. **Redis**
- **Тип:** Cache Database
- **Статус:** ✅ Активен
- **Назначение:** Кэширование и сессии
- **Подключение:** npx @modelcontextprotocol/server-redis
- **URL:** redis://localhost:6379

---

### 9. **Filesystem**
- **Тип:** Local File System
- **Статус:** ✅ Активен
- **Назначение:** Чтение и запись в локальную файловую систему
- **Подключение:** npx @modelcontextprotocol/server-filesystem
- **Доступные пути:**
  - `/Users/aleksandr/Desktop/AgentCare`
  - `/Users/aleksandr/Desktop`

---

### 10. **Docker**
- **Тип:** Container Management
- **Статус:** ✅ Активен
- **Назначение:** Управление Docker контейнерами
- **Подключение:** npx @quantgeekdev/docker-mcp

---

### 11. **Context7** (Upstash)
- **Тип:** Context Management
- **Статус:** ✅ Активен
- **Назначение:** Управление контекстом между сессиями
- **Подключение:** npx @upstash/context7-mcp

---

## 🛠️ Configuration File

**Путь:** `~/Library/Application Support/Claude/claude_desktop_config.json`

Все MCP серверы определены в этом файле JSON. Структура:

```json
{
  "mcpServers": {
    "server-name": {
      "command": "npx|docker",
      "args": ["arg1", "arg2"],
      "env": {
        "ENV_VAR": "value"
      }
    }
  }
}
```

---

## 🔧 Как добавить новый MCP Server

1. **Найти MCP сервер** на https://github.com/modelcontextprotocol/servers

2. **Добавить в claude_desktop_config.json:**
   ```json
   "my-server": {
     "command": "npx",
     "args": ["-y", "@myorg/my-server"],
     "env": { "API_KEY": "value" }
   }
   ```

3. **Перезагрузить Claude Desktop** (полный перезапуск приложения)

4. **Проверить в Claude Code:**
   ```
   Доступные инструменты → должен быть виден новый сервер
   ```

---

## 🚀 Использование MCP Servers в AgentCare

### Пример: Работа с Supabase через AI

1. **Через меню Google Sheets:**
   - Откройте документ
   - Нажмите на меню "🗄️ Supabase"
   - Выберите нужную операцию

2. **Через Claude AI директно:**
   ```
   "Получи список всех пользователей из таблицы users в Supabase"
   ```

3. **Автоматизация:**
   - Используется MCP сервер для выполнения запроса
   - Результат возвращается в AI контекст
   - AI обрабатывает результат и представляет пользователю

---

## 📊 Статус Проверки

Все серверы можно проверить командой:

```bash
# Будет добавлена в дашборд
curl http://localhost:8000/api/v1/mcp/status
```

---

## ⚠️ Важные Замечания

1. **Безопасность:**
   - API ключи хранятся в claude_desktop_config.json
   - Никогда не коммитьте этот файл в git
   - Используйте переменные окружения для чувствительных данных

2. **Docker требования:**
   - Docker должен быть установлен и запущен
   - Образы скачиваются автоматически при первом использовании

3. **Google Credentials:**
   - Файл `/config/credentials.json` должен содержать валидные Google API credentials
   - Включить необходимые API в Google Cloud Console

---

## 🐛 Troubleshooting

### MCP сервер не работает

```bash
# 1. Проверить claude_desktop_config.json синтаксис
python3 -m json.tool ~/Library/Application\ Support/Claude/claude_desktop_config.json

# 2. Перезагрузить Claude Desktop (полный перезапуск)

# 3. Проверить логи в Claude Console (F1 → Developer → Console)
```

### Supabase подключение не работает

```bash
# Проверить переменные окружения
echo $SUPABASE_URL
echo $SUPABASE_KEY

# Тестировать подключение через функцию
testSupabaseConnection()
```

---

## 📚 Полезные Ссылки

- [MCP Protocol Spec](https://modelcontextprotocol.io)
- [Supabase Documentation](https://supabase.com/docs)
- [Claude API Documentation](https://claude.com/docs)
- [AgentCare Project Repository](https://github.com/your-repo/AgentCare)

---

**Последнее обновление:** 2026-01-14
**Автор:** Claude AI
**Статус:** ✅ Production Ready
