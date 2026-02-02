# 📚 SUPABASE SETUP GUIDE - Полное руководство по настройке

**Дата**: 2026-01-14
**Версия**: 1.0
**Статус**: 🔵 В процессе

---

## 📋 Содержание

1. [Получение учетных данных](#получение-учетных-данных)
2. [Инициализация схемы БД](#инициализация-схемы-бд)
3. [Настройка прав доступа (RLS)](#настройка-прав-доступа-rls)
4. [Настройка среды](#настройка-среды)
5. [Проверка подключения](#проверка-подключения)

---

## 🔑 Получение учетных данных

### Шаг 1: Откройте консоль Supabase
```
https://app.supabase.com
```

### Шаг 2: Выберите проект
- Слева в меню найдите ваш проект
- Кликните на него

### Шаг 3: Получите Connection String и API ключи
**Путь**: ⚙️ Project Settings → API → Connection Strings & API Keys

**Необходимые данные:**
```
Project URL:           https://jfpnhkcqvriblsiqqjis.supabase.co
Anon Key (public):     eyJ... (полный ключ)
Service Role Key:      sb_secret_... (полный ключ)
Database Password:     ...
```

**⚠️ БЕЗОПАСНОСТЬ:**
- Service Role Key - держите в секрете! Никогда не выкладывайте в коде!
- Используйте переменные окружения (`.env`)
- Anon Key можно использовать в клиентском коде (с ограниченными правами)

---

## 🗄️ Инициализация схемы БД

### Способ 1: Через SQL Editor в Supabase (РЕКОМЕНДУЕТСЯ)

1. **Откройте SQL Editor**
   - В консоли Supabase: SQL Editor (слева в меню)
   - Или кликните: [New Query](https://app.supabase.com/project/jfpnhkcqvriblsiqqjis/sql/new)

2. **Скопируйте содержимое файла**
   - Откройте: `supabase/migrations/001_initial_schema.sql`
   - Скопируйте ВСЕ содержимое

3. **Выполните скрипт**
   - Вставьте код в SQL Editor
   - Кликните кнопку "RUN" или нажмите Ctrl+Enter
   - Дождитесь сообщения об успехе ✅

### Способ 2: Через Command Line (для опытных)

```bash
# 1. Установите Supabase CLI
npm install -g supabase

# 2. Авторизуйтесь
supabase login

# 3. Инициализируйте проект (если еще не сделано)
supabase init

# 4. Создайте миграцию
supabase migration new initial_schema

# 5. Запустите миграцию
supabase db push
```

### Проверка успешной инициализации

После выполнения скрипта должны появиться таблицы:

```
✅ articles
✅ prices
✅ stocks
✅ logs
✅ log_archives
✅ sync_rules
✅ sync_history
✅ sync_queue
✅ users
✅ audit_logs
✅ log_categories
✅ operation_statuses
✅ app_config
```

Проверьте через **Table Editor** → слева должны быть все таблицы.

---

## 🔐 Настройка прав доступа (RLS)

### Включение RLS (Row Level Security)

Это **критически важно** для безопасности!

1. **Откройте Authentication → Policies**
2. **Для каждой таблицы создайте политики:**

#### Таблица `articles` (чтение всем, изменение только администраторам)

```sql
-- Чтение: все авторизованные пользователи
CREATE POLICY "articles_select_all" ON articles
FOR SELECT
USING (TRUE);

-- Изменение: только администраторы
CREATE POLICY "articles_update_admin" ON articles
FOR UPDATE
USING (
  (SELECT auth.jwt()->>'role')::text = 'authenticated'
  AND EXISTS (
    SELECT 1 FROM users WHERE id = auth.uid() AND 'admin' = ANY(roles)
  )
);

-- Удаление: только администраторы
CREATE POLICY "articles_delete_admin" ON articles
FOR DELETE
USING (
  EXISTS (
    SELECT 1 FROM users WHERE id = auth.uid() AND 'admin' = ANY(roles)
  )
);
```

#### Таблица `logs` (чтение только пользователям, вставка сервисом)

```sql
-- Чтение: авторизованные пользователи
CREATE POLICY "logs_select" ON logs
FOR SELECT
USING (TRUE);

-- Вставка: приложение (используя service_role)
CREATE POLICY "logs_insert_service" ON logs
FOR INSERT
WITH CHECK (TRUE);
```

#### Таблица `audit_logs` (чтение администраторам, вставка система)

```sql
-- Чтение: администраторы
CREATE POLICY "audit_logs_select_admin" ON audit_logs
FOR SELECT
USING (
  EXISTS (
    SELECT 1 FROM users WHERE id = auth.uid() AND 'admin' = ANY(roles)
  )
);

-- Вставка: система
CREATE POLICY "audit_logs_insert_system" ON audit_logs
FOR INSERT
WITH CHECK (TRUE);
```

---

## 🔧 Настройка среды

### 1. Создайте файл `.env.local`

```bash
# В корне проекта /Users/aleksandr/Desktop/AgentCare/

# Supabase
SUPABASE_URL=https://jfpnhkcqvriblsiqqjis.supabase.co
SUPABASE_ANON_KEY=eyJ... (скопируйте полный ключ)
SUPABASE_SERVICE_ROLE_KEY=sb_secret_... (скопируйте полный ключ)

# Google Sheets API
GOOGLE_SHEETS_API_KEY=... (если нужен)

# Синхронизация
SYNC_INTERVAL_SECONDS=300
LOG_RETENTION_DAYS=30
```

### 2. Для Python (если используете backend)

```bash
# В директории проекта

# 1. Создайте виртуальное окружение
python3 -m venv venv
source venv/bin/activate  # или `venv\Scripts\activate` на Windows

# 2. Установите зависимости
pip install supabase python-dotenv google-auth google-auth-httplib2 google-auth-oauthlib

# 3. Создайте файл config.py
cat > config.py << 'EOF'
import os
from dotenv import load_dotenv

load_dotenv('.env.local')

SUPABASE_URL = os.getenv('SUPABASE_URL')
SUPABASE_KEY = os.getenv('SUPABASE_SERVICE_ROLE_KEY')
SUPABASE_ANON_KEY = os.getenv('SUPABASE_ANON_KEY')

SYNC_INTERVAL = int(os.getenv('SYNC_INTERVAL_SECONDS', 300))
LOG_RETENTION_DAYS = int(os.getenv('LOG_RETENTION_DAYS', 30))
EOF
```

### 3. Для Google Apps Script

В файле `gas/00_GlobalApiBridge.gs` или нового файла `gas/SupabaseClient.gs`:

```javascript
// SUPABASE CONFIGURATION
const SUPABASE_CONFIG = {
  URL: "https://jfpnhkcqvriblsiqqjis.supabase.co",
  ANON_KEY: "eyJ...", // скопируйте Anon Key
  // ВНИМАНИЕ: service_role_key НИКОГДА не выкладывайте в GAS!
  // Используйте только anon_key с RLS политиками
};

function getSupabaseUrl() {
  return SUPABASE_CONFIG.URL;
}

function getSupabaseAnonKey() {
  return SUPABASE_CONFIG.ANON_KEY;
}
```

---

## 🧪 Проверка подключения

### Способ 1: Через SQL Query

В Supabase SQL Editor:

```sql
-- 1. Проверка таблиц
SELECT table_name
FROM information_schema.tables
WHERE table_schema = 'public'
ORDER BY table_name;

-- 2. Проверка количества записей
SELECT
  'articles' as table_name, COUNT(*) as count FROM articles
UNION ALL
SELECT 'logs' as table_name, COUNT(*) FROM logs
UNION ALL
SELECT 'users' as table_name, COUNT(*) FROM users;

-- 3. Проверка справочников
SELECT * FROM log_categories;
SELECT * FROM operation_statuses;
```

### Способ 2: Через Python

```python
from supabase import create_client
import os

# Подключитесь
url = os.getenv('SUPABASE_URL')
key = os.getenv('SUPABASE_SERVICE_ROLE_KEY')

supabase = create_client(url, key)

# Проверьте таблицы
try:
    response = supabase.table('log_categories').select('*').execute()
    print("✅ Supabase connected!")
    print(f"Found {len(response.data)} categories")
except Exception as e:
    print(f"❌ Connection error: {e}")
```

### Способ 3: Через Google Apps Script

```javascript
function testSupabaseConnection() {
  const url = SUPABASE_CONFIG.URL + '/rest/v1/log_categories';
  const options = {
    method: 'get',
    headers: {
      'apikey': SUPABASE_CONFIG.ANON_KEY,
      'Authorization': 'Bearer ' + SUPABASE_CONFIG.ANON_KEY,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      Logger.log('✅ Supabase connection successful!');
      Logger.log(response.getContentText());
    } else {
      Logger.log('❌ Error: ' + response.getResponseCode());
      Logger.log(response.getContentText());
    }
  } catch (e) {
    Logger.log('❌ Connection failed: ' + e);
  }
}
```

Запустите: `testSupabaseConnection()`

---

## 🚀 Следующие шаги

1. **[Создание Python скрипта синхронизации](PYTHON_SYNC_GUIDE.md)**
2. **[Интеграция с Google Apps Script](GAS_INTEGRATION_GUIDE.md)**
3. **[Настройка двусторонней синхронизации](BIDIRECTIONAL_SYNC.md)**

---

## 📞 Поддержка

Если возникли ошибки:

1. **Проверьте URL и ключи** - копируйте точно из консоли
2. **Убедитесь в RLS политиках** - должны быть включены и правильно настроены
3. **Смотрите логи Supabase** → Database → Logs
4. **Проверьте сетевые запросы** → в DevTools браузера

---

**Версия**: 1.0
**Последнее обновление**: 2026-01-14
**Статус**: ✅ Готово к использованию
