# 🚀 AGENTCARE SUPABASE INTEGRATION

**Версия**: 1.0
**Дата**: 2026-01-14
**Статус**: ✅ Готово к использованию

---

## 📖 Описание

Полная интеграция **Supabase** для AgentCare - облачная база данных PostgreSQL для хранения:
- 📦 Артикулов и товаров (MT, SS, SK проекты)
- 💰 Цен и остатков
- 📋 Логов и журналов операций
- ⚙️ Конфигурации и правил синхронизации

---

## 🎯 Преимущества Supabase

✅ **Масштабируемость** - легко растет с вашими данными
✅ **Real-time** - слушайте изменения в реальном времени
✅ **Безопасность** - Row Level Security (RLS) политики
✅ **Бесплатный tier** - начните без затрат
✅ **Open-source** - контролируйте свои данные
✅ **PostgreSQL** - мощная базовая база данных

---

## 📁 Структура папок

```
supabase/
├── migrations/
│   └── 001_initial_schema.sql       # Все таблицы и функции
├── SETUP_GUIDE.md                   # Пошаговое руководство по настройке
├── README.md                         # Этот файл
└── .env.example                      # Пример переменных окружения
```

---

## 🚀 Быстрый старт (5 минут)

### Шаг 1: Получите Supabase URL и ключи
1. Перейдите на https://app.supabase.com
2. Откройте свой проект
3. ⚙️ Project Settings → API
4. Скопируйте:
   - **Project URL**: `https://jfpnhkcqvriblsiqqjis.supabase.co`
   - **Anon Key**: (для Google Apps Script)
   - **Service Role Key**: (для Python backend)

### Шаг 2: Инициализируйте БД
1. В консоли Supabase: **SQL Editor**
2. Создайте **New Query**
3. Скопируйте содержимое `migrations/001_initial_schema.sql`
4. Нажмите **RUN** ✅

### Шаг 3: Создайте `.env.local` файл
```bash
cd /Users/aleksandr/Desktop/AgentCare

cat > .env.local << 'EOF'
# Supabase
SUPABASE_URL=https://jfpnhkcqvriblsiqqjis.supabase.co
SUPABASE_ANON_KEY=eyJ...
SUPABASE_SERVICE_ROLE_KEY=sb_secret_...

# Settings
SYNC_INTERVAL_SECONDS=300
LOG_RETENTION_DAYS=30
EOF
```

### Шаг 4: Протестируйте подключение
**Для Google Apps Script:**
```javascript
// В Google Sheets скрипте нажмите Ctrl+Enter
testSupabaseConnection()
```

**Для Python:**
```bash
python3 scripts/supabase_sync.py
```

---

## 📊 Архитектура БД

### Основные таблицы

#### 1️⃣ Articles (Артикулы)
```sql
-- Основные данные о товарах
CREATE TABLE articles (
  id UUID PRIMARY KEY,
  article_id VARCHAR UNIQUE,    -- Уникальный ID артикула
  project_key VARCHAR,          -- MT, SS, SK
  article_rus VARCHAR,          -- Название на русском
  code_base VARCHAR,            -- Базовый код
  status VARCHAR DEFAULT 'ACTIVE',
  created_at TIMESTAMP,
  updated_at TIMESTAMP,
  google_sheet_id VARCHAR       -- Для синхронизации с Sheets
);

-- Примеры запросов:
SELECT * FROM articles WHERE project_key = 'MT' AND status = 'ACTIVE';
SELECT COUNT(*) FROM articles WHERE project_key = 'SK';
```

#### 2️⃣ Prices (Цены)
```sql
CREATE TABLE prices (
  id UUID PRIMARY KEY,
  article_id UUID REFERENCES articles(id),
  project_key VARCHAR,
  price_base DECIMAL,
  price_sk DECIMAL,
  price_mt DECIMAL,
  date_from DATE,
  date_to DATE,
  created_at TIMESTAMP,
  updated_at TIMESTAMP
);
```

#### 3️⃣ Stocks (Остатки)
```sql
CREATE TABLE stocks (
  id UUID PRIMARY KEY,
  article_id UUID REFERENCES articles(id),
  project_key VARCHAR,
  quantity_total INT,
  quantity_available INT,
  supplier_name VARCHAR,
  warehouse_location VARCHAR,
  created_at TIMESTAMP,
  updated_at TIMESTAMP
);
```

#### 4️⃣ Logs (Журналы)
```sql
CREATE TABLE logs (
  id UUID PRIMARY KEY,
  timestamp TIMESTAMP,
  level VARCHAR,                -- DEBUG, INFO, WARN, ERROR
  category VARCHAR,             -- Article, Sync, Price, etc.
  message TEXT,
  function_name VARCHAR,
  variables JSONB,              -- Параметры операции
  result JSONB,                 -- Результаты операции
  created_at TIMESTAMP
);

-- Пример: Получить все ошибки за день
SELECT * FROM logs
WHERE level = 'ERROR'
  AND timestamp > NOW() - INTERVAL '1 day'
ORDER BY timestamp DESC;
```

#### 5️⃣ Sync Rules (Правила синхронизации)
```sql
CREATE TABLE sync_rules (
  id UUID PRIMARY KEY,
  name VARCHAR,
  source_type VARCHAR,          -- GOOGLE_SHEETS, SUPABASE
  target_type VARCHAR,
  sync_direction VARCHAR,       -- ONE_WAY, TWO_WAY
  schedule VARCHAR,             -- CRON или MANUAL
  enabled BOOLEAN,
  last_synced_at TIMESTAMP,
  next_sync_at TIMESTAMP
);
```

---

## 🔄 Синхронизация данных

### Способ 1: Python Backend (РЕКОМЕНДУЕТСЯ)

Синхронизируйте данные из Google Sheets в Supabase:

```bash
# 1. Активируйте виртуальное окружение
cd /Users/aleksandr/Desktop/AgentCare
source venv/bin/activate

# 2. Запустите скрипт синхронизации
python3 scripts/supabase_sync.py

# Результаты:
# ✅ MT: 450 артикулов синхронизировано
# ✅ SS: 320 артикулов синхронизировано
# ✅ SK: 280 артикулов синхронизировано
```

### Способ 2: Google Apps Script (Google Sheets)

```javascript
// В Google Sheets скрипте:

// 1. Вставить логи в Supabase
insertLogToSupabase({
  timestamp: new Date().toISOString(),
  level: 'INFO',
  category: 'Article',
  message: 'New article added',
  functionName: 'addArticleManually',
  variables: { articleId: 'ART_12345' },
  result: { success: true },
  projectKey: 'MT'
});

// 2. Получить артикулы из Supabase
const articles = getArticlesFromSupabase('MT');
articles.forEach(article => {
  Logger.log(`${article.article_id}: ${article.article_rus}`);
});

// 3. Вставить артикулы в Supabase
insertArticlesToSupabase(articles);

// 4. Обновить артикулы в Supabase
updateArticlesInSupabase(updatedArticles);
```

### Способ 3: API (Direct)

```bash
# Получить артикулы (curl)
curl -X GET "https://jfpnhkcqvriblsiqqjis.supabase.co/rest/v1/articles?project_key=eq.MT" \
  -H "apikey: YOUR_ANON_KEY" \
  -H "Authorization: Bearer YOUR_ANON_KEY"

# Вставить артикул
curl -X POST "https://jfpnhkcqvriblsiqqjis.supabase.co/rest/v1/articles" \
  -H "apikey: YOUR_ANON_KEY" \
  -H "Authorization: Bearer YOUR_ANON_KEY" \
  -H "Content-Type: application/json" \
  -d '{
    "article_id": "ART_12345",
    "project_key": "MT",
    "article_rus": "Мой товар",
    "status": "ACTIVE"
  }'
```

---

## 🔒 Безопасность

### Row Level Security (RLS)

Все таблицы защищены RLS политиками:

```sql
-- Пример: Чтение артикулов разрешено всем, но изменение только админам
CREATE POLICY "articles_select_all" ON articles
FOR SELECT USING (true);

CREATE POLICY "articles_update_admin" ON articles
FOR UPDATE USING (
  auth.jwt()->>'role'::text = 'authenticated'
  AND EXISTS (SELECT 1 FROM users WHERE id = auth.uid() AND 'admin' = ANY(roles))
);
```

### Переменные окружения

**НИКОГДА** не выкладывайте в коде:
```javascript
// ❌ ПЛОХО
const KEY = "sb_secret_LLXE5vAw5EM9tnLrm6SA6Q_3Cp7YySQ";

// ✅ ХОРОШО
const KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
```

### API Ключи

- **Anon Key**: Использование в клиентском коде (Google Apps Script) - ограниченный доступ
- **Service Role Key**: Только в backend (Python) - полный доступ, СЕКРЕТНО!

---

## 🧪 Тестирование

### Проверить подключение
```javascript
// Google Apps Script
testSupabaseConnection()
// Результат:
// ✅ Retrieved 15 categories
// ✅ Log inserted to Supabase
// ✅ Retrieved 5 logs
```

### Проверить данные в БД
```sql
-- SQL в Supabase Editor
SELECT COUNT(*) as total_articles FROM articles;
SELECT COUNT(*) as total_logs FROM logs;
SELECT * FROM log_categories LIMIT 5;
SELECT * FROM operation_statuses LIMIT 5;
```

### Проверить синхронизацию
```bash
# Python
python3 scripts/supabase_sync.py

# Результат в логе:
# 🔄 Starting synchronization (MT)
# ✅ Retrieved 450 articles from Google Sheets
# ✅ Synced 450 articles to Supabase: SUCCESS
# ✅ Synchronization completed
```

---

## 📈 Использование в Production

### 1. Автоматическая синхронизация

**Способ A: Cloud Tasks (Google Cloud)**
```yaml
# Запускает скрипт каждый час
schedule: "0 * * * *"
endpoint: https://your-domain.com/api/sync
```

**Способ B: Cron Job (Linux)**
```bash
# В crontab
0 * * * * cd /path/to/AgentCare && python3 scripts/supabase_sync.py
```

### 2. Резервные копии
```bash
# Резервная копия БД
pg_dump postgresql://user:password@jfpnhkcqvriblsiqqjis.supabase.co:5432/postgres \
  > backup_$(date +%Y%m%d).sql

# Восстановление
psql postgresql://user:password@jfpnhkcqvriblsiqqjis.supabase.co:5432/postgres < backup.sql
```

### 3. Мониторинг
```sql
-- Проверить размер таблиц
SELECT
  schemaname,
  tablename,
  pg_size_pretty(pg_total_relation_size(schemaname||'.'||tablename)) as size
FROM pg_tables
WHERE schemaname NOT IN ('pg_catalog', 'information_schema')
ORDER BY pg_total_relation_size(schemaname||'.'||tablename) DESC;

-- Проверить индексы
SELECT * FROM pg_stat_user_indexes ORDER BY idx_scan DESC;
```

---

## 📞 Часто задаваемые вопросы

### Q: Где взять Supabase URL и ключи?
**A**: В консоли Supabase → Project Settings → API (вкладка)

### Q: Можно ли использовать с Google Apps Script?
**A**: Да! Используйте `SupabaseClient.gs` с Anon Key (ограниченный доступ)

### Q: Что делать если видю ошибку 401?
**A**: Проверьте:
1. Правильный ли URL (без слешей в конце)?
2. Правильный ли API ключ?
3. Включены ли RLS политики?

### Q: Сколько стоит Supabase?
**A**: Бесплатный tier на 500MB. Pro план $25/месяц. [Подробнее](https://supabase.com/pricing)

### Q: Как удалить старые логи?
```sql
DELETE FROM logs WHERE created_at < NOW() - INTERVAL '30 days';
```

---

## 📚 Дополнительные ресурсы

- [Supabase Documentation](https://supabase.com/docs)
- [PostgreSQL Reference](https://www.postgresql.org/docs/)
- [Google Apps Script Reference](https://developers.google.com/apps-script)
- [REST API Reference](https://supabase.com/docs/guides/api)

---

## 📝 Версионирование

| Версия | Дата | Изменения |
|--------|------|-----------|
| 1.0 | 2026-01-14 | Начальная версия - все таблицы, функции, документация |

---

## 👨‍💻 Поддержка

Если возникли проблемы:
1. Проверьте логи: `supabase_sync.log`
2. Смотрите Database Logs в консоли Supabase
3. Используйте SQL Editor для тестирования запросов
4. Проверьте DevTools (F12) в браузере для сетевых ошибок

---

**Версия**: 1.0
**Последнее обновление**: 2026-01-14
**Статус**: ✅ Production Ready
