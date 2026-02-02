# 🎯 SUPABASE ИНТЕГРАЦИЯ - ПОШАГОВОЕ РУКОВОДСТВО

**Версия**: 1.0
**Дата**: 2026-01-14
**Время установки**: ~15 минут

---

## 📋 Содержание

1. [Шаг 1: Получение учетных данных](#шаг-1-получение-учетных-данных) (2 минуты)
2. [Шаг 2: Инициализация БД](#шаг-2-инициализация-бд) (3 минуты)
3. [Шаг 3: Настройка окружения](#шаг-3-настройка-окружения) (2 минуты)
4. [Шаг 4: Python интеграция](#шаг-4-python-интеграция) (4 минуты)
5. [Шаг 5: Google Apps Script](#шаг-5-google-apps-script) (3 минуты)
6. [Шаг 6: Проверка](#шаг-6-проверка-работоспособности) (1 минута)

---

## 🔑 ШАГ 1: Получение учетных данных

### 1.1 Откройте консоль Supabase
```
https://app.supabase.com
```

### 1.2 Если у вас еще нет проекта:
1. Кликните **"New Project"**
2. Заполните:
   - **Name**: AgentCare
   - **Database Password**: (сгенерируется автоматически или введите свой)
   - **Region**: Select the closest region to you (например, `eu-west-1` для Европы)
3. Кликните **"Create new project"**
4. Дождитесь инициализации (2-3 минуты)

### 1.3 Получите учетные данные

**Путь**: В левой панели нажмите ⚙️ **Settings** → выберите **API**

Вы найдете:

```
Project URL:
  https://jfpnhkcqvriblsiqqjis.supabase.co

Anon Key (Публичный):
  eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJ...
  [полный ключ начинается с "eyJ"]

Service Role Key (Секретный):
  sb_secret_LLXE5vAw5EM9tnLrm6SA6Q_3Cp7YySQ...
  [полный ключ начинается с "sb_secret_"]
```

**⚠️ ВАЖНО**: Скопируйте **полные** ключи (не обрезанные)!

### 1.4 Сохраните в заметки (временно)
```
🔐 Ваши учетные данные (держите в безопасности!):

Project URL: https://jfpnhkcqvriblsiqqjis.supabase.co
Anon Key: sb_publishable_jfCetOVRKF9LuIs7uWPoaw__VSv-17O
[скопируйте полностью из консоли]
Service Role Key: [скопируйте полностью из консоли]
```

---

## 🗄️ ШАГ 2: Инициализация БД

### 2.1 Откройте SQL Editor
В консоли Supabase найдите левом меню:
- **SQL Editor** → **New Query**

Или откройте напрямую: https://app.supabase.com/project/jfpnhkcqvriblsiqqjis/sql/new

### 2.2 Скопируйте SQL скрипт

**Путь локально**: `supabase/migrations/001_initial_schema.sql`

1. Откройте этот файл в редакторе
2. **Выделите ВСЕ содержимое** (Ctrl+A)
3. **Скопируйте** (Ctrl+C)

### 2.3 Выполните скрипт

1. В SQL Editor вставьте код (Ctrl+V)
2. Нажмите кнопку **"RUN"** (или Ctrl+Enter)
3. Дождитесь появления ✅ **Query succeeded**

### 2.4 Проверьте таблицы

Нажмите в левом меню **Database** → **Tables**

Должны появиться таблицы:
- ✅ articles
- ✅ prices
- ✅ stocks
- ✅ logs
- ✅ log_archives
- ✅ sync_rules
- ✅ sync_history
- ✅ users
- ✅ audit_logs
- ✅ log_categories
- ✅ operation_statuses
- ✅ app_config
- ✅ sync_queue

---

## ⚙️ ШАГ 3: Настройка окружения

### 3.1 Создайте файл `.env.local`

```bash
# Перейдите в папку проекта
cd /Users/aleksandr/Desktop/AgentCare

# Создайте файл .env.local
cat > .env.local << 'EOF'
# ============================================================================
# SUPABASE CONFIG
# ============================================================================

# Скопируйте из https://app.supabase.com → Settings → API
SUPABASE_URL=https://jfpnhkcqvriblsiqqjis.supabase.co

# Anon Key (публичный, используется в Google Apps Script)
SUPABASE_ANON_KEY=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsImp...
# ☝️ Замените на ПОЛНЫЙ Anon Key из консоли

# Service Role Key (СЕКРЕТНЫЙ, только для Python backend!)
SUPABASE_SERVICE_ROLE_KEY=sb_secret_LLXE5vAw5EM9tnLrm6SA6Q_3Cp7YySQ...
# ☝️ Замените на ПОЛНЫЙ Service Role Key из консоли
# ⚠️ НИКОГДА не выкладывайте в публичные репозитории!

# ============================================================================
# GOOGLE SHEETS CONFIG (опционально)
# ============================================================================

# Если используете Google Sheets API
GOOGLE_SHEETS_API_KEY=AIzaSy...
GOOGLE_DRIVE_FOLDER_ID=1234567890...

# ============================================================================
# СИНХРОНИЗАЦИЯ
# ============================================================================

# Интервал синхронизации (в секундах)
SYNC_INTERVAL_SECONDS=300

# Сколько дней хранить логи
LOG_RETENTION_DAYS=30

# ============================================================================
# РЕЖИМ ОТЛАДКИ
# ============================================================================

# Уровень логирования: DEBUG, INFO, WARN, ERROR
LOG_LEVEL=INFO

# Разработчик?
DEBUG=false
EOF
```

### 3.2 Проверьте файл

```bash
# Просмотрите содержимое
cat .env.local

# Результат должен быть похож на:
# SUPABASE_URL=https://...
# SUPABASE_ANON_KEY=eyJ...
# SUPABASE_SERVICE_ROLE_KEY=sb_secret_...
```

### 3.3 Добавьте в `.gitignore`

```bash
# Убедитесь, что .env.local не выкладывается в git
echo ".env.local" >> .gitignore
```

---

## 🐍 ШАГ 4: Python интеграция

### 4.1 Установите зависимости

```bash
# Убедитесь, что находитесь в папке проекта
cd /Users/aleksandr/Desktop/AgentCare

# 1. Активируйте виртуальное окружение
source venv/bin/activate

# 2. Обновите pip
pip install --upgrade pip

# 3. Установите зависимости
pip install supabase python-dotenv google-auth google-auth-httplib2 google-auth-oauthlib
```

### 4.2 Проверьте импорты

```bash
python3 << 'EOF'
try:
    from supabase import create_client
    import dotenv
    print("✅ Суpabase SDK установлен корректно")
except Exception as e:
    print(f"❌ Ошибка: {e}")
EOF
```

### 4.3 Запустите синхронизацию

```bash
# Убедитесь, что .env.local содержит правильные ключи
python3 scripts/supabase_sync.py
```

**Ожидаемый результат:**
```
2026-01-14 10:00:00 - root - INFO - ✅ Supabase клиент инициализирован: https://jfpnhkcqvriblsiqqjis.supabase.co
2026-01-14 10:00:01 - root - INFO - 📦 Syncing project MT...
2026-01-14 10:00:05 - root - INFO - ✅ Retrieved 450 articles from Google Sheets
2026-01-14 10:00:10 - root - INFO - ✅ Synced 450 articles to Supabase: SUCCESS
2026-01-14 10:00:10 - root - INFO - ✅ SYNCHRONIZATION COMPLETED
```

### 4.4 Настройте автоматическую синхронизацию (опционально)

**Linux/Mac - Cron Job:**
```bash
# Отредактируйте cron
crontab -e

# Добавьте строку (синхронизация каждый час)
0 * * * * cd /Users/aleksandr/Desktop/AgentCare && source venv/bin/activate && python3 scripts/supabase_sync.py >> supabase_sync.log 2>&1

# Сохраните (Ctrl+X, Y, Enter)
```

---

## 📑 ШАГ 5: Google Apps Script

### 5.1 Откройте Google Sheet

1. Откройте одну из ваших Google Sheets (например, для проекта MT)
2. **Инструменты** → **Редактор скриптов** (Apps Script Editor)

### 5.2 Создайте новый файл

1. В редакторе слева кликните **"+"** рядом с файлами
2. Выберите **"Script"**
3. Назовите его **"SupabaseClient"**

### 5.3 Скопируйте код

**Локально**: `gas/SupabaseClient.gs`

1. Откройте этот файл
2. Скопируйте ВСЕ содержимое
3. Вставьте в Apps Script Editor

### 5.4 Обновите конфигурацию

В файле найдите строку:
```javascript
const SUPABASE_CONFIG = {
  URL: "https://jfpnhkcqvriblsiqqjis.supabase.co",
  ANON_KEY: "YOUR_ANON_KEY_HERE",  // ← Замените!
};
```

Замените `YOUR_ANON_KEY_HERE` на ваш **Anon Key** из Supabase консоли.

### 5.5 Сохраните

```
Ctrl+S  или  File → Save
```

### 5.6 Протестируйте подключение

1. В редакторе найдите функцию `testSupabaseConnection`
2. Нажмите **Run** (или выберите функцию из dropdown и нажмите ▶️)
3. Проверьте логи (View → Logs)

**Ожидаемый результат:**
```
✅ Supabase connected!
✅ Test log inserted to Supabase
✅ Retrieved 5 logs
✅ All tests completed!
```

### 5.7 Добавьте меню в Google Sheet

Теперь в Google Sheet появилось новое меню **⚡ Supabase** с кнопками:
- Test Connection
- Get Articles
- Insert Test Log
- Sync All Articles

---

## ✅ ШАГ 6: Проверка работоспособности

### 6.1 Проверьте данные в БД

**В консоли Supabase** (SQL Editor):
```sql
-- Проверьте количество артикулов
SELECT COUNT(*) as count FROM articles;

-- Проверьте логи
SELECT * FROM logs ORDER BY timestamp DESC LIMIT 10;

-- Проверьте справочники
SELECT * FROM log_categories;
SELECT * FROM operation_statuses;
```

### 6.2 Проверьте синхронизацию

**В Python:**
```bash
python3 scripts/supabase_sync.py
```

Должны синхронизироваться артикулы из трех проектов (MT, SS, SK).

### 6.3 Проверьте Google Apps Script

**В Google Sheet:**
1. Нажмите **⚡ Supabase** → **Test Connection**
2. Откройте Logs (Ctrl+Enter в редакторе)
3. Должны увидеть ✅ сообщения об успехе

### 6.4 Проверьте двустороннюю синхронизацию

**Тестовый сценарий:**
1. Добавьте новый артикул в Google Sheet
2. Запустите `python3 scripts/supabase_sync.py`
3. Проверьте в Supabase консоли - артикул должен появиться
4. Добавьте лог через GAS - он должен появиться в Supabase

---

## 🎉 Готово!

Вы успешно настроили Supabase интеграцию! Теперь вы можете:

✅ **Синхронизировать данные** между Google Sheets и Supabase
✅ **Хранить логи** в облаке
✅ **Масштабировать** базу данных без ограничений Google Sheets
✅ **Анализировать** данные с помощью SQL запросов
✅ **Автоматизировать** синхронизацию через Cron

---

## 🆘 Решение проблем

### Ошибка: "401 Unauthorized"
```
❌ Error: 401 - Unauthorized
```
**Решение**: Проверьте Anon Key - скопируйте полный ключ (не обрезанный)

### Ошибка: "Connection refused"
```
❌ Connection refused to https://...
```
**Решение**: Убедитесь, что Supabase project инициализирован

### Ошибка: "SUPABASE_URL не установлен"
```
ValueError: SUPABASE_URL и SUPABASE_SERVICE_ROLE_KEY должны быть установлены в .env.local
```
**Решение**: Проверьте файл `.env.local` - должны быть установлены все переменные

### Google Apps Script выдает ошибку
```
❌ Error: PERMISSION_DENIED
```
**Решение**:
1. Убедитесь, что URL правильный (без слешей в конце)
2. Проверьте ANON_KEY - должен быть полный ключ
3. Включите RLS политики в Supabase (см. SETUP_GUIDE.md)

---

## 📚 Дополнительные ресурсы

- 📖 [Полная документация](supabase/README.md)
- 🔧 [Гайд по установке](supabase/SETUP_GUIDE.md)
- 🐍 [Python скрипт синхронизации](scripts/supabase_sync.py)
- 📑 [Google Apps Script клиент](gas/SupabaseClient.gs)
- 🗄️ [SQL схема](supabase/migrations/001_initial_schema.sql)

---

## 📞 Контакты и поддержка

Если возникли проблемы:
1. Проверьте логи: `tail -f supabase_sync.log`
2. Смотрите Supabase Dashboard → Database → Logs
3. Используйте DevTools (F12) в браузере
4. Проверьте Database Editor в консоли Supabase

---

**Версия**: 1.0
**Дата**: 2026-01-14
**Статус**: ✅ Production Ready
