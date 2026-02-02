# 🕐 АВТОМАТИЧЕСКАЯ СИНХРОНИЗАЦИЯ - CRON JOB

**Версия**: 1.0
**Дата**: 2026-01-14
**Статус**: ✅ Активна

---

## ✅ Статус Cron Job

### Установленная конфигурация

```bash
0 * * * * cd /Users/aleksandr/Desktop/AgentCare && source venv/bin/activate && python3 scripts/supabase_sync.py >> supabase_sync_cron.log 2>&1
```

**Означает:**
- ✅ **Запускается**: каждый час в 00:00 (полный час)
- ✅ **Рабочая папка**: `/Users/aleksandr/Desktop/AgentCare`
- ✅ **Окружение**: Активирует Python venv перед запуском
- ✅ **Логи**: Сохраняются в `supabase_sync_cron.log` с временными метками
- ✅ **Ошибки**: Перенаправляются в тот же лог файл

---

## 📊 Мониторинг синхронизации

### Просмотр логов в реальном времени

```bash
# Последние 20 строк логов
tail -n 20 /Users/aleksandr/Desktop/AgentCare/supabase_sync_cron.log

# Следить за логами в реальном времени (обновляется каждую секунду)
tail -f /Users/aleksandr/Desktop/AgentCare/supabase_sync_cron.log

# Выход из режима наблюдения: Ctrl+C
```

### Проверка последней синхронизации

```bash
# Последние 5 успешных синхронизаций
tail -n 50 /Users/aleksandr/Desktop/AgentCare/supabase_sync_cron.log | grep "СИНХРОНИЗАЦИЯ ЗАВЕРШЕНА"

# Количество статей в последней синхронизации
tail -n 50 /Users/aleksandr/Desktop/AgentCare/supabase_sync_cron.log | grep "Всего:"

# Поиск ошибок
grep -i "ошиб\|error\|failed" /Users/aleksandr/Desktop/AgentCare/supabase_sync_cron.log | tail -n 10
```

### Проверка в Supabase

```sql
-- SQL запросы в Supabase SQL Editor

-- 1. Проверить время последней синхронизации
SELECT
  sync_rule_id,
  status,
  started_at,
  completed_at,
  total_records,
  processed_records,
  failed_records
FROM sync_history
ORDER BY started_at DESC
LIMIT 5;

-- 2. Проверить общее количество артикулов
SELECT
  project_key,
  COUNT(*) as total_articles,
  MAX(updated_at) as last_update
FROM articles
GROUP BY project_key;

-- 3. Проверить последние логи синхронизации
SELECT
  timestamp,
  level,
  category,
  message,
  project_key
FROM logs
WHERE category = 'Sync'
ORDER BY timestamp DESC
LIMIT 10;
```

---

## 🔧 Управление Cron Job

### Просмотр всех cron jobs

```bash
crontab -l
```

### Редактирование расписания

Если нужно изменить время запуска (например, каждые 30 минут):

```bash
# Открыть редактор crontab
crontab -e

# Измените строку на:
# Каждые 30 минут:
*/30 * * * * cd /Users/aleksandr/Desktop/AgentCare && source venv/bin/activate && python3 scripts/supabase_sync.py >> supabase_sync_cron.log 2>&1

# Каждые 15 минут:
*/15 * * * * cd /Users/aleksandr/Desktop/AgentCare && source venv/bin/activate && python3 scripts/supabase_sync.py >> supabase_sync_cron.log 2>&1

# Каждые 5 минут:
*/5 * * * * cd /Users/aleksandr/Desktop/AgentCare && source venv/bin/activate && python3 scripts/supabase_sync.py >> supabase_sync_cron.log 2>&1

# Сохраните: Ctrl+X, Y, Enter (если используете nano)
```

### Удаление cron job (если нужно отключить)

```bash
# Полностью удалить все cron jobs
crontab -r

# Или открыть редактор и удалить вручную строку синхронизации
crontab -e
```

---

## 📈 Рекомендуемые расписания

| Интервал | Назначение | Использование |
|----------|-----------|----------------|
| **Каждый час** (0 часов) | Стандартная синхронизация | 👈 Текущая конфигурация |
| **Каждые 30 минут** | Более частые обновления | Для активных проектов |
| **Каждые 15 минут** | Почти реал-тайм | Для критичных данных |
| **Каждые 5 минут** | Максимальная актуальность | Не рекомендуется (нагрузка) |
| **2 раза в день** (00:00, 12:00) | Минимальная нагрузка | Для стабильных данных |

---

## 🚨 Решение проблем

### Проблема: Cron job не запускается

**Признак**: Логи не обновляются по расписанию

**Решение:**
```bash
# 1. Проверить, что venv существует
ls -la /Users/aleksandr/Desktop/AgentCare/venv/bin/activate

# 2. Проверить, что .env.local существует
ls -la /Users/aleksandr/Desktop/AgentCare/.env.local

# 3. Проверить, что скрипт существует
ls -la /Users/aleksandr/Desktop/AgentCare/scripts/supabase_sync.py

# 4. Проверить логи системы (Mac)
log stream --predicate 'process == "cron"' --level debug
```

### Проблема: "Permission denied" ошибка

**Решение:**
```bash
# Убедитесь, что скрипт исполняемый
chmod +x /Users/aleksandr/Desktop/AgentCare/scripts/supabase_sync.py

# Проверьте права на папку
chmod 755 /Users/aleksandr/Desktop/AgentCare
```

### Проблема: "SUPABASE_URL не установлен"

**Решение:**
```bash
# Убедитесь, что .env.local содержит все переменные
cat /Users/aleksandr/Desktop/AgentCare/.env.local

# Должно быть:
# SUPABASE_URL=https://...
# SUPABASE_SERVICE_ROLE_KEY=sb_secret_...
# SUPABASE_ANON_KEY=sb_publishable_...
```

---

## 📝 Примеры использования

### Запустить синхронизацию вручную

```bash
cd /Users/aleksandr/Desktop/AgentCare
source venv/bin/activate
python3 scripts/supabase_sync.py
```

### Проверить, работает ли синхронизация

```bash
# Запустить и вывести логи
python3 scripts/supabase_sync.py

# Ожидаемый результат:
# ✅ Supabase клиент инициализирован: https://jfpnhkcqvriblsiqqjis.supabase.co
# 📦 Синхронизация проекта MT...
# ✅ СИНХРОНИЗАЦИЯ ЗАВЕРШЕНА
```

### Экспортировать логи за неделю

```bash
# Создать резервную копию логов
cp /Users/aleksandr/Desktop/AgentCare/supabase_sync_cron.log \
   /Users/aleksandr/Desktop/AgentCare/backups/sync_logs_$(date +%Y-%m-%d).log
```

---

## 📊 Статистика синхронизации

### Проверить загрузку на базу данных

```bash
# Количество синхронизаций за сегодня
grep "$(date +%Y-%m-%d)" /Users/aleksandr/Desktop/AgentCare/supabase_sync_cron.log | wc -l

# Среднее время синхронизации (в секундах)
grep "СИНХРОНИЗАЦИЯ ЗАВЕРШЕНА" /Users/aleksandr/Desktop/AgentCare/supabase_sync_cron.log | tail -n 10
```

---

## 🎯 Что дальше?

После установки автоматической синхронизации:

1. **Мониторинг** (опционально):
   - Настройте email уведомления при ошибках
   - Создайте дашборд для отслеживания синхронизации

2. **Google Apps Script**:
   - Интегрируйте в Google Sheets для реал-тайм обновлений
   - Добавьте кнопки синхронизации в меню

3. **Row Level Security (RLS)**:
   - Настройте политики доступа в Supabase
   - Ограничьте доступ по пользователям и проектам

4. **Мониторинг и алерты**:
   - Настройте логирование ошибок
   - Создайте алерты для критичных сбоев

---

## ✅ Чек-лист

- [x] Cron job установлен и активен
- [ ] Проверить логи за последний час
- [ ] Убедиться, что синхронизация работает по расписанию
- [ ] Настроить мониторинг (опционально)
- [ ] Интегрировать Google Apps Script (опционально)

---

**Версия**: 1.0
**Дата создания**: 2026-01-14
**Последнее обновление**: 2026-01-14
**Статус**: ✅ Production Ready
