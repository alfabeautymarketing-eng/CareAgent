# 💡 ПРИМЕРЫ ИСПОЛЬЗОВАНИЯ SUPABASE

**Версия**: 1.0
**Дата**: 2026-01-14

---

## 📋 Содержание

1. [Google Apps Script примеры](#google-apps-script-примеры)
2. [Python примеры](#python-примеры)
3. [SQL примеры](#sql-примеры)
4. [API примеры (curl)](#api-примеры)

---

## 🔶 Google Apps Script примеры

### Пример 1: Вставить лог при выполнении функции

```javascript
function myCustomFunction() {
  try {
    // Начало функции
    insertLogToSupabase({
      timestamp: new Date().toISOString(),
      level: 'INFO',
      category: 'Article',
      message: 'Начало обработки артикулов',
      functionName: 'myCustomFunction',
      projectKey: 'MT'
    });

    // Ваш код здесь...
    const articles = getArticlesFromSupabase('MT');
    Logger.log(`Получено ${articles.length} артикулов`);

    // Успешное завершение
    insertLogToSupabase({
      timestamp: new Date().toISOString(),
      level: 'INFO',
      category: 'Article',
      message: 'Обработка артикулов завершена успешно',
      functionName: 'myCustomFunction',
      variables: { count: articles.length },
      result: { success: true },
      projectKey: 'MT'
    });

  } catch (error) {
    // Ошибка
    insertLogToSupabase({
      timestamp: new Date().toISOString(),
      level: 'ERROR',
      category: 'Article',
      message: 'Ошибка при обработке артикулов',
      functionName: 'myCustomFunction',
      details: error.message,
      projectKey: 'MT'
    });

    throw error;
  }
}
```

### Пример 2: Синхронизировать данные между Google Sheets и Supabase

```javascript
/**
 * Синхронизировать все артикулы текущего листа в Supabase
 */
function syncCurrentSheetToSupabase() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getDataRange();
  const values = range.getValues();

  if (values.length < 2) {
    Logger.log("❌ No data to sync");
    return;
  }

  const headers = values[0];
  const articles = [];

  // Начало синхронизации
  insertLogToSupabase({
    timestamp: new Date().toISOString(),
    level: 'INFO',
    category: 'Sync',
    message: `Начало синхронизации ${values.length - 1} артикулов`,
    functionName: 'syncCurrentSheetToSupabase',
    projectKey: 'MT'
  });

  // Преобразуйте строки в объекты
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const article = {};

    for (let j = 0; j < headers.length; j++) {
      article[headers[j]] = row[j];
    }

    articles.push(article);
  }

  // Вставьте в Supabase
  const count = insertArticlesToSupabase(articles);

  // Конец синхронизации
  insertLogToSupabase({
    timestamp: new Date().toISOString(),
    level: 'INFO',
    category: 'Sync',
    message: `Синхронизация завершена: ${count} артикулов загружено`,
    functionName: 'syncCurrentSheetToSupabase',
    variables: { totalRows: articles.length },
    result: { synced: count },
    projectKey: 'MT'
  });

  Logger.log(`✅ Synced ${count} articles`);
}
```

### Пример 3: Получить и отобразить логи

```javascript
/**
 * Получить последние логи и отобразить их
 */
function displayRecentLogs() {
  const client = new SupabaseClient();

  // Получить последние 20 логов
  const logs = client.select('logs', {
    select: '*',
    order: 'timestamp.desc',
    limit: 20
  });

  if (!logs) {
    Logger.log("❌ Failed to fetch logs");
    return;
  }

  // Создать новый лист для логов
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .insertSheet('Logs Report');

  // Заголовки
  const headers = ['Время', 'Уровень', 'Категория', 'Сообщение', 'Функция'];
  sheet.appendRow(headers);

  // Добавить логи
  for (const log of logs) {
    const row = [
      new Date(log.timestamp).toLocaleString('ru-RU'),
      log.level,
      log.category,
      log.message,
      log.function_name
    ];
    sheet.appendRow(row);
  }

  // Отформатировать заголовки
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('white');

  // Автоширина колонок
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  Logger.log(`✅ Created logs report with ${logs.length} rows`);
}
```

### Пример 4: Обновить артикул при изменении в Google Sheet

```javascript
/**
 * Триггер для синхронизации при изменении
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();

  // Убедитесь, что изменение произошло в листе "Главная"
  if (sheet.getName() !== "Главная") return;

  // Убедитесь, что это не заголовок
  if (row === 1) return;

  // Получите данные строки
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Преобразуйте в объект
  const article = {};
  for (let i = 0; i < headers.length; i++) {
    article[headers[i]] = rowData[i];
  }

  // Обновите в Supabase
  const client = new SupabaseClient();
  const result = client.update('articles', {
    article_rus: article['Арт. Рус'],
    article_producer: article['Арт. произв.'],
    status: article['Статус'] || 'ACTIVE',
    synced_from_sheet: new Date().toISOString()
  }, {
    article_id: article['ID']
  });

  if (result) {
    Logger.log(`✅ Updated article ${article['ID']} in Supabase`);

    // Логируйте изменение
    insertLogToSupabase({
      timestamp: new Date().toISOString(),
      level: 'INFO',
      category: 'Article',
      message: `Артикул обновлен: ${article['ID']}`,
      functionName: 'onEdit',
      variables: { articleId: article['ID'], rowNumber: row },
      projectKey: 'MT'
    });
  }
}
```

---

## 🐍 Python примеры

### Пример 1: Базовое подключение и запрос

```python
from supabase import create_client
import os
from dotenv import load_dotenv

# Загрузить переменные окружения
load_dotenv('.env.local')

# Создать клиент
url = os.getenv('SUPABASE_URL')
key = os.getenv('SUPABASE_SERVICE_ROLE_KEY')
supabase = create_client(url, key)

# Получить все артикулы проекта MT
response = supabase.table('articles') \
    .select('*') \
    .eq('project_key', 'MT') \
    .limit(100) \
    .execute()

print(f"✅ Retrieved {len(response.data)} articles")
for article in response.data[:5]:
    print(f"  - {article['article_id']}: {article['article_rus']}")
```

### Пример 2: Вставить новую запись

```python
from supabase import create_client
from datetime import datetime

supabase = create_client(url, key)

# Вставить новый артикул
new_article = {
    'article_id': 'ART_NEW_12345',
    'project_key': 'MT',
    'article_rus': 'Новый товар',
    'code_base': 'CB_12345',
    'category': 'Drinks',
    'status': 'ACTIVE',
    'synced_from_sheet': datetime.now().isoformat()
}

response = supabase.table('articles').insert(new_article).execute()

if response.data:
    print(f"✅ Inserted article: {response.data[0]['article_id']}")
else:
    print(f"❌ Failed to insert article")
```

### Пример 3: Обновить записи

```python
# Обновить статус артикула
response = supabase.table('articles') \
    .update({'status': 'ARCHIVED'}) \
    .eq('article_id', 'ART_NEW_12345') \
    .execute()

print(f"✅ Updated {len(response.data)} records")
```

### Пример 4: Удалить записи

```python
# Удалить артикул
response = supabase.table('articles') \
    .delete() \
    .eq('article_id', 'ART_NEW_12345') \
    .execute()

print(f"✅ Deleted {len(response.data)} records")
```

### Пример 5: Фильтрованные запросы

```python
# Получить дорогие товары (цена > 1000)
response = supabase.table('prices') \
    .select('*') \
    .gt('price_base', 1000) \
    .execute()

print(f"✅ Found {len(response.data)} expensive items")

# Получить товары в определённом диапазоне дат
response = supabase.table('prices') \
    .select('*') \
    .gte('date_from', '2026-01-01') \
    .lte('date_to', '2026-12-31') \
    .execute()

print(f"✅ Found {len(response.data)} items in date range")
```

### Пример 6: Работа с логами

```python
import json
from datetime import datetime

# Вставить лог
log_entry = {
    'timestamp': datetime.now().isoformat(),
    'level': 'INFO',
    'category': 'Sync',
    'message': 'Синхронизация завершена успешно',
    'function_name': 'sync_articles_from_sheets',
    'variables': json.dumps({'total_articles': 450, 'project': 'MT'}),
    'result': json.dumps({'status': 'SUCCESS', 'synced_count': 450}),
    'project_key': 'MT'
}

response = supabase.table('logs').insert(log_entry).execute()
print(f"✅ Log inserted: {response.data[0]['id']}")

# Получить логи за последний час
from datetime import timedelta
one_hour_ago = (datetime.now() - timedelta(hours=1)).isoformat()

response = supabase.table('logs') \
    .select('*') \
    .gte('timestamp', one_hour_ago) \
    .order('timestamp', desc=True) \
    .execute()

print(f"✅ Retrieved {len(response.data)} logs from last hour")
```

---

## 📊 SQL примеры

### Пример 1: Базовые SELECT запросы

```sql
-- Получить все активные артикулы проекта MT
SELECT * FROM articles
WHERE project_key = 'MT' AND status = 'ACTIVE';

-- Получить количество артикулов по проектам
SELECT project_key, COUNT(*) as count
FROM articles
GROUP BY project_key
ORDER BY count DESC;

-- Получить дорогие товары
SELECT a.article_id, a.article_rus, p.price_base
FROM articles a
LEFT JOIN prices p ON a.id = p.article_id
WHERE p.price_base > 1000
ORDER BY p.price_base DESC;
```

### Пример 2: Статистика логов

```sql
-- Получить статистику по уровням логирования
SELECT
  level,
  COUNT(*) as count,
  ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) as percentage
FROM logs
WHERE timestamp > NOW() - INTERVAL '1 day'
GROUP BY level
ORDER BY count DESC;

-- Получить самые активные функции
SELECT
  function_name,
  COUNT(*) as log_count,
  COUNT(CASE WHEN level = 'ERROR' THEN 1 END) as error_count
FROM logs
WHERE timestamp > NOW() - INTERVAL '7 days'
GROUP BY function_name
ORDER BY log_count DESC
LIMIT 10;
```

### Пример 3: Аналитика остатков

```sql
-- Получить товары с низкими остатками
SELECT
  a.article_id,
  a.article_rus,
  s.quantity_available,
  s.warehouse_location
FROM articles a
LEFT JOIN stocks s ON a.id = s.article_id
WHERE s.quantity_available < 10
ORDER BY s.quantity_available ASC;

-- Получить общий объем по складам
SELECT
  warehouse_location,
  SUM(quantity_total) as total_qty,
  SUM(quantity_available) as available_qty,
  COUNT(DISTINCT article_id) as unique_items
FROM stocks
GROUP BY warehouse_location
ORDER BY total_qty DESC;
```

### Пример 4: Удаление старых логов (обслуживание)

```sql
-- Удалить логи старше 30 дней
DELETE FROM logs
WHERE timestamp < NOW() - INTERVAL '30 days';

-- Проверить сколько удалим
SELECT COUNT(*) as old_logs_count
FROM logs
WHERE timestamp < NOW() - INTERVAL '30 days';

-- Архивировать логи перед удалением
INSERT INTO log_archives (archive_date, archive_type, logs_data, total_logs)
SELECT
  CURRENT_DATE - INTERVAL '30 days',
  'MONTHLY',
  JSON_AGG(ROW_TO_JSON(logs.*)),
  COUNT(*)
FROM logs
WHERE timestamp < NOW() - INTERVAL '30 days';
```

---

## 🌐 API примеры (curl)

### Пример 1: GET запрос (получить артикулы)

```bash
# Получить первые 10 артикулов проекта MT
curl -X GET \
  "https://jfpnhkcqvriblsiqqjis.supabase.co/rest/v1/articles?project_key=eq.MT&limit=10" \
  -H "apikey: YOUR_ANON_KEY" \
  -H "Authorization: Bearer YOUR_ANON_KEY"

# Результат (JSON):
# [
#   {
#     "id": "550e8400-e29b-41d4-a716-446655440000",
#     "article_id": "ART_12345",
#     "project_key": "MT",
#     "article_rus": "Мой товар",
#     ...
#   }
# ]
```

### Пример 2: POST запрос (вставить артикул)

```bash
curl -X POST \
  "https://jfpnhkcqvriblsiqqjis.supabase.co/rest/v1/articles" \
  -H "apikey: YOUR_ANON_KEY" \
  -H "Authorization: Bearer YOUR_ANON_KEY" \
  -H "Content-Type: application/json" \
  -d '{
    "article_id": "ART_NEW_123",
    "project_key": "MT",
    "article_rus": "Новый товар",
    "code_base": "CB_123",
    "category": "Drinks",
    "status": "ACTIVE"
  }'
```

### Пример 3: PATCH запрос (обновить артикул)

```bash
curl -X PATCH \
  "https://jfpnhkcqvriblsiqqjis.supabase.co/rest/v1/articles?article_id=eq.ART_NEW_123" \
  -H "apikey: YOUR_ANON_KEY" \
  -H "Authorization: Bearer YOUR_ANON_KEY" \
  -H "Content-Type: application/json" \
  -d '{
    "status": "ARCHIVED",
    "notes": "Снято с производства"
  }'
```

### Пример 4: DELETE запрос (удалить артикул)

```bash
curl -X DELETE \
  "https://jfpnhkcqvriblsiqqjis.supabase.co/rest/v1/articles?article_id=eq.ART_NEW_123" \
  -H "apikey: YOUR_ANON_KEY" \
  -H "Authorization: Bearer YOUR_ANON_KEY"
```

### Пример 5: Фильтрованный запрос с сортировкой

```bash
# Получить дорогие товары, отсортированные по цене
curl -X GET \
  "https://jfpnhkcqvriblsiqqjis.supabase.co/rest/v1/prices?price_base=gt.1000&order=price_base.desc&limit=20" \
  -H "apikey: YOUR_ANON_KEY" \
  -H "Authorization: Bearer YOUR_ANON_KEY"
```

---

## 🎯 Практические сценарии

### Сценарий 1: Ежедневная синхронизация с отчетом

```python
# daily_sync.py
from supabase_sync import DataSynchronizer
from datetime import datetime
import logging

logger = logging.getLogger(__name__)

def daily_sync_with_report():
    """Выполнить ежедневную синхронизацию и отправить отчет"""

    synchronizer = DataSynchronizer()
    projects = ['MT', 'SS', 'SK']

    report = {
        'date': datetime.now().isoformat(),
        'results': []
    }

    for project_key in projects:
        logger.info(f"🔄 Syncing {project_key}...")

        result = synchronizer.sync_articles_from_sheets(
            spreadsheet_id=SHEETS[project_key],
            sheet_range=f"Главная!A1:Z10000",
            project_key=project_key,
            upsert=True
        )

        report['results'].append({
            'project': project_key,
            'status': result.status.value,
            'total': result.total_records,
            'synced': result.processed_records,
            'errors': result.failed_records
        })

    # Отправить отчет (например, по email)
    send_sync_report(report)
    logger.info(f"✅ Daily sync completed")

if __name__ == '__main__':
    daily_sync_with_report()
```

### Сценарий 2: Мониторинг ошибок

```javascript
// Google Apps Script
function monitorErrors() {
  const client = new SupabaseClient();

  // Получить ошибки за последний час
  const oneHourAgo = new Date(Date.now() - 3600000).toISOString();

  const errors = client.select('logs', {
    select: '*',
    filter: `level=eq.ERROR,timestamp=gte.${oneHourAgo}`,
    order: 'timestamp.desc'
  });

  if (!errors || errors.length === 0) {
    Logger.log("✅ No errors in the last hour");
    return;
  }

  Logger.log(`⚠️ Found ${errors.length} errors:`);

  // Отправить уведомление (например, в Slack)
  sendSlackNotification({
    text: `⚠️ AgentCare: ${errors.length} ошибок за последний час`,
    errors: errors.map(e => ({
      message: e.message,
      function: e.function_name,
      time: new Date(e.timestamp).toLocaleString('ru-RU')
    }))
  });
}
```

---

**Версия**: 1.0
**Последнее обновление**: 2026-01-14
**Статус**: ✅ Ready to use
