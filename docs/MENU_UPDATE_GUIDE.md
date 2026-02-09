# 📋 Руководство обновления меню

## ✅ Статус: Готово к использованию

Сервер развёрнут и работает на `http://46.226.167.153:8000`

### Проверка работы сервера

```bash
curl http://46.226.167.153:8000/health
```

Ответ должен быть:
```json
{
  "status": "healthy",
  "version": "0.1.0",
  "checks": {
    "redis": "ok",
    "google_sheets": "ok",
    "gemini": "ok"
  }
}
```

---

## 🔄 Обновление меню в Google Sheets

### Для КАЖДОГО проекта (MT, SK, SS):

**Шаг 1️⃣: Откройте Google Sheets**
- Откройте вашу таблицу (Montibello, Carmado или San)

**Шаг 2️⃣: Откройте Apps Script**
- Нажмите: **Extensions → Apps Script**
- Откроется редактор в новой вкладке

**Шаг 3️⃣: Очистите кэш и обновите меню**
- В редакторе нажмите **Ctrl+Enter** (открывается Console)
- В Console выполните (скопируйте и вставьте):

```javascript
clearMenuCache();
refreshMenu();
```

- Нажмите **Enter**

**Результат:**
Вы должны увидеть два уведомления в таблице:
1. "Кэш меню очищен!"
2. "Меню обновлено!"

**Шаг 4️⃣: Проверьте меню**
- Вернитесь на вкладку с Google Sheets
- Обновите страницу (Ctrl+R)
- Откройте меню - должны видеть **новую структуру** ✨

---

## 📊 Новая структура меню

```
🏷️ Прайс-лист
🛒 Заказ & Документация
📦 Экспорт & Выгрузка
🤖 AI Агент
⚙️ Настройки
   ├─ 🔄 Синхронизация
   ├─ 📋 Логи
   ├─ 🗄️ Supabase
   └─ 🟢 Ecosystem
```

---

## 🔧 Если меню всё ещё не обновилось

**Вариант 1: Полная очистка**

```javascript
PropertiesService.getUserProperties().deleteAllProperties();
refreshMenu();
```

**Вариант 2: Hard refresh Google Sheets**
- Нажмите **Ctrl+Shift+R** (полная перезагрузка кэша браузера)
- Откройте таблицу заново

**Вариант 3: Проверьте Console**

```javascript
// Проверить текущий кэш
const props = PropertiesService.getUserProperties();
const keys = props.getKeys();
Logger.log('Cached properties:', keys);

// Проверить SERVER_URL
Logger.log('Server URL:', SERVER_URL);
```

---

## 📱 API Endpoint (для разработчиков)

```
GET /api/v1/menu/config?spreadsheet_id=SPREADSHEET_ID
```

**Пример:**
```bash
curl "http://46.226.167.153:8000/api/v1/menu/config?spreadsheet_id=13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ"
```

---

## 📝 ID проектов для использования

```
MT (CosmeticaBar):
  - 13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ (Main)
  - 1fMOjUE7oZV96fCY5j5rPxnhWGJkDqg-GfwPZ8jUVgPw (Alt)

SK (Carmado):
  - 1CpYYLvRYslsyCkuLzL9EbbjsvbNpWCEZcmhKqMoX5zw (Main)
  - 1hSsS9_Iu_MgKWsoE19hAMouQInLGVFaBF6ZFG4Bsm1s (Production)

SS (San):
  - 12yIL1CuESZxeUUd-oKK2brtN1FnXE9q95N7SqzNc7vk (Main)
  - 1Bq2Pq0P1SQZfJNBZC3yduYCJmnyc4L4vmbLtvsVUkcg (Production)
```

---

## 🎯 Принципы новой структуры

✅ **Одна кнопка = одна функция** (без диалогов выбора)
✅ **По жизненному циклу пользователя** (Прайс → Заказ → Документация → Выгрузка)
✅ **Управление спрятано в подменю** (⚙️ Настройки)
✅ **Чистое интуитивное главное меню** (5 основных пунктов)

---

## 🚀 Документация

- [MENU_STRUCTURE.md](docs/MENU_STRUCTURE.md) - Полная документация структуры меню
- [src/api/endpoints.py:2613-3038](src/api/endpoints.py#L2613-L3038) - Исходный код меню
- [gas/Menu.gs](gas/Menu.gs) - Google Apps Script реализация

---

## 📞 Поддержка

Если возникают проблемы:

1. **Проверьте лог сервера:**
   ```bash
   docker logs agentcare -f
   ```

2. **Проверьте Console в Apps Script (Ctrl+Enter):**
   ```javascript
   console.log('Current config:', getMenuConfig());
   ```

3. **Очистите всё и начните заново:**
   ```javascript
   clearMenuCache();
   refreshMenu();
   ```

---

**Статус:** ✅ Готово к использованию
**Сервер:** http://46.226.167.153:8000
**Документация:** docs/MENU_STRUCTURE.md
