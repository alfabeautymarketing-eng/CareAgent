# AgentCare - Быстрая инструкция миграции на новый MacBook

**Время выполнения:** 20-30 минут
**Сложность:** Минимальная (просто запустить один скрипт)

---

## ✅ Что вам нужно ДО миграции

### На текущем MacBook (прямо сейчас)

1. **Подготовить credentials:**
   ```bash
   # Просто убедитесь, что есть:
   ls ~/Desktop/AgentCare/config/credentials.json
   ls ~/Desktop/AgentCare/.env
   ```

2. **Сохранить в безопасное место (выбрать один вариант):**

   **Вариант A: Google Drive** ⭐ Рекомендуется
   ```
   1. Открыть Google Drive
   2. Создать папку "AgentCare_Migration"
   3. Загрузить туда:
      - config/credentials.json
      - .env
   4. Скопировать ссылку для скачивания
   ```

   **Вариант B: AirDrop** (если новый MacBook рядом)
   ```bash
   # Просто откройте папку AgentCare и используйте AirDrop
   ```

   **Вариант C: Encrypted USB**
   ```bash
   cp config/credentials.json /Volumes/USB_NAME/
   cp .env /Volumes/USB_NAME/
   ```

---

## 🚀 На новом MacBook - 3 простых шага

### Шаг 1️⃣: Клонировать проект
```bash
git clone <ваш-git-repo> ~/Desktop/AgentCare
cd ~/Desktop/AgentCare
```

### Шаг 2️⃣: Поместить credentials
```bash
# Если используете Google Drive:
# 1. Скачать credentials.json
# 2. Поместить в: ~/Desktop/AgentCare/config/credentials.json

# Если используете AirDrop:
# 1. Принять файлы
# 2. Поместить в: ~/Desktop/AgentCare/config/ и ~/Desktop/AgentCare/

# Если используете USB:
# 1. Подключить USB
# 2. Скопировать файлы
cp /Volumes/USB_NAME/credentials.json ~/Desktop/AgentCare/config/
cp /Volumes/USB_NAME/.env ~/Desktop/AgentCare/
```

### Шаг 3️⃣: Запустить автоматическую настройку
```bash
cd ~/Desktop/AgentCare
bash setup_new_mac.sh
```

**Вот и всё!** ✨

Скрипт автоматически:
- ✅ Установит все необходимые инструменты
- ✅ Загрузит все зависимости
- ✅ Запустит все сервисы (FastAPI, Redis и т.д.)
- ✅ Проверит все соединения
- ✅ Убедится, что всё работает

---

## 📊 Чего ожидать

### Во время запуска:
```
╔════════════════════════════════════════╗
║  AgentCare - New MacBook Setup        ║
╚════════════════════════════════════════╝

ℹ Phase 1: Checking prerequisites...
ℹ Phase 2: Repository setup...
ℹ Phase 3: Installing Python dependencies...
ℹ Phase 4: Setting up credentials...
ℹ Phase 5: Setting up Docker...
ℹ Phase 6: Starting services...
ℹ Phase 7: Running health checks...
ℹ Phase 8: Service information...

========================================
✓ Startup Successful
========================================
```

### Время по этапам:
- Prerequisites: 2-5 мин (установка brew, git, python)
- Repository: 1 мин
- Python deps: 3-5 мин
- Docker: 5-10 мин
- Services: 2-3 мин
- **Итого: 15-25 минут**

---

## ✨ После завершения

### Проверить, что всё работает:
```bash
# Автоматически запустится в конце setup_new_mac.sh
# Но можно повторить в любой момент:
python scripts/health_check.py
```

**Результат должен быть:**
```
✓ OK: 18
⚠ WARNINGS: 1 (Supabase - опционально)
✗ FAILURES: 0
```

### Доступ к системе:

| Что | Где |
|-----|-----|
| **API Dashboard** | http://localhost:8000/docs |
| **Health Check** | http://localhost:8000/health |
| **Logs** | `tail -f logs/app.log` |
| **Services** | `docker-compose ps` |

---

## 🆘 Если что-то не работает

### Быстрая диагностика:
```bash
# Проверить статус сервисов
docker-compose ps

# Просмотреть логи
docker-compose logs -f

# Запустить полную проверку
python scripts/health_check.py
```

### Типичные проблемы:

**❌ "Docker daemon not running"**
```bash
# Запустить Docker Desktop
open /Applications/Docker.app
# Подождать 10 секунд, затем повторить setup
```

**❌ "credentials.json not found"**
```bash
# Проверить путь
ls -la config/credentials.json

# Если нет - скачать и скопировать в нужное место
# config/credentials.json должен быть в этой директории
```

**❌ "Module not found"**
```bash
# Активировать Python окружение
source .venv/bin/activate

# Переустановить зависимости
poetry install
```

**❌ "Cannot connect to http://localhost:8000"**
```bash
# Проверить контейнер
docker-compose ps agentcare

# Просмотреть логи
docker-compose logs agentcare

# Перезагрузить
docker-compose restart agentcare
```

---

## 📝 Что передается на новое устройство

✅ **Код:**
- Весь исходный код из репозитория
- Все скрипты и конфигурации
- Google Apps Script файлы

✅ **Настройки:**
- Конфигурации проектов (MT, SK, SS)
- Правила синхронизации
- Переменные окружения

✅ **Сервисы:**
- FastAPI приложение
- Redis кеш
- Docker контейнеры
- Соединения с Google Sheets
- Соединения с Supabase

✅ **Готовность к работе:**
- Все API ключи настроены
- Все базы данных доступны
- Все сервисы запущены
- Все проверки пройдены

---

## 🎯 Чек-лист успешной миграции

После завершения проверьте:

- [ ] Нет ошибок в setup.log
- [ ] `health_check.py` выводит 18+ OK
- [ ] `docker-compose ps` показывает 2 работающих контейнера
- [ ] API доступен: http://localhost:8000/docs
- [ ] Redis работает: `docker exec agentcare-redis redis-cli ping`
- [ ] Credentials загружены: `ls config/credentials.json`
- [ ] .env настроен: `cat .env | grep GEMINI_API_KEY`

---

## 📚 Полная документация

Если понадобятся более подробные инструкции:

- **SETUP_SUMMARY.md** - Полный обзор (с примерами команд)
- **MIGRATION_GUIDE.md** - Детальное руководство (40+ разделов)
- **scripts/health_check.py** - Запустить для полной диагностики

---

## ⏱️ Резюме

**Всё просто:**
1. Клонировать репо
2. Скопировать credentials
3. Запустить: `bash setup_new_mac.sh`
4. Дождаться завершения (~20 мин)
5. Наслаждаться новым MacBook! 🎉

**Вопросы?** Смотрите MIGRATION_GUIDE.md для полного руководства.

---

**Last Updated:** January 14, 2026
