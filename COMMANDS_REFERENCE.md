# AgentCare - Справочник команд

Полный список полезных команд для работы с AgentCare на новом MacBook.

## 🚀 Запуск и остановка сервисов

### Запустить все сервисы
```bash
cd ~/Desktop/AgentCare

# С автоматическими проверками (рекомендуется)
bash scripts/startup.sh

# Или просто через Docker Compose
docker-compose up -d

# С видимыми логами в консоли
docker-compose up
```

### Остановить все сервисы
```bash
docker-compose down
```

### Перезагрузить сервисы
```bash
docker-compose restart

# Или перезагрузить конкретный сервис
docker-compose restart agentcare
docker-compose restart agentcare-redis
```

### Пересобрать контейнеры
```bash
docker-compose up -d --build
```

---

## 🔍 Проверка статуса и здоровья системы

### Полная проверка здоровья (все соединения)
```bash
python scripts/health_check.py
```

### Быстрая проверка статуса контейнеров
```bash
docker-compose ps
```

### API health endpoint
```bash
curl http://localhost:8000/health | python -m json.tool
```

### Проверить Redis
```bash
docker exec agentcare-redis redis-cli ping
# Должно вывести: PONG
```

### Информация о Redis
```bash
docker exec agentcare-redis redis-cli info stats
```

---

## 📝 Просмотр логов

### Все логи (все контейнеры)
```bash
docker-compose logs -f
```

### Логи конкретного сервиса
```bash
# FastAPI логи
docker-compose logs -f agentcare

# Redis логи
docker-compose logs -f agentcare-redis
```

### Последние N строк логов
```bash
docker-compose logs --tail 50 agentcare
```

### Логи с временной меткой
```bash
docker-compose logs -f --timestamps agentcare
```

### Логи приложения (файл)
```bash
tail -f logs/app.log

# Последние 100 строк
tail -100 logs/app.log

# В реальном времени
tail -f logs/app.log
```

### Логи синхронизации
```bash
# Просмотреть директорию логов
ls -lah data/sync_logs/

# Последний лог
cat data/sync_logs/$(ls -t data/sync_logs/ | head -1)
```

---

## 🔧 Python и зависимости

### Активировать виртуальное окружение
```bash
source .venv/bin/activate
```

### Деактивировать
```bash
deactivate
```

### Установить/обновить зависимости
```bash
poetry install

# Обновить все пакеты
poetry update

# Установить новый пакет
poetry add package_name
```

### Проверить установленные пакеты
```bash
poetry show

# Только основные
poetry show --only main
```

---

## 🌐 Доступ к сервисам

### FastAPI Dashboard (Swagger UI)
```
http://localhost:8000/docs
```

### ReDoc (альтернативная документация)
```
http://localhost:8000/redoc
```

### Health check endpoint
```
http://localhost:8000/health
```

### OpenAPI JSON
```
http://localhost:8000/openapi.json
```

---

## 🧪 Тестирование API

### Тестировать health endpoint
```bash
curl http://localhost:8000/health
```

### С красивым форматом JSON
```bash
curl http://localhost:8000/health | python -m json.tool
```

### С заголовками
```bash
curl -i http://localhost:8000/health
```

### POST запрос
```bash
curl -X POST http://localhost:8000/api/v1/endpoint \
  -H "Content-Type: application/json" \
  -d '{"key": "value"}'
```

---

## 🗄️ Работа с данными

### Просмотреть структуру данных
```bash
ls -la data/

# Со включением скрытых файлов
ls -lah data/
```

### Размер синхронизированных данных
```bash
du -sh data/
du -sh data/sync_logs/
du -sh data/function_logs/
```

### Очистить логи (старше 90 дней - автоматически)
```bash
# Скрипт автоматически управляет ротацией логов
# Но можно очистить вручную:
rm -rf data/sync_logs/old_*.json
```

---

## 🔐 Работа с credentials

### Проверить, есть ли credentials
```bash
ls -la config/credentials.json
```

### Просмотреть информацию о credentials (без раскрытия секрета)
```bash
python -c "
import json
with open('config/credentials.json') as f:
    data = json.load(f)
    print(f'Service Account: {data.get(\"client_email\")}')
    print(f'Project ID: {data.get(\"project_id\")}')
"
```

### Тестировать Google API подключение
```bash
python -c "
from google.oauth2.service_account import Credentials

creds = Credentials.from_service_account_file('config/credentials.json')
print(f'✓ Credentials loaded successfully')
print(f'  Service Account: {creds.service_account_email}')
"
```

---

## 🌍 Работа с .env файлом

### Просмотреть все переменные окружения
```bash
cat .env
```

### Редактировать .env
```bash
nano .env
# или
vim .env
```

### Проверить конкретную переменную
```bash
cat .env | grep GEMINI_API_KEY
cat .env | grep REDIS_URL
```

### Перезагрузить .env (если изменили в Docker)
```bash
docker-compose down
docker-compose up -d
```

---

## 🐳 Docker команды

### Просмотреть все контейнеры
```bash
docker ps -a
```

### Просмотреть образы
```bash
docker images | grep agentcare
```

### Просмотреть объёмы данных
```bash
docker volume ls
```

### Удалить неиспользуемые ресурсы
```bash
docker system prune -a
```

### Просмотреть ресурсы контейнера
```bash
docker stats agentcare
```

### Войти в контейнер (bash)
```bash
docker exec -it agentcare bash

# Выход: exit или Ctrl+D
```

### Выполнить команду в контейнере
```bash
docker exec agentcare python -m src.main --help
```

---

## 🔄 Git команды (для обновления кода)

### Проверить статус
```bash
git status
```

### Получить последние изменения
```bash
git pull origin main
```

### Просмотреть историю коммитов
```bash
git log --oneline | head -10
```

### Просмотреть ветки
```bash
git branch -a
```

### Переключиться на другую ветку
```bash
git checkout branch_name
```

### Создать новую ветку
```bash
git checkout -b new_branch_name
```

---

## 📊 Мониторинг производительности

### CPU и память контейнеров
```bash
docker stats
```

### Использование диска
```bash
df -h
# или для конкретной директории
du -sh ~/Desktop/AgentCare
```

### Redis память
```bash
docker exec agentcare-redis redis-cli info memory
```

### Статистика Redis
```bash
docker exec agentcare-redis redis-cli info stats
```

---

## 🛠️ Настройка автозапуска (опционально)

### Создать launchd сервис
```bash
# Создать файл
sudo nano /Library/LaunchDaemons/com.agentcare.plist
```

### Содержимое файла
```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>com.agentcare.startup</string>
  <key>ProgramArguments</key>
  <array>
    <string>bash</string>
    <string>/Users/USERNAME/Desktop/AgentCare/scripts/startup.sh</string>
  </array>
  <key>RunAtLoad</key>
  <true/>
  <key>StandardErrorPath</key>
  <string>/Users/USERNAME/Desktop/AgentCare/logs/launchd.log</string>
  <key>StandardOutPath</key>
  <string>/Users/USERNAME/Desktop/AgentCare/logs/launchd.log</string>
</dict>
</plist>
```

### Загрузить сервис
```bash
sudo launchctl load /Library/LaunchDaemons/com.agentcare.plist
```

### Выгрузить сервис
```bash
sudo launchctl unload /Library/LaunchDaemons/com.agentcare.plist
```

---

## 🔍 Полезные трюки

### Следить за файлом логов в реальном времени
```bash
tail -f logs/app.log | grep ERROR
```

### Найти ошибки в логах
```bash
docker-compose logs agentcare | grep -i error
```

### Посчитать строки в логе
```bash
wc -l logs/app.log
```

### Выполнить Python скрипт с окружением
```bash
source .venv/bin/activate && python your_script.py
```

### Запустить скрипт в фоне
```bash
bash scripts/startup.sh &
```

### Посмотреть все процессы Python
```bash
ps aux | grep python
```

### Убить процесс по порту
```bash
# Найти процесс на порту 8000
lsof -i :8000

# Убить процесс (replace PID)
kill -9 <PID>
```

---

## 🆘 Troubleshooting команды

### Полная диагностика системы
```bash
python scripts/health_check.py 2>&1 | tee diagnostic.log
```

### Проверить все соединения
```bash
# API
curl -s http://localhost:8000/health | python -m json.tool

# Redis
docker exec agentcare-redis redis-cli ping

# Docker
docker-compose ps

# Python deps
python -c "import fastapi; import redis; print('OK')"
```

### Пересоздать всю систему с нуля
```bash
# Остановить
docker-compose down

# Удалить образы
docker-compose rm

# Заново собрать
docker-compose up -d --build

# Проверить
python scripts/health_check.py
```

---

## 📱 Быстрые ссылки на важное

| Что нужно | Команда |
|-----------|---------|
| Запустить всё | `bash scripts/startup.sh` |
| Проверить здоровье | `python scripts/health_check.py` |
| Просмотреть логи | `docker-compose logs -f` |
| Остановить | `docker-compose down` |
| API документация | `http://localhost:8000/docs` |
| Помощь | `cat QUICK_MIGRATION.md` |

---

## 💡 Советы

1. **Всегда запускайте через `startup.sh`** - он делает все проверки
2. **Проверяйте логи при ошибках** - большинство проблем там описаны
3. **Используйте `health_check.py` для диагностики** - показывает полное состояние
4. **Активируйте venv перед Python командами** - чтобы работали нужные версии
5. **Сохраняйте backup перед изменениями** - чтобы можно было откатиться

---

**Last Updated:** January 14, 2026
**Questions?** See MIGRATION_GUIDE.md or run: `python scripts/health_check.py`
