---
description: Тестирование и деплой AgentCare сервера
---

# Workflow: Test & Deploy

Этот workflow позволяет быстро протестировать изменения локально и задеплоить на продакшн.

## Шаги

### ⚡ ЕДИНАЯ КОМАНДА: Локальная разработка + Таблицы
// turbo
```bash
cd /Users/aleksandr/Desktop/AgentCare && make dev
```
Эта команда:
1. Запустит локальный сервер.
2. Поднимет туннель ngrok.
3. **Сама** обновит URL во всех ваших Google Таблицах.
4. При нажатии `Ctrl+C` — вернет всё как было (на VPS).

### 2. Запуск тестов API
// turbo
```bash
cd /Users/aleksandr/Desktop/AgentCare && .venv/bin/python scripts/test_menu_endpoints.py
```

### 3. Проверка health локального сервера
// turbo
```bash
curl -s http://localhost:8000/health | python3 -m json.tool
```

### 4. Деплой на VPS (когда всё готово)
```bash
cd /Users/aleksandr/Desktop/AgentCare && make deploy
```
⚠️ Это синхронизирует код на VPS и перезапустит Docker контейнер.

### 5. Проверка health VPS сервера
// turbo
```bash
curl -s http://46.226.167.153:8000/health | python3 -m json.tool
```

### 6. Просмотр логов VPS (опционально)
```bash
ssh root@46.226.167.153 "cd ~/AgentCare && docker-compose logs -f app --tail=50"
```

---

## Быстрые команды

| Действие | Команда |
|----------|---------|
| Локальный сервер | `.venv/bin/uvicorn src.main:app --reload` |
| Тесты | `.venv/bin/python scripts/test_menu_endpoints.py` |
| Деплой | `make deploy` |
| Логи VPS | `ssh root@46.226.167.153 "docker-compose logs -f app"` |
