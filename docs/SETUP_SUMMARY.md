# AgentCare - Быстрый Старт (SETUP_SUMMARY.md)

Краткий обзор команд и процедур для запуска системы на новом устройстве.

## 🏁 Команда для ленивых (запуск всего)
```bash
bash scripts/startup.sh
```

## 📋 Чек-лист проверки системы
1. [ ] Файл `.env` существует и заполнен
2. [ ] Файл `config/credentials.json` содержит валидный Service Account
3. [ ] Выполнен `docker-compose up -d`
4. [ ] Результат `python scripts/health_check.py` — "Result: 6/6 checks passed"

---

## 🛠 Полезные команды (Справочник)

| Действие | Команда |
|-----------|---------|
| **Запуск всего** | `bash scripts/startup.sh` |
| **Остановка всего** | `docker-compose down` |
| **Проверка здоровья**| `python scripts/health_check.py` |
| **Просмотр логов** | `docker-compose logs -f agentcare` |
| **Перезагрузка** | `docker-compose restart` |
| **Swagger/Документация** | http://localhost:8000/docs |

---

## 🔍 Куда смотреть при ошибках?
- **Логи приложения**: `logs/app.log`
- **Логи синхронизации**: `data/sync_logs/`
- **Консоль Docker**: `docker-compose logs -f`

## 🧪 Экспресс-проверка API
```bash
curl http://localhost:8000/health
```
Ожидаемый ответ: `{"status": "ok", "version": "..."}`

---
**Обновлено:** 14 января 2026 г.
