# Cloud Deployment Strategy для AgentCare

**Дата:** 2026-01-14
**Статус:** 📋 Planning
**Приоритет:** 🔴 HIGH (обеспечение reliability)

---

## 🎯 Проблема с текущей архитектурой

### ❌ Текущее состояние: Локальный сервер
```
Компьютер 1          Компьютер 2
    │                    │
    └─── Docker ────────┘
         (локальный)
```

**Проблемы:**
1. ⚠️ Зависания обновлений - не доходят до сервера
2. 📱 Синхронизация между 2 компьютерами нарушена
3. 🔌 Зависимость от конкретного компьютера (если он выключен - ничего не работает)
4. 🐛 Сложная отладка проблем синхронизации
5. 💾 Нет единого источника данных
6. 🚀 Невозможно масштабировать

---

## ✅ Целевая архитектура: Облачный сервер

```
Компьютер 1              Облачный VPS (DigitalOcean/AWS/Hetzner)
    │                              │
    ├─ Git Push ─────────────────────┤
    ├─ HTTP API Calls ──────────────→│ Docker Compose
    ├─ Google Sheets (via API) ────→│ ├─ FastAPI App
    └─ MCP Requests ─────────────────┤ ├─ PostgreSQL
                                      │ ├─ Redis
                                      │ └─ Nginx

Компьютер 2                            │
    │                                  │
    ├─ Git Pull ◄─────────────────────┤
    ├─ HTTP API Calls ──────────────→│
    ├─ Google Sheets (via API) ────→│
    └─ MCP Requests ─────────────────┤
```

---

## 📊 Сравнение архитектур

| Критерий | Локальный | Облачный |
|----------|-----------|----------|
| **Доступность** | 24/7 если PC включен | ✅ 24/7 гарантировано |
| **Синхронизация** | Нестабильна | ✅ Вся данных в БД |
| **Масштабируемость** | Нет | ✅ Легко (вертикальное/горизонтальное) |
| **CI/CD** | Ручной деплой | ✅ Автоматический на push |
| **Мониторинг** | Сложный | ✅ Встроенный (Prometheus, ELK) |
| **Резервные копии** | Только локальные | ✅ Автоматические (S3, Google Cloud) |
| **Затраты** | 0 (если есть PC) | 💵 5-20 USD/месяц |
| **Надежность** | Низкая | ✅ 99.9% SLA |

---

## 🚀 План миграции (Фазы)

### Фаза 1: Подготовка (1-2 дня)
- [ ] Выбрать облачного провайдера
- [ ] Создать VPS инстанс
- [ ] Установить Docker, Docker Compose
- [ ] Создать PostgreSQL БД (Supabase или облачная)
- [ ] Настроить домен и SSL сертификат

### Фаза 2: Развертывание (1 день)
- [ ] Загрузить код на облачный сервер
- [ ] Скопировать .env конфиги
- [ ] Запустить docker-compose
- [ ] Проверить здоровье сервиса
- [ ] Настроить автоматические резервные копии

### Фаза 3: Синхронизация (1 день)
- [ ] Мигрировать данные из локального Redis в облачный
- [ ] Синхронизировать Google Sheets данные
- [ ] Тестировать обе машины одновременно
- [ ] Убедиться что всё синхронизируется через облако

### Фаза 4: CI/CD (1-2 дня)
- [ ] Настроить GitHub Actions для автодеплоя
- [ ] Добавить проверки перед деплоем (тесты, линтер)
- [ ] Настроить rollback на случай ошибок
- [ ] Добавить мониторинг и алерты

### Фаза 5: Оптимизация (1 неделя)
- [ ] Настроить логирование (ELK или Datadog)
- [ ] Добавить кэширование на уровне CDN
- [ ] Оптимизировать запросы к БД
- [ ] Настроить автоскейлинг при необходимости

---

## 🌩️ Рекомендуемые провайдеры

### 1. **DigitalOcean** ⭐⭐⭐ (РЕКОМЕНДУЕМ)
```
Плюсы:
- Простой и понятный интерфейс
- Docker уже встроен в droplets
- Отличная документация
- Управляемые БД по хорошей цене

Цены:
- Droplet 2GB RAM, 2 CPU: $12/месяц
- PostgreSQL managed: $15/месяц
- Итого: ~$27/месяц

Настройка:
1. Создать droplet (Ubuntu 24.04)
2. Установить Docker Compose
3. Клонировать репо
4. Запустить docker-compose
```

### 2. **Hetzner**
```
Плюсы:
- Самая дешевая цена в Европе
- Отличное соотношение цена/производительность
- Немецкие серверы (хорошая скорость)

Цены:
- Cloud VPS 2GB: €4/месяц
- Managed Database (PostgreSQL): €5/месяц
- Итого: ~€9/месяц (~10 USD)
```

### 3. **AWS/Google Cloud**
```
Плюсы:
- Масштабируемость
- Управляемые сервисы (RDS, Cloud SQL)
- Интеграция с AWS/Google экосистемой

Минусы:
- Сложнее в настройке
- Дороже на малых объемах (~$30-50/месяц)

Рекомендуется когда нужна мощность
```

---

## 📋 Улучшенный docker-compose.yml для облака

```yaml
version: "3.8"

services:
  app:
    build: .
    container_name: agentcare
    ports:
      - "8000:8000"
    env_file:
      - .env
    volumes:
      - ./config:/app/config:rw
      - ./logs:/app/logs
    environment:
      - REDIS_URL=redis://redis:6379/0
      - DATABASE_URL=${DATABASE_URL}  # Облачная БД
      - ENVIRONMENT=production
    depends_on:
      redis:
        condition: service_healthy
    restart: always
    networks:
      - agentcare
    healthcheck:
      test: ["CMD", "curl", "-f", "http://localhost:8000/api/v1/health"]
      interval: 30s
      timeout: 10s
      retries: 3

  redis:
    image: redis:7-alpine
    container_name: agentcare-redis
    ports:
      - "127.0.0.1:6379:6379"  # Только локально в контейнере
    volumes:
      - redis_data:/data
    healthcheck:
      test: ["CMD", "redis-cli", "ping"]
      interval: 10s
      timeout: 5s
      retries: 3
    restart: always
    networks:
      - agentcare
    command: redis-server --appendonly yes  # Persistence

  # Nginx reverse proxy (опционально, но рекомендуем)
  nginx:
    image: nginx:alpine
    ports:
      - "80:80"
      - "443:443"
    volumes:
      - ./nginx.conf:/etc/nginx/nginx.conf:ro
      - ./certs:/etc/nginx/certs:ro  # SSL сертификаты
    depends_on:
      - app
    restart: always
    networks:
      - agentcare

networks:
  agentcare:
    driver: bridge

volumes:
  redis_data:
```

---

## 🔧 Скрипты для облачного развертывания

### setup-cloud.sh
```bash
#!/bin/bash
# Автоматическая настройка облачного сервера

# 1. Обновить систему
sudo apt update && sudo apt upgrade -y

# 2. Установить Docker
curl -fsSL https://get.docker.com -o get-docker.sh
sudo sh get-docker.sh
sudo usermod -aG docker $USER

# 3. Установить Docker Compose
sudo curl -L "https://github.com/docker/compose/releases/download/v2.20.0/docker-compose-$(uname -s)-$(uname -m)" -o /usr/local/bin/docker-compose
sudo chmod +x /usr/local/bin/docker-compose

# 4. Клонировать репо
git clone https://github.com/your-repo/AgentCare.git
cd AgentCare

# 5. Скопировать переменные окружения
cp .env.template .env
# EDIT .env WITH YOUR CLOUD CREDENTIALS

# 6. Запустить приложение
docker-compose up -d

# 7. Проверить здоровье
docker-compose exec app python scripts/health_check.py
```

---

## 🔄 CI/CD Pipeline (GitHub Actions)

```yaml
name: Deploy to Cloud

on:
  push:
    branches: [main]

jobs:
  deploy:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3

      - name: Deploy to VPS
        env:
          DEPLOY_KEY: ${{ secrets.CLOUD_DEPLOY_KEY }}
          CLOUD_HOST: ${{ secrets.CLOUD_HOST }}
        run: |
          mkdir -p ~/.ssh
          echo "$DEPLOY_KEY" > ~/.ssh/id_ed25519
          chmod 600 ~/.ssh/id_ed25519

          ssh -i ~/.ssh/id_ed25519 root@$CLOUD_HOST << 'EOF'
            cd /app/AgentCare
            git pull origin main
            docker-compose up -d --build
            docker-compose exec app python scripts/health_check.py
          EOF
```

---

## 💡 Преимущества облачной архитектуры

1. **Надежность:** Всегда доступно, даже если оба локальных ПК выключены
2. **Синхронизация:** Все данные в единой базе, автоматическая синхронизация
3. **Масштабируемость:** Легко добавить ресурсы при необходимости
4. **Безопасность:** SSL/TLS, автоматические обновления, резервные копии
5. **Отладка:** Логи в одном месте, простой мониторинг
6. **Автоматизация:** CI/CD автоматически деплоит изменения
7. **Стоимость:** Очень дешево на малых объемах (~10-30 USD/месяц)

---

## 📈 Примерная стоимость

### Вариант 1: DigitalOcean (РЕКОМЕНДУЕМ)
- Droplet 2GB RAM: $12/месяц
- Managed PostgreSQL: $15/месяц
- Spaces (S3-like, для резервных копий): $5/месяц
- **Итого: ~$32/месяц**

### Вариант 2: Hetzner (БЮДЖЕТНЫЙ)
- Cloud VPS 2GB: €4/месяц
- Managed PostgreSQL: €5/месяц
- S3-like storage: €3/месяц
- **Итого: ~€12/месяц (~13 USD)**

### Вариант 3: AWS (КОРПОРАТИВНЫЙ)
- EC2 t3.small: ~$20/месяц
- RDS PostgreSQL: ~$30/месяц
- Другие сервисы: ~$10/месяц
- **Итого: ~$60/месяц**

---

## 🎬 Следующие шаги

1. **Немедленно:**
   - [ ] Решить на DigitalOcean или Hetzner?
   - [ ] Создать облачный инстанс
   - [ ] Выделить себе 2-3 часа на настройку

2. **В ближайшие дни:**
   - [ ] Запустить сервер на облаке
   - [ ] Мигрировать данные
   - [ ] Тестировать с обе машины

3. **На этой неделе:**
   - [ ] Настроить CI/CD
   - [ ] Убедиться что всё синхронизируется

4. **В дальнейшем:**
   - [ ] Оптимизировать производительность
   - [ ] Добавить мониторинг
   - [ ] Настроить резервные копии

---

## ✅ Заключение

Облачное развертывание решит все текущие проблемы:
- ✅ Нет зависаний обновлений
- ✅ Синхронизация между машинами автоматическая
- ✅ Сервер всегда доступен
- ✅ Просто масштабировать
- ✅ Дешево (10-30 USD/месяц)

**Рекомендация:** Переходить на облако в этой неделе.

---

**Готовы ли вы начать миграцию?**
