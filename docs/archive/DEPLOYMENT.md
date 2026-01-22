# AgentCare - Deployment Guide

## Требования

### Сервер

| Параметр | Минимум | Рекомендуется |
|----------|---------|---------------|
| CPU | 2 cores | 4 cores |
| RAM | 4 GB | 8 GB |
| SSD | 40 GB | 80 GB |
| OS | Ubuntu 22.04+ | Ubuntu 24.04 |

### Программное обеспечение

- Docker 24+
- Docker Compose 2.20+
- (альтернативно) Python 3.11+, Redis 7+

---

## Способ 1: Docker (рекомендуется)

### 1.1 Клонирование

```bash
git clone https://github.com/your-repo/agentcare.git
cd agentcare
```

### 1.2 Настройка переменных окружения

```bash
cp .env.example .env
nano .env
```

**.env файл:**
```bash
# App
APP_ENV=production
APP_DEBUG=false
APP_SECRET_KEY=your-secret-key-here

# Server
SERVER_HOST=0.0.0.0
SERVER_PORT=8000
SERVER_WORKERS=4

# Google
GOOGLE_CREDENTIALS_BASE64=<base64 encoded credentials.json>
# или путь к файлу
# GOOGLE_CREDENTIALS_FILE=/app/config/credentials.json

# Gemini AI
GEMINI_API_KEY=your-gemini-api-key

# Redis
REDIS_URL=redis://redis:6379/0

# Webhook Security
WEBHOOK_SECRET=your-webhook-secret

# Notifications (опционально)
TELEGRAM_BOT_TOKEN=your-bot-token
TELEGRAM_CHAT_ID=your-chat-id
```

### 1.3 Запуск

```bash
# Сборка и запуск
docker compose up -d

# Проверка логов
docker compose logs -f app

# Проверка здоровья
curl http://localhost:8000/health
```

### 1.4 docker-compose.yml

```yaml
version: '3.8'

services:
  app:
    build: .
    ports:
      - "8000:8000"
    environment:
      - APP_ENV=production
    env_file:
      - .env
    volumes:
      - ./config:/app/config:ro
      - ./logs:/app/logs
    depends_on:
      redis:
        condition: service_healthy
    restart: unless-stopped
    healthcheck:
      test: ["CMD", "curl", "-f", "http://localhost:8000/health"]
      interval: 30s
      timeout: 10s
      retries: 3

  redis:
    image: redis:7-alpine
    volumes:
      - redis_data:/data
    restart: unless-stopped
    healthcheck:
      test: ["CMD", "redis-cli", "ping"]
      interval: 10s
      timeout: 5s
      retries: 3

  # Опционально: Nginx reverse proxy
  nginx:
    image: nginx:alpine
    ports:
      - "80:80"
      - "443:443"
    volumes:
      - ./nginx.conf:/etc/nginx/nginx.conf:ro
      - ./certs:/etc/nginx/certs:ro
    depends_on:
      - app
    restart: unless-stopped

volumes:
  redis_data:
```

### 1.5 Dockerfile

```dockerfile
FROM python:3.11-slim

WORKDIR /app

# System dependencies
RUN apt-get update && apt-get install -y \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Python dependencies
COPY pyproject.toml poetry.lock ./
RUN pip install poetry && \
    poetry config virtualenvs.create false && \
    poetry install --no-dev --no-interaction

# Application
COPY src/ ./src/
COPY config/ ./config/

# User
RUN useradd -m appuser && chown -R appuser:appuser /app
USER appuser

# Run
EXPOSE 8000
CMD ["uvicorn", "src.main:app", "--host", "0.0.0.0", "--port", "8000"]
```

---

## Способ 2: Ручная установка

### 2.1 Установка зависимостей

```bash
# Python
sudo apt update
sudo apt install python3.11 python3.11-venv python3-pip

# Redis
sudo apt install redis-server
sudo systemctl enable redis-server
sudo systemctl start redis-server
```

### 2.2 Настройка проекта

```bash
# Клонирование
git clone https://github.com/your-repo/agentcare.git
cd agentcare

# Virtual environment
python3.11 -m venv venv
source venv/bin/activate

# Зависимости
pip install poetry
poetry install --no-dev
```

### 2.3 Systemd сервис

```bash
sudo nano /etc/systemd/system/agentcare.service
```

```ini
[Unit]
Description=AgentCare Service
After=network.target redis.service

[Service]
Type=simple
User=www-data
WorkingDirectory=/opt/agentcare
Environment="PATH=/opt/agentcare/venv/bin"
EnvironmentFile=/opt/agentcare/.env
ExecStart=/opt/agentcare/venv/bin/uvicorn src.main:app --host 0.0.0.0 --port 8000 --workers 4
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl daemon-reload
sudo systemctl enable agentcare
sudo systemctl start agentcare
```

---

## Настройка Google Cloud

### 3.1 Создание проекта

1. Перейдите в [Google Cloud Console](https://console.cloud.google.com/)
2. Создайте новый проект (например, "agentcare-prod")
3. Запомните Project ID

### 3.2 Включение API

```bash
# Установите gcloud CLI
# https://cloud.google.com/sdk/docs/install

gcloud auth login
gcloud config set project YOUR_PROJECT_ID

# Включите необходимые API
gcloud services enable sheets.googleapis.com
gcloud services enable drive.googleapis.com
gcloud services enable generativelanguage.googleapis.com
```

### 3.3 Service Account

```bash
# Создание service account
gcloud iam service-accounts create agentcare \
    --display-name="AgentCare Service Account"

# Скачивание ключа
gcloud iam service-accounts keys create credentials.json \
    --iam-account=agentcare@YOUR_PROJECT_ID.iam.gserviceaccount.com
```

### 3.4 Доступ к таблицам

**Важно:** Дайте service account доступ к вашим Google Sheets:
1. Откройте таблицу
2. Нажмите "Share"
3. Добавьте email service account (из credentials.json)
4. Дайте права "Editor"

### 3.5 Gemini API Key

1. Перейдите в [Google AI Studio](https://makersuite.google.com/app/apikey)
2. Создайте API key
3. Добавьте в .env: `GEMINI_API_KEY=...`

---

## Настройка Webhook

### 4.1 Google Apps Script (минимальный)

В каждой таблице (MT/SK/SS) создайте Apps Script:

```javascript
// Webhook URL вашего сервера
const WEBHOOK_URL = "https://your-server.com/webhook/sheets/mt";
const WEBHOOK_SECRET = "your-webhook-secret";

function onEdit(e) {
  sendWebhook("onEdit", e);
}

function onChange(e) {
  sendWebhook("onChange", e);
}

function sendWebhook(eventType, e) {
  const payload = {
    event: eventType,
    project: "mt",  // или sk, ss
    sheet: e.source.getActiveSheet().getName(),
    range: e.range ? e.range.getA1Notation() : null,
    timestamp: new Date().toISOString(),
    user: Session.getActiveUser().getEmail()
  };

  const signature = Utilities.computeHmacSha256Signature(
    JSON.stringify(payload),
    WEBHOOK_SECRET
  );
  const signatureHex = signature.map(b =>
    ('0' + (b & 0xFF).toString(16)).slice(-2)
  ).join('');

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    headers: {
      "X-Webhook-Signature": "sha256=" + signatureHex,
      "X-Webhook-Timestamp": Date.now().toString()
    },
    muteHttpExceptions: true
  };

  try {
    UrlFetchApp.fetch(WEBHOOK_URL, options);
  } catch (error) {
    console.error("Webhook error:", error);
  }
}
```

### 4.2 Установка триггеров

В Apps Script:
1. Triggers → Add Trigger
2. Function: `onEdit`, Event: `On edit`
3. Function: `onChange`, Event: `On change`

---

## SSL/HTTPS

### 5.1 Let's Encrypt (рекомендуется)

```bash
# Установка certbot
sudo apt install certbot python3-certbot-nginx

# Получение сертификата
sudo certbot --nginx -d your-domain.com

# Автообновление
sudo certbot renew --dry-run
```

### 5.2 Nginx конфигурация

```nginx
server {
    listen 80;
    server_name your-domain.com;
    return 301 https://$server_name$request_uri;
}

server {
    listen 443 ssl http2;
    server_name your-domain.com;

    ssl_certificate /etc/letsencrypt/live/your-domain.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/your-domain.com/privkey.pem;

    location / {
        proxy_pass http://localhost:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }

    location /ws/ {
        proxy_pass http://localhost:8000;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
    }
}
```

---

## Мониторинг

### 6.1 Логи

```bash
# Docker
docker compose logs -f app

# Systemd
journalctl -u agentcare -f

# Файлы логов
tail -f /opt/agentcare/logs/app.log
```

### 6.2 Health check

```bash
# Добавьте в cron
*/5 * * * * curl -sf http://localhost:8000/health || echo "AgentCare is down" | mail -s "Alert" admin@example.com
```

### 6.3 Prometheus metrics (опционально)

Добавьте endpoint `/metrics` для Prometheus:

```python
# src/api/metrics.py
from prometheus_client import Counter, Histogram, generate_latest

REQUEST_COUNT = Counter('requests_total', 'Total requests', ['endpoint', 'status'])
REQUEST_LATENCY = Histogram('request_latency_seconds', 'Request latency')
```

---

## Backup

### 7.1 Redis данные

```bash
# Backup
redis-cli BGSAVE
cp /var/lib/redis/dump.rdb /backup/redis-$(date +%Y%m%d).rdb

# Restore
sudo systemctl stop redis
cp /backup/redis-20240115.rdb /var/lib/redis/dump.rdb
sudo systemctl start redis
```

### 7.2 Конфигурация

```bash
# Backup configs
tar -czf /backup/agentcare-config-$(date +%Y%m%d).tar.gz \
    /opt/agentcare/.env \
    /opt/agentcare/config/
```

---

## Обновление

```bash
# Остановка
docker compose down

# Обновление кода
git pull origin main

# Пересборка
docker compose build --no-cache

# Запуск
docker compose up -d

# Проверка
docker compose logs -f app
```

---

## Troubleshooting

### Ошибка: "Google API quota exceeded"

```bash
# Проверьте лимиты в Google Cloud Console
# Увеличьте квоты или добавьте задержки между запросами
```

### Ошибка: "Redis connection refused"

```bash
# Проверьте статус Redis
sudo systemctl status redis

# Проверьте подключение
redis-cli ping
```

### Ошибка: "Webhook signature invalid"

```bash
# Убедитесь что WEBHOOK_SECRET одинаковый
# в .env и в Google Apps Script
```

### Высокое использование памяти

```bash
# Проверьте память
docker stats

# Ограничьте в docker-compose.yml
services:
  app:
    deploy:
      resources:
        limits:
          memory: 2G
```

---

## Рекомендуемые VPS провайдеры

| Провайдер | План | Цена | Комментарий |
|-----------|------|------|-------------|
| Hetzner | CX31 | €8/мес | Лучшая цена в Европе |
| DigitalOcean | Basic 4GB | $24/мес | Простой UI |
| Vultr | High Frequency | $18/мес | Быстрые диски |
| Linode | Dedicated 4GB | $30/мес | Стабильный |

**Выбор:** Hetzner CX31 — оптимальное соотношение цена/качество.
