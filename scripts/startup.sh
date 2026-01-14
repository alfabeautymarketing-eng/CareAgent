#!/bin/bash

# =================================================================
# AgentCare Startup Script
# Умный запуск всех сервисов с проверками
# =================================================================

# Цвета для вывода
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

echo -e "${GREEN}🚀 Запуск инфраструктуры AgentCare...${NC}"

# 1. Проверка окружения
echo -e "${YELLOW}Phase 1: Проверка окружения...${NC}"
if [ ! -f .env ]; then
    echo -e "${RED}❌ Ошибка: .env файл не найден!${NC}"
    echo -e "Создайте его из шаблона: cp .env.template .env"
    exit 1
fi

# 2. Проверка Docker
echo -e "${YELLOW}Phase 2: Запуск Docker контейнеров...${NC}"
if ! docker info > /dev/null 2>&1; then
    echo -e "${RED}❌ Ошибка: Docker не запущен!${NC}"
    exit 1
fi

# Запуск Redis и приложения через docker-compose
docker-compose up -d

# 3. Ожидание готовности Redis
echo -e "${YELLOW}Phase 3: Ожидание готовности Redis...${NC}"
RETRIES=10
while ! docker exec agentcare-redis redis-cli ping | grep PONG > /dev/null 2>&1; do
    echo "Ожидание Redis... ($RETRIES)"
    sleep 2
    RETRIES=$((RETRIES-1))
    if [ $RETRIES -eq 0 ]; then
        echo -e "${RED}❌ Redis не запустился вовремя${NC}"
        exit 1
    fi
done
echo -e "${GREEN}✅ Redis готов!${NC}"

# 4. Проверка Python зависимостей (в venv)
echo -e "${YELLOW}Phase 4: Проверка зависимостей...${NC}"
if [ -d .venv ]; then
    source .venv/bin/activate
else
    echo -e "${YELLOW}Виртуальное окружение не найдено. Попытка использовать глобальный Python...${NC}"
fi

# 5. Запуск FastAPI сервера (если не в докере, или проверка порта)
echo -e "${YELLOW}Phase 5: Запуск API сервера...${NC}"
# Если контейнер agentcare уже запущен через docker-compose, просто ждем
# Если нет, запускаем локально
if ! docker ps | grep agentcare-app > /dev/null 2>&1; then
    echo "Запуск FastAPI локально через poetry..."
    poetry run uvicorn src.main:app --host 0.0.0.0 --port 8000 --reload &
    SERVER_PID=$!
    echo "Сервер запущен с PID: $SERVER_PID"
fi

# 6. Финальная проверка здоровья
echo -e "${YELLOW}Phase 6: Выполнение Health Checks...${NC}"
sleep 5
python scripts/health_check.py

echo -e "${GREEN}========================================${NC}"
echo -e "${GREEN}🎉 Все сервисы AgentCare запущены!${NC}"
echo -e "${GREEN}API Dashboard: http://localhost:8000/docs${NC}"
echo -e "${GREEN}========================================${NC}"
