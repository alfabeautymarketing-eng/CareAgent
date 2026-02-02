.PHONY: help install dev run test lint format clean docker-build docker-up docker-down deploy

help:
	@echo "AgentCare - Available commands:"
	@echo ""
	@echo "  install     Install dependencies"
	@echo "  dev         Run in development mode"
	@echo "  run         Run in production mode"
	@echo "  test        Run tests"
	@echo "  lint        Run linters"
	@echo "  format      Format code"
	@echo "  clean       Clean cache files"
	@echo ""
	@echo "  docker-build  Build Docker image"
	@echo "  docker-up     Start Docker containers"
	@echo "  docker-down   Stop Docker containers"
	@echo "  gas-push      Push GAS to ALL projects (MT, SS, SK)"
	@echo "  deploy        Full deploy (Server + GAS)"

# === Основные команды ===

install:
	poetry install

dev:
	bash ./scripts/dev_full.sh

test:
	poetry run pytest -v --cov=src --cov-report=term-missing

lint:
	poetry run ruff check src tests
	poetry run mypy src

format:
	poetry run black src tests
	poetry run ruff check --fix src tests

clean:
	find . -type d -name "__pycache__" -exec rm -rf {} + 2>/dev/null || true
	find . -type f -name "*.pyc" -delete 2>/dev/null || true
	rm -rf .pytest_cache .mypy_cache .coverage htmlcov 2>/dev/null || true

docker-build:
	docker compose build

docker-up:
	docker compose up -d

docker-down:
	docker compose down

docker-logs:
	docker compose logs -f app

gas-push:
	bash ./scripts/push_all_gas.sh

deploy:
	./scripts/deploy_all.sh

# === Локальная разработка ===

# Запуск локального сервера (venv)
local:
	.venv/bin/uvicorn src.main:app --host 0.0.0.0 --port 8000 --reload

# Тестирование API endpoints
test-api:
	.venv/bin/python scripts/test_menu_endpoints.py

# Проверка health локального сервера
health-local:
	@curl -s http://localhost:8000/health | python3 -m json.tool

# Проверка health VPS сервера
health-vps:
	@curl -s http://46.226.167.153:8000/health | python3 -m json.tool

# Перезагрузка основного сервера на VPS
restart-vps:
	@echo "🔄 Restarting AgentCare server on VPS..."
	ssh root@46.226.167.153 "cd ~/AgentCare && docker-compose restart app"
	@echo "✅ Server restarted."

# Полный цикл: тест + деплой (если тесты прошли)
test-deploy:
	@echo "🔍 Running local tests..."
	.venv/bin/python scripts/test_menu_endpoints.py && \
	echo "" && \
	echo "✅ Tests passed! Deploying to VPS..." && \
	./scripts/deploy_all.sh

# Запуск локального сервера + туннеля (информационная команда)
dev-full:
	@echo "1. В одном терминале: make local"
	@echo "2. В другом терминале: make tunnel"


