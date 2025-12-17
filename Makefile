.PHONY: help install dev run test lint format clean docker-build docker-up docker-down

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

install:
	poetry install

dev:
	poetry run uvicorn src.main:app --reload --host 0.0.0.0 --port 8000

run:
	poetry run uvicorn src.main:app --host 0.0.0.0 --port 8000 --workers 4

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
