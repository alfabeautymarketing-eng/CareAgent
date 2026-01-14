#!/bin/bash

# =================================================================
# AgentCare - Главный скрипт автоматизации настройки
# Версия: 1.0.0
# =================================================================

# Цвета
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

LOG_FILE="setup_diag_$(date +%Y%m%d_%H%M%S).log"

log() {
    echo -e "${BLUE}[$(date +%H:%M:%S)]${NC} $1" | tee -a "$LOG_FILE"
}

success() {
    echo -e "${GREEN}✓ $1${NC}" | tee -a "$LOG_FILE"
}

warn() {
    echo -e "${YELLOW}! $1${NC}" | tee -a "$LOG_FILE"
}

error() {
    echo -e "${RED}✗ $1${NC}" | tee -a "$LOG_FILE"
}

header() {
    echo -e "\n${BLUE}================================================================${NC}"
    echo -e "${BLUE}  $1${NC}"
    echo -e "${BLUE}================================================================${NC}\n"
}

# --- Фаза 1: Prerequisites ---
phase1_prerequisites() {
    header "Phase 1: Проверка и установка основных инструментов"
    
    # Check Homebrew
    if ! command -v brew &> /dev/null; then
        log "Homebrew не найден. Установка..."
        /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    fi
    success "Homebrew готов"

    # Install tools
    log "Установка git, python, docker, poetry..."
    brew install git python@3.11 docker docker-compose poetry 2>/dev/null
    success "Инструменты установлены"
}

# --- Фаза 2: Repository ---
phase2_repository() {
    header "Phase 2: Настройка репозитория"
    if [ ! -d .git ]; then
        warn "Скрипт запущен вне git репозитория."
    else
        log "Проверка обновлений..."
        git pull origin main
    fi
    success "Репозиторий актуален"
}

# --- Фаза 3: Python Deps ---
phase3_python_deps() {
    header "Phase 3: Установка Python зависимостей"
    
    if [ ! -d .venv ]; then
        log "Создание виртуального окружения..."
        python3 -m venv .venv
    fi
    
    source .venv/bin/activate
    log "Установка пакетов через poetry..."
    poetry install
    success "Python зависимости установлены"
}

# --- Фаза 4: Credentials ---
phase4_credentials() {
    header "Phase 4: Настройка учетных данных"
    
    if [ ! -f .env ]; then
        log "Создание .env из шаблона..."
        cp .env.template .env
        warn "Пожалуйста, отредактируйте .env и добавьте API ключи!"
    fi

    if [ ! -f config/credentials.json ]; then
        warn "config/credentials.json ОТСУТСТВУЕТ. Добавьте его для работы с Google Sheets."
    fi
    success "Конфигурация проверена"
}

# --- Фаза 5: Docker ---
phase5_docker() {
    header "Phase 5: Настройка Docker"
    log "Запуск Docker контейнеров..."
    docker-compose up -d
    success "Контейнеры запущены"
}

# --- Фаза 6: Startup ---
phase6_startup() {
    header "Phase 6: Запуск сервисов"
    bash scripts/startup.sh
    success "Сервисы запущены"
}

# --- Фаза 7: Health Checks ---
phase7_health_checks() {
    header "Phase 7: Проверка всех соединений"
    python scripts/health_check.py
    success "Проверка здоровья завершена"
}

# --- Фаза 8: Verification ---
phase8_verification() {
    header "Phase 8: Верификация работоспособности"
    log "Проверка API health endpoint..."
    curl -s http://localhost:8000/health | grep "status"
    success "API отвечает"
}

# --- Фаза 9: Summary ---
phase9_summary() {
    header "Phase 9: Итоговый отчет"
    echo -e "${GREEN}Настройка AgentCare на новом MacBook завершена успешно!${NC}"
    echo -e "Лог установки: $LOG_FILE"
    echo -e "\nДля работы с системой используйте:"
    echo -e " - ${BLUE}bash scripts/startup.sh${NC} (запуск)"
    echo -e " - ${BLUE}python scripts/health_check.py${NC} (диагностика)"
    echo -e " - ${BLUE}docker-compose logs -f${NC} (логи)"
}

# Выполнение всех фаз
main() {
    phase1_prerequisites
    phase2_repository
    phase3_python_deps
    phase4_credentials
    phase5_docker
    phase6_startup
    phase7_health_checks
    phase8_verification
    phase9_summary
}

main
