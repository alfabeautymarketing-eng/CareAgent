#!/bin/bash

# ===============================================
# СКРИПТ ПРОВЕРКИ СИСТЕМЫ ЛОГИРОВАНИЯ AgentCare
# ===============================================

echo ""
echo "╔════════════════════════════════════════════════════════╗"
echo "║  📜 Проверка системы логирования AgentCare v1.0       ║"
echo "║     Дата: 3 января 2026                               ║"
echo "╚════════════════════════════════════════════════════════╝"
echo ""

# Цвета для вывода
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

TESTS_PASSED=0
TESTS_FAILED=0

# Функция для проверки
check() {
    echo ""
    echo "▶ $1"
}

# Функция для успеха
pass() {
    echo -e "${GREEN}✅ PASSED${NC}: $1"
    ((TESTS_PASSED++))
}

# Функция для ошибки
fail() {
    echo -e "${RED}❌ FAILED${NC}: $1"
    ((TESTS_FAILED++))
}

# ===================================================
# ЭТАП 1: Проверка файлов
# ===================================================
echo ""
echo "═══════════════════════════════════════════════════════"
echo "📁 ЭТАП 1: Проверка файлов"
echo "═══════════════════════════════════════════════════════"

check "Проверка src/services/function_log_service.py"
if [ -f "src/services/function_log_service.py" ]; then
    pass "Файл function_log_service.py существует"
else
    fail "Файл function_log_service.py не найден"
fi

check "Проверка config/logs_manager.html"
if [ -f "config/logs_manager.html" ]; then
    pass "Файл logs_manager.html существует"
else
    fail "Файл logs_manager.html не найден"
fi

check "Проверка документации"
for doc in "docs/LOGS_IMPLEMENTATION.md" "docs/LOGS_QUICK_REFERENCE.md" "docs/LOGS_RELEASE_NOTES.md" "docs/INTEGRATION_COMPLETE.md" "docs/TESTING_CHECKLIST.md"; do
    if [ -f "$doc" ]; then
        pass "Документ $doc существует"
    else
        fail "Документ $doc не найден"
    fi
done

# ===================================================
# ЭТАП 2: Проверка Python синтаксиса
# ===================================================
echo ""
echo "═══════════════════════════════════════════════════════"
echo "🐍 ЭТАП 2: Проверка Python синтаксиса"
echo "═══════════════════════════════════════════════════════"

check "Компиляция src/services/function_log_service.py"
if python3 -m py_compile src/services/function_log_service.py 2>/dev/null; then
    pass "function_log_service.py синтаксис OK"
else
    fail "function_log_service.py содержит ошибки синтаксиса"
fi

check "Компиляция src/services/logging.py"
if python3 -m py_compile src/services/logging.py 2>/dev/null; then
    pass "logging.py синтаксис OK"
else
    fail "logging.py содержит ошибки синтаксиса"
fi

check "Компиляция src/services/sync.py"
if python3 -m py_compile src/services/sync.py 2>/dev/null; then
    pass "sync.py синтаксис OK"
else
    fail "sync.py содержит ошибки синтаксиса"
fi

check "Компиляция src/api/endpoints.py"
if python3 -m py_compile src/api/endpoints.py 2>/dev/null; then
    pass "endpoints.py синтаксис OK"
else
    fail "endpoints.py содержит ошибки синтаксиса"
fi

# ===================================================
# ЭТАП 3: Проверка импортов
# ===================================================
echo ""
echo "═══════════════════════════════════════════════════════"
echo "📦 ЭТАП 3: Проверка импортов"
echo "═══════════════════════════════════════════════════════"

check "Импорт FunctionLogService"
if python3 -c "from src.services.function_log_service import FunctionLogService" 2>/dev/null; then
    pass "FunctionLogService импортируется корректно"
else
    fail "Ошибка импорта FunctionLogService"
fi

check "Импорт FunctionLogContext"
if python3 -c "from src.services.function_log_service import FunctionLogContext" 2>/dev/null; then
    pass "FunctionLogContext импортируется корректно"
else
    fail "Ошибка импорта FunctionLogContext"
fi

check "Импорт LoggingService"
if python3 -c "from src.services.logging import LoggingService" 2>/dev/null; then
    pass "LoggingService импортируется корректно"
else
    fail "Ошибка импорта LoggingService"
fi

# ===================================================
# ЭТАП 4: Проверка структуры данных
# ===================================================
echo ""
echo "═══════════════════════════════════════════════════════"
echo "📊 ЭТАП 4: Проверка структуры данных"
echo "═══════════════════════════════════════════════════════"

check "Проверка SyncLogEntry"
if python3 -c "from src.services.sync_log_service import SyncLogEntry" 2>/dev/null; then
    pass "SyncLogEntry модель определена"
else
    fail "SyncLogEntry не найдена"
fi

check "Проверка FunctionStep"
if python3 -c "from src.services.function_log_service import FunctionStep" 2>/dev/null; then
    pass "FunctionStep модель определена"
else
    fail "FunctionStep не найдена"
fi

# ===================================================
# ЭТАП 5: Проверка директорий хранилища
# ===================================================
echo ""
echo "═══════════════════════════════════════════════════════"
echo "💾 ЭТАП 5: Проверка директорий хранилища"
echo "═══════════════════════════════════════════════════════"

check "Создание директории /data/sync_logs"
mkdir -p "data/sync_logs" 2>/dev/null
if [ -d "data/sync_logs" ]; then
    pass "Директория /data/sync_logs готова"
else
    fail "Не удалось создать /data/sync_logs"
fi

check "Создание директории /data/function_logs"
mkdir -p "data/function_logs" 2>/dev/null
if [ -d "data/function_logs" ]; then
    pass "Директория /data/function_logs готова"
else
    fail "Не удалось создать /data/function_logs"
fi

# ===================================================
# ЭТАП 6: Проверка размеров файлов
# ===================================================
echo ""
echo "═══════════════════════════════════════════════════════"
echo "📏 ЭТАП 6: Проверка размеров файлов"
echo "═══════════════════════════════════════════════════════"

check "Размер function_log_service.py"
SIZE=$(wc -l < src/services/function_log_service.py)
if [ "$SIZE" -gt 300 ]; then
    pass "function_log_service.py: $SIZE строк"
else
    fail "function_log_service.py: $SIZE строк (ожидается >300)"
fi

check "Размер logs_manager.html"
SIZE=$(wc -l < config/logs_manager.html)
if [ "$SIZE" -gt 800 ]; then
    pass "logs_manager.html: $SIZE строк"
else
    fail "logs_manager.html: $SIZE строк (ожидается >800)"
fi

# ===================================================
# ЭТАП 7: Вывод результатов
# ===================================================
echo ""
echo "═══════════════════════════════════════════════════════"
echo "📋 ФИНАЛЬНЫЙ РЕЗУЛЬТАТ"
echo "═══════════════════════════════════════════════════════"
echo ""

TOTAL=$((TESTS_PASSED + TESTS_FAILED))

echo -e "✅ Пройдено:  ${GREEN}$TESTS_PASSED/$TOTAL${NC}"
echo -e "❌ Ошибок:   ${RED}$TESTS_FAILED/$TOTAL${NC}"
echo ""

if [ $TESTS_FAILED -eq 0 ]; then
    echo "╔════════════════════════════════════════════════════════╗"
    echo "║         🚀 ВСЕ ТЕСТЫ ПРОЙДЕНЫ УСПЕШНО! 🚀           ║"
    echo "║                                                        ║"
    echo "║   Система логирования готова к использованию!         ║"
    echo "║                                                        ║"
    echo "║   Далее:                                              ║"
    echo "║   1. Откройте веб-UI:                                 ║"
    echo "║      config/logs_manager.html?id=<SPREADSHEET_ID>    ║"
    echo "║                                                        ║"
    echo "║   2. Смотрите документацию:                           ║"
    echo "║      docs/LOGS_QUICK_REFERENCE.md                    ║"
    echo "║                                                        ║"
    echo "║   3. Запустите полный чек-лист:                      ║"
    echo "║      docs/TESTING_CHECKLIST.md                       ║"
    echo "╚════════════════════════════════════════════════════════╝"
    echo ""
    exit 0
else
    echo "╔════════════════════════════════════════════════════════╗"
    echo "║           ⚠️  НЕКОТОРЫЕ ТЕСТЫ НЕ ПРОЙДЕНЫ  ⚠️        ║"
    echo "║                                                        ║"
    echo "║   Пожалуйста, проверьте ошибки выше и повторите.    ║"
    echo "║                                                        ║"
    echo "║   Помощь:                                             ║"
    echo "║   - docs/LOGS_RELEASE_NOTES.md                       ║"
    echo "║   - docs/INTEGRATION_COMPLETE.md                     ║"
    echo "╚════════════════════════════════════════════════════════╝"
    echo ""
    exit 1
fi
