# AgentCare AI Development Handbook

**Version:** 4.0  
**Last Updated:** 2026-01-29  
**Architecture:** [ARCHITECTURE.md](./ARCHITECTURE.md)

---

## 🎯 Core Principles

### 1. Server-First Philosophy
- **Python Server (FastAPI)** — единственный источник бизнес-логики
- **Google Sheets (GAS)** — тонкий клиент, только UI и события
- **Supabase** — хранилище логов, конфигурации, кэша

### 2. Menu-Driven Architecture
- **Структура кода = структура меню** — файлы роутеров соответствуют пунктам меню
- **Правила именования**: `<номер>m-<название>.py` для функций меню, `s_<название>.py` для серверных функций
- **Детали**: См. [docs/NAMING_CONVENTIONS.md](./docs/NAMING_CONVENTIONS.md)

### 3. Logging-First Debugging
- Детальное логирование каждой функции
- Журнал UI для анализа проблем
- Правило: "Сначала лог, потом фикс"

---

## 📂 Project Structure

```
AgentCare/
├── src/
│   ├── api/
│   │   ├── routes/
│   │   │   ├── 01m-zakaz.py          # 🧾 Заказ
│   │   │   ├── 02m-stadii_zakaza.py  # 📊 Стадии
│   │   │   ├── ...                   # (8 файлов для меню)
│   │   │   ├── s_sync.py             # Синхронизация
│   │   │   └── s_*.py                # Серверные функции
│   │   ├── models/                   # Pydantic модели
│   │   ├── router.py                 # Главный роутер
│   │   └── ...
│   ├── services/                     # Бизнес-логика
│   ├── config/                       # Конфигурация
│   │   └── project_menus.py          # Структура меню
│   └── main.py                       # Точка входа
├── gas/                              # Google Apps Script
├── docs/                             # Документация
│   └── NAMING_CONVENTIONS.md         # Правила именования
└── tests/                            # Тесты
```

---

## 🔧 Development Workflow

### Before Starting Work

1. **Check existing documentation:**
   - Read [ARCHITECTURE.md](./ARCHITECTURE.md)
   - Check [NAMING_CONVENTIONS.md](./docs/NAMING_CONVENTIONS.md)
   - Review [task.md](./task.md) if exists

2. **Understand the menu structure:**
   - Open `src/config/project_menus.py`
   - Find which menu group relates to your task
   - Locate corresponding router file (`01m-zakaz.py`, etc.)

3. **Create task.md** (if complex work):
   - Break down work into checklist items
   - Track progress as you go

### During Development

1. **Follow naming conventions:**
   - Menu functions → `<номер>m-<название>.py`
   - Server functions → `s_<название>.py`

2. **Document everything:**
   - Add docstrings with button mapping
   - Explain what, why, and how
   - Include example usage

3. **Log extensively:**
   - Use `logger.info()` for important steps
   - Use `logger.error()` for failures
   - Include context in logs

4. **Test thoroughly:**
   - Run existing tests: `pytest tests/`
   - Test via GAS buttons
   - Check `/docs` for API correctness

### After Completion

1. **Update documentation:**
   - Update relevant .md files
   - Keep ARCHITECTURE.md in sync

2. **Verify backwards compatibility:**
   - API paths unchanged
   - JSON formats unchanged
   - GAS functions still work

---

## 🚫 What NOT to Do

1. ❌ **Never** put business logic in GAS
2. ❌ **Never** change API paths without migration plan
3. ❌ **Never** skip logging in critical functions
4. ❌ **Never** create files without following naming conventions
5. ❌ **Never** modify code without understanding menu structure

---

## 📋 Quick Reference

### Finding Code by Menu Button

1. Look at button label in Google Sheets menu
2. Match to menu group in `src/config/project_menus.py`
3. Open corresponding `<номер>m-*.py` file
4. Search for function name from menu config

**Example:**
- Button: "📥 Обработка" in "🧾 Заказ" menu
- Function: `serverProcessPrimaryData`
- File: `src/api/routes/01m-zakaz.py`
- Endpoint: `POST /api/v1/price/process/{project}`

### Creating New Menu Function

1. Add to `src/config/project_menus.py`
2. Add endpoint to corresponding `<номер>m-*.py`
3. Add function in GAS that calls endpoint
4. Document mapping in router docstring

### Creating New Server Function

1. Create/update `s_<название>.py`
2. Add endpoints with proper documentation
3. Update router imports in `src/api/router.py`
4. Add tests if needed

---

## 🔗 Important Links

- [Architecture Overview](./ARCHITECTURE.md)
- [Naming Conventions](./docs/NAMING_CONVENTIONS.md)
- [API Documentation](./docs/API.md)
- [Deployment Guide](./DEPLOYMENT_STATUS.md)

---

- Always check menu structure first
- Use Context7 for library documentation
- Follow the established patterns
- When in doubt, ask the user
- Keep changes minimal and focused
- **Язык**: Всегда отвечать на русском. Планы (implementation_plan.md) и артефакты должны быть на русском языке.
