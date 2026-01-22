"""
Unified menu configuration for all projects.

Combines static menu groups (BASE_MENU_GROUPS) with project-specific items (MENU_CONFIGS).
Single source of truth for all menu structures across MT, SS, SK projects.
"""

from typing import Dict, List, Any

# ============================================================================
# UNIFIED PROJECT MENUS CONFIGURATION
# ============================================================================
# Merges BASE_MENU_GROUPS (static) + MENU_CONFIGS (project-specific)

PROJECT_MENUS: Dict[str, Dict[str, Any]] = {
    # Default/fallback configuration (applies to all projects unless overridden)
    "default": {
        "sort_columns": {
            "manufacturer": "Производитель",
            "price": "Цена",
        },
        "order_sheet": "Заказ",
        # Static menu groups (same for all projects)
        "menu_groups": [
            # ============== ЦИКЛ 1: ПРАЙС-ЛИСТ ==============
            {
                "title": "🏷️ Прайс-лист",
                "items": [
                    {"label": "📥 Загрузить и обработать прайс", "function_name": "serverProcessPrimaryData"},
                    {"label": "📊 Загрузить остатки", "function_name": "serverLoadStockData"},
                    {"separator": True},
                    {"label": "💰 Анализировать цены", "function_name": "menuAnalyzePrices"},
                    {"label": "✅ Сформировать свежий прайс", "function_name": "serverGenerateFreshPriceList"},
                ],
            },
            # ============== ЦИКЛ 2: ЗАКАЗ & ДОКУМЕНТАЦИЯ ==============
            {
                "title": "🛒 Заказ & Документация",
                "items": [
                    {"label": "📋 Подготовить заказ", "function_name": "serverPrepareOrder"},
                    {"separator": True},
                    {"label": "📦 Подготовить документы для таможни", "function_name": "serverPrepareCustomsDocuments"},
                    {"label": "✏️ Заполнить разрешительную документацию", "function_name": "serverFillPermitDocumentation"},
                    {"separator": True},
                    {"label": "🔬 Сертификация", "function_name": "serverCertificationWorkbench"},
                    {"label": "📄 Лист новинки", "function_name": "serverCreateNewsSheet"},
                    {"label": "📋 Протоколы (353пп)", "function_name": "serverGenerateProtocols353pp"},
                    {"label": "📋 ДС Макеты (353пп)", "function_name": "serverGenerateDsLayouts"},
                    {"label": "📦 Собрать документы для заявки", "function_name": "structureDocuments_353pp"},
                    {"separator": True},
                    {"label": "🥃 Спирты - Посчитать", "function_name": "serverCalculateSpiritNumbers"},
                    {"label": "🥃 Спирты - Создать Макеты", "function_name": "generateSpiritProtocols"},
                    {"label": "🔄 Спирты - Пересчитать каскады", "function_name": "serverRecalculateCascades"},
                ],
            },
            # ============== ЦИКЛ 3: ЭКСПОРТ & ВЫГРУЗКА ==============
            {
                "title": "📦 Экспорт & Выгрузка",
                "items": [
                    {"label": "📤 Выгрузить Акции", "function_name": "serverExportPromotions"},
                    {"label": "📤 Выгрузить Наборы", "function_name": "serverExportSets"},
                    {"separator": True},
                    {"label": "📄 Форматировать лист 'Ордер'", "function_name": "serverFormatOrderSheet"},
                    {"label": "📄 Создать лист 'Для инвойса'", "function_name": "serverCreateFullInvoice"},
                    {"label": "📄 Собрать документы для инвойса", "function_name": "collectAndCopyDocuments"},
                ],
            },
            # ============== ЦИКЛ 4: AI АГЕНТ ==============
            {
                "title": "🤖 AI Агент",
                "items": [
                    {"label": "🔑 Ввести API ключ Gemini", "function_name": "setupGeminiComplete"},
                    {"label": "📋 Показать настройки AI", "function_name": "showGeminiSettings"},
                ],
            },
            # ============== ЦИКЛ 5: НАСТРОЙКА ==============
            {
                "title": "⚙️ Настройка",
                "items": [
                    {"label": "🔄 Синхронизировать конфиги", "function_name": "serverSyncConfigs"},
                    {"label": "📊 Показать логи", "function_name": "serverShowLogs"},
                ],
            },
            # ============== ЦИКЛ 6: РАЗРАБОТКА ==============
            {
                "title": "🛠️ Инструменты разработчика",
                "items": [
                    {"label": "🔗 Установить URL туннеля (ngrok)", "function_name": "serverSetLocalTunnel"},
                    {"label": "🚀 Вернуть на Production (VPS)", "function_name": "serverResetToProduction"},
                    {"separator": True},
                    {"label": "ℹ️ Показать текущий SERVER_URL", "function_name": "serverShowCurrentUrl"},
                    {"label": "🔄 Принудительно обновить меню", "function_name": "refreshMenu"},
                ],
            },
        ],
    },
    # ========================================================================
    # PROJECT-SPECIFIC OVERRIDES
    # ========================================================================
    "MT": {
        "menu_title": "MT CosmeticaBar",
        "order_sheet": "Заказ",
        "sort_columns": {
            "manufacturer": "Производитель",
            "price": "EXW ALFASPA текущая, €",
        },
        # Project-specific submenu: 🧾 Заказ (part of primary_menu in MENU_CONFIGS)
        "primary_menu": {
            "title": "🧾 Заказ",
            "items": {
                "MAIN": "Обработка Б/З поставщик",
                "TESTER": "Обработка Тестер",
                "SAMPLES": "Обработка Пробники",
                "STOCKS": "Загрузить остатки",
                "NEW_PRICE_YEAR": "New год для динамика",
            },
        },
        # Project-specific submenu: 📊 Стадии по заказ
        "order_stages_menu": {
            "title": "📊 Стадии по заказ",
            "items": {
                "SORT_MANUFACTURER": "Сортировать по производителю",
                "SORT_PRICE": "Сортировать по прайсу",
                "STAGE_ALL": "1. Все данные",
                "STAGE_ORDER": "2. Заказ",
                "STAGE_PROMOTIONS": "3. Акции",
                "STAGE_SET": "4. Набор",
                "STAGE_PRICE": "5. Прайс",
            },
        },
    },
    "SK": {
        "menu_title": "SK Carmado",
        "order_sheet": "Заказ",
        "sort_columns": {
            "manufacturer": "Производитель",
            "price": "Цена",
        },
        "primary_menu": {
            "title": "🧾 Заказ",
            "items": {
                "MAIN": "Обработка Б/З поставщик",
                "PROBES": "Обработка пробники",
                "STOCKS": "Загрузить остатки",
                "NEW_PRICE_YEAR": "New год для динамика",
            },
        },
        "order_stages_menu": {
            "title": "📊 Стадии по заказ",
            "items": {
                "SORT_MANUFACTURER": "Сортировать по производителю",
                "SORT_PRICE": "Сортировать по прайсу",
                "STAGE_ALL": "1. Все данные",
                "STAGE_ORDER": "2. Заказ",
                "STAGE_PROMOTIONS": "3. Акции",
                "STAGE_SET": "4. Набор",
                "STAGE_PRICE": "5. Прайс",
            },
        },
    },
    "SS": {
        "menu_title": "SS San",
        "order_sheet": "Заказ",
        "sort_columns": {
            "manufacturer": "Производитель",
            "price": "Цена",
        },
        "primary_menu": {
            "title": "🧾 Заказ",
            "items": {
                "MAIN": "Обработка Б/З поставщик",
                "STOCKS": "Загрузить остатки",
                "NEW_PRICE_YEAR": "New год для динамика",
            },
        },
        "order_stages_menu": {
            "title": "📊 Стадии по заказ",
            "items": {
                "SORT_MANUFACTURER": "Сортировать по производителю",
                "SORT_PRICE": "Сортировать по прайсу",
                "STAGE_ALL": "1. Все данные",
                "STAGE_ORDER": "2. Заказ",
                "STAGE_PROMOTIONS": "3. Акции",
                "STAGE_SET": "4. Набор",
                "STAGE_PRICE": "5. Прайс",
            },
        },
    },
}

# ============================================================================
# ACTION MAPPING (Function References)
# ============================================================================
# Maps menu item keys to their server function implementations
# Replaces PRIMARY_DATA_MENU_ACTIONS from endpoints.py

PRIMARY_DATA_MENU_ACTIONS = {
    # Main processing functions (MAIN mode)
    "MAIN": {
        "fn_by_project": {
            "SK": "serverProcessSkMain",
            "SS": "serverProcessSsMain",
            "MT": "serverProcessMtMain",
        }
    },
    # Processing variants (TESTER, SAMPLES, PROBES) - project-specific
    "TESTER": {"fn_by_project": {"MT": "serverProcessMtTester"}},
    "SAMPLES": {"fn_by_project": {"MT": "serverProcessMtSamples"}},
    "PROBES": {"fn_by_project": {"SK": "serverProcessSkProbes"}},
    # Stock loading (project-specific)
    "STOCKS": {
        "fn_by_project": {
            "SK": "loadSkStockData",
            "SS": "loadSsStockData",
            "MT": "loadMtStockData",
        }
    },
    # Price/Date operations (generic)
    "NEW_PRICE_YEAR": {"fn": "serverAddNewYearColumns"},
    # Sorting operations (generic)
    "SORT_MANUFACTURER": {"fn": "sortByManufacturer"},
    "SORT_PRICE": {"fn": "sortByPrice"},
    # Stage view operations (generic)
    "STAGE_ALL": {"fn": "serverShowAllOrderData"},
    "STAGE_ORDER": {"fn": "serverShowOrderStage"},
    "STAGE_PROMOTIONS": {"fn": "serverShowPromotionsStage"},
    "STAGE_SET": {"fn": "serverShowSetStage"},
    "STAGE_PRICE": {"fn": "serverShowPriceStage"},
}

# ============================================================================
# MENU ITEM ORDERING
# ============================================================================
# Defines the order in which menu items appear

PRIMARY_DATA_MENU_ORDER = ["MAIN", "TESTER", "SAMPLES", "PROBES", "STOCKS", "NEW_PRICE_YEAR"]
ORDER_STAGES_MENU_ORDER = [
    "SORT_MANUFACTURER",
    "SORT_PRICE",
    "STAGE_ALL",
    "STAGE_ORDER",
    "STAGE_PROMOTIONS",
    "STAGE_SET",
    "STAGE_PRICE",
]


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================


def get_project_menu(project: str) -> Dict[str, Any]:
    """
    Get complete menu configuration for a project.

    Args:
        project: Project code (MT, SS, SK)

    Returns:
        Merged configuration with project-specific overrides
    """
    if project not in PROJECT_MENUS:
        raise ValueError(f"Unknown project: {project}")

    # Start with default config
    config = PROJECT_MENUS["default"].copy()

    # Merge project-specific config
    project_config = PROJECT_MENUS[project]
    config.update(project_config)

    return config


def get_menu_groups(project: str) -> List[Dict[str, Any]]:
    """Get static menu groups for a project (same for all projects)."""
    return PROJECT_MENUS["default"]["menu_groups"]


def get_primary_menu(project: str) -> Dict[str, Any]:
    """Get project-specific primary menu (🧾 Заказ submenu)."""
    if project not in PROJECT_MENUS:
        raise ValueError(f"Unknown project: {project}")
    return PROJECT_MENUS[project].get("primary_menu", {})


def get_order_stages_menu(project: str) -> Dict[str, Any]:
    """Get project-specific order stages menu (📊 Стадии по заказ submenu)."""
    if project not in PROJECT_MENUS:
        raise ValueError(f"Unknown project: {project}")
    return PROJECT_MENUS[project].get("order_stages_menu", {})


def get_menu_title(project: str) -> str:
    """Get project menu title."""
    if project not in PROJECT_MENUS:
        raise ValueError(f"Unknown project: {project}")
    return PROJECT_MENUS[project].get("menu_title", project)


def get_menu_action(menu_key: str, project: str = None) -> Dict[str, Any]:
    """
    Get function action for a menu item.

    Args:
        menu_key: Menu item key (e.g., "MAIN", "STOCKS")
        project: Project code (optional, required for project-specific functions)

    Returns:
        Action configuration with function name(s)
    """
    if menu_key not in PRIMARY_DATA_MENU_ACTIONS:
        raise ValueError(f"Unknown menu action: {menu_key}")

    action = PRIMARY_DATA_MENU_ACTIONS[menu_key]

    # If it has project-specific functions, get the one for this project
    if "fn_by_project" in action and project:
        fn = action["fn_by_project"].get(project)
        if not fn:
            raise ValueError(f"No function for {menu_key} on project {project}")
        return {"fn": fn}

    # Otherwise return the generic function
    return action


# ============================================================================
# EXPORT FOR BACKWARDS COMPATIBILITY
# ============================================================================
# Keep these exports for code that still references the old names

BASE_MENU_GROUPS = PROJECT_MENUS["default"]["menu_groups"]
MENU_CONFIGS = {
    project: {
        k: v
        for k, v in PROJECT_MENUS[project].items()
        if k not in ["menu_groups"]
    }
    for project in ["MT", "SS", "SK"]
}
