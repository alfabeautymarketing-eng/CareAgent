"""
Unified menu configuration for all projects.

Menu is served ONLY from server - no local GAS menu creation.
Single source of truth for all menu structures across MT, SS, SK projects.
"""

from typing import Dict, List, Any

# ============================================================================
# UNIFIED PROJECT MENUS CONFIGURATION
# ============================================================================
# All menus are served from server via /api/v1/menu/config endpoint
# GAS loads this dynamically on document open

PROJECT_MENUS: Dict[str, Dict[str, Any]] = {
    # Default/fallback configuration (applies to all projects unless overridden)
    "default": {
        "sort_columns": {
            "manufacturer": "Производитель",
            "price": "Цена",
        },
        "order_sheet": "Заказ",
        # ========================================================================
        # MENU GROUPS - same structure for ALL projects
        # ========================================================================
        "menu_groups": [
            # ============== ГРУППА 1: ЗАКАЗ ==============
            # Единые кнопки, сервер определяет проект автоматически
            {
                "title": "🧾 Заказ",
                "items": [
                    {"label": "📥 Обработка", "function_name": "serverProcessPrimaryData"},
                    {"label": "📊 Загрузить остатки", "function_name": "serverLoadStockData"},
                ],
            },
            # ============== ГРУППА 2: СОРТИРОВКА ==============
            {
                "title": "🔄 Сортировка",
                "items": [
                    {"label": "Сортировать по производителю", "function_name": "sortByManufacturer"},
                    {"label": "Сортировать по прайсу", "function_name": "sortByPrice"},
                    {"separator": True},
                    {"label": "1. Все данные", "function_name": "serverShowAllOrderData"},
                    {"label": "2. Заказ", "function_name": "serverShowOrderStage"},
                    {"label": "3. Акции", "function_name": "serverShowPromotionsStage"},
                    {"label": "4. Набор", "function_name": "serverShowSetStage"},
                    {"label": "5. Прайс", "function_name": "serverShowPriceStage"},
                ],
            },
            # ============== ГРУППА 3: ПРАЙС-ЛИСТ ==============
            {
                "title": "🏷️ Прайс-лист",
                "items": [
                    {"label": "💰 Анализировать цены", "function_name": "menuAnalyzePrices"},
                    {"label": "✅ Сформировать свежий прайс", "function_name": "serverGenerateFreshPriceList"},
                    {"separator": True},
                    {"label": "📅 New год для динамика", "function_name": "serverAddNewYearColumns"},
                ],
            },
            # ============== ГРУППА 4: ЗАКАЗ & ДОКУМЕНТАЦИЯ ==============
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
            # ============== ГРУППА 5: ЭКСПОРТ & ВЫГРУЗКА ==============
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
            # ============== ГРУППА 6: AI АГЕНТ ==============
            {
                "title": "🤖 AI Агент",
                "items": [
                    {"label": "🔑 Ввести API ключ Gemini", "function_name": "setupGeminiComplete"},
                    {"label": "📋 Показать настройки AI", "function_name": "showGeminiSettings"},
                ],
            },
            # ============== ГРУППА 7: НАСТРОЙКА ==============
            {
                "title": "⚙️ Настройка",
                "items": [
                    {"label": "⚙️ Настроить правила синхронизации", "function_name": "showSyncRulesManagerDialog"},
                    {"label": "🔄 Установить триггеры", "function_name": "setupTriggers_proxy"},
                    {"separator": True},
                    {"label": "📜 Открыть Журнал (UI)", "function_name": "openLogDashboard_proxy"},
                    {"label": "🔍 Проверить статус сервера", "function_name": "showAllServicesStatus_proxy"},
                ],
            },
            # ============== ГРУППА 8: РАЗРАБОТКА ==============
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
    # PROJECT-SPECIFIC CONFIGURATION (only metadata, NOT menu items)
    # ========================================================================
    "MT": {
        "menu_title": "MT CosmeticaBar",
        "order_sheet": "Заказ",
        "sort_columns": {
            "manufacturer": "Производитель",
            "price": "EXW ALFASPA текущая, €",
        },
        "primary_menu": {
            "title": "🧾 Заказ",
            "items": {
                "MAIN": "📥 Обработка Main",
                "TESTER": "🧪 Обработка Tester",
                "SAMPLES": "🎁 Обработка Samples",
                "STOCKS": "📊 Загрузить остатки",
                "NEW_PRICE_YEAR": "📅 New год для динамика",
            }
        },
        "order_stages_menu": {
            "title": "🔄 Сортировка",
            "items": {
                "SORT_MANUFACTURER": "Сортировать по производителю",
                "SORT_PRICE": "Сортировать по прайсу",
                "STAGE_ALL": "1. Все данные",
                "STAGE_ORDER": "2. Заказ",
                "STAGE_PROMOTIONS": "3. Акции",
                "STAGE_SET": "4. Набор",
                "STAGE_PRICE": "5. Прайс",
            }
        }
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
                "MAIN": "📥 Обработка Б/З поставщик",
                "PROBES": "🧪 Обработка пробники",
                "STOCKS": "📊 Загрузить остатки",
                "NEW_PRICE_YEAR": "📅 New год для динамика",
            }
        },
        "order_stages_menu": {
            "title": "🔄 Сортировка",
            "items": {
                "SORT_MANUFACTURER": "Сортировать по производителю",
                "SORT_PRICE": "Сортировать по прайсу",
                "STAGE_ALL": "1. Все данные",
                "STAGE_ORDER": "2. Заказ",
                "STAGE_PROMOTIONS": "3. Акции",
                "STAGE_SET": "4. Набор",
                "STAGE_PRICE": "5. Прайс",
            }
        }
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
                "MAIN": "📥 Обработка Б/З поставщик",
                "STOCKS": "📊 Загрузить остатки",
                "NEW_PRICE_YEAR": "📅 New год для динамика",
            }
        },
        "order_stages_menu": {
            "title": "🔄 Сортировка",
            "items": {
                "SORT_MANUFACTURER": "Сортировать по производителю",
                "SORT_PRICE": "Сортировать по прайсу",
                "STAGE_ALL": "1. Все данные",
                "STAGE_ORDER": "2. Заказ",
                "STAGE_PROMOTIONS": "3. Акции",
                "STAGE_SET": "4. Набор",
                "STAGE_PRICE": "5. Прайс",
            }
        }
    },
}


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

    # Merge project-specific config (metadata only)
    project_config = PROJECT_MENUS[project]
    for key, value in project_config.items():
        config[key] = value

    return config


def get_menu_groups(project: str) -> List[Dict[str, Any]]:
    """Get menu groups (same structure for all projects)."""
    return PROJECT_MENUS["default"]["menu_groups"]


def get_menu_title(project: str) -> str:
    """Get project menu title."""
    if project not in PROJECT_MENUS:
        raise ValueError(f"Unknown project: {project}")
    return PROJECT_MENUS[project].get("menu_title", project)


# ============================================================================
# EXPORT FOR BACKWARDS COMPATIBILITY
# ============================================================================

BASE_MENU_GROUPS = PROJECT_MENUS["default"]["menu_groups"]
MENU_CONFIGS = {
    project: {
        k: v
        for k, v in PROJECT_MENUS[project].items()
        if k not in ["menu_groups"]
    }
    for project in ["MT", "SS", "SK"]
}

# Legacy constants for endpoints.py compatibility - NOW UPDATED FOR FULL FUNCTIONALITY
PRIMARY_DATA_MENU_ORDER = ["MAIN", "TESTER", "SAMPLES", "PROBES", "STOCKS", "NEW_PRICE_YEAR"]

PRIMARY_DATA_MENU_ACTIONS = {
    "MAIN": {
        "fn_by_project": {
            "SK": "processSkPriceSheet",
            "SS": "processSsPriceSheet",
            "MT": "processMtMainPrice"
        }
    },
    "TESTER": {
        "fn_by_project": {
            "MT": "processMtTesterPrice"
        }
    },
    "SAMPLES": {
        "fn_by_project": {
            "MT": "processMtSamplesPrice"
        }
    },
    "PROBES": {
        "fn_by_project": {
            "SK": "processSkPriceProbes"
        }
    },
    "STOCKS": {
        "fn_by_project": {
            "SK": "loadSkStockData",
            "SS": "loadSsStockData",
            "MT": "loadMtStockData"
        }
    },
    "NEW_PRICE_YEAR": {
        "fn": "serverAddNewYearColumns"
    },
    # Order Stages Actions
    "SORT_MANUFACTURER": {
        "fn_by_project": {
            "SK": "sortSkOrderByManufacturer",
            "SS": "sortSsOrderByManufacturer",
            "MT": "sortMtOrderByManufacturer"
        }
    },
    "SORT_PRICE": {
        "fn_by_project": {
            "SK": "sortSkOrderByPrice",
            "SS": "sortSsOrderByPrice",
            "MT": "sortMtOrderByPrice"
        }
    },
    "STAGE_ALL": {"fn": "serverShowAllOrderData"},
    "STAGE_ORDER": {"fn": "serverShowOrderStage"},
    "STAGE_PROMOTIONS": {"fn": "serverShowPromotionsStage"},
    "STAGE_SET": {"fn": "serverShowSetStage"},
    "STAGE_PRICE": {"fn": "serverShowPriceStage"},
}

ORDER_STAGES_MENU_ORDER = ["SORT_MANUFACTURER", "SORT_PRICE", "STAGE_ALL", "STAGE_ORDER", "STAGE_PROMOTIONS", "STAGE_SET", "STAGE_PRICE"]


def get_primary_menu(project: str) -> List[Dict[str, Any]]:
    """
    Legacy helper for endpoints.py.
    This is not really used by new endpoints logic which builds properly,
    but kept for potential backward compatibility or other consumers.
    """
    return []


def get_order_stages_menu(project: str) -> List[Dict[str, Any]]:
    """Legacy helper for endpoints.py"""
    return []


def get_menu_action(project: str, group_idx: int, item_idx: int) -> Dict[str, Any]:
    """Get single menu item action."""
    config = get_project_menu(project)
    try:
        return config["menu_groups"][group_idx]["items"][item_idx]
    except (IndexError, KeyError):
        return {"label": "Unknown", "function_name": "unknown"}
