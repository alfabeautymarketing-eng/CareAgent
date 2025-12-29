// =======================================================================================
// Menu.gs — DYNAMIC MENU BUILDER
// =======================================================================================
// DESCRIPTION:
//   This file creates the main menu for the spreadsheet.
//   Menu configuration is loaded from the Python server based on spreadsheet ID.
//   Each project (MT, SK, SS) gets its own customized menu.
// =======================================================================================

/**
 * @OnlyCurrentDoc
 */

/**
 * Логирует шаг в лист "Логи"
 */
function _logMenuStep_(action, details, status) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;

    var logSheetName = "Логи";
    var sh = ss.getSheetByName(logSheetName);

    if (!sh) {
      // Создаём лист если его нет
      sh = ss.insertSheet(logSheetName);
      var headers = ["🕒 Время", "🏷️ Категория", "💬 Действие", "📝 Детали", "🔘 Статус"];
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
      sh.setFrozenRows(1);
    }

    var timestamp = Utilities.formatDate(new Date(), "Europe/Moscow", "dd.MM.yyyy HH:mm:ss");
    sh.appendRow([timestamp, "MENU", action, details || "", status || "✅ OK"]);
  } catch (e) {
    console.error("_logMenuStep_ error:", e);
  }
}

/**
 * Creates the main Ecosystem menu.
 * Loads configuration from server to build project-specific menu.
 */
function createAgentMenu(options) {
  _logMenuStep_("Начало создания меню", "options: " + JSON.stringify(options || {}), "🔄 В ПРОЦЕССЕ");

  try {
    const ui = SpreadsheetApp.getUi();
    const allowNetwork = !options || options.allowNetwork !== false;

    _logMenuStep_("Загрузка конфигурации", "allowNetwork: " + allowNetwork, "🔄 В ПРОЦЕССЕ");

    // Try to load menu config from server
    const config = getMenuConfig({ useCacheOnFail: true, skipFetch: !allowNetwork });

    if (config) {
      _logMenuStep_("Конфигурация получена", "project: " + config.project + ", groups: " + (config.menus ? config.menus.length : 0), "✅ OK");
      // Build dynamic menu from server config
      buildDynamicMenu(ui, config);
      _logMenuStep_("Меню создано", "Динамическое меню загружено с сервера", "✅ ГОТОВО");
    } else {
      _logMenuStep_("Сервер недоступен", "Используем fallback меню", "⚠️ ВНИМАНИЕ");
      // Fallback to static menu if server unavailable
      buildFallbackMenu(ui);
      _logMenuStep_("Fallback меню создано", "Offline режим", "✅ ГОТОВО");
    }
  } catch (e) {
    _logMenuStep_("ОШИБКА создания меню", e.message, "❌ ОШИБКА");
    Logger.log('Error creating Agent menu: ' + e.toString());
    console.error('Error creating Agent menu: ' + e.toString());

    // Fallback menu on any error
    try {
      buildFallbackMenu(SpreadsheetApp.getUi());
      _logMenuStep_("Fallback меню после ошибки", "Резервное меню создано", "⚠️ ВНИМАНИЕ");
    } catch (e2) {
      _logMenuStep_("КРИТИЧЕСКАЯ ОШИБКА", e2.message, "❌ ОШИБКА");
      console.error('Failed to create fallback menu:', e2);
    }
  }
}

/**
 * Build dynamic menu from server configuration.
 * @param {GoogleAppsScript.Base.Ui} ui - UI service
 * @param {Object} config - Menu configuration from server
 */
function buildDynamicMenu(ui, config) {
  const registry = (config && (config.menus || config.menu_registry)) || null;

  if (registry && registry.length) {
    _logMenuStep_("Построение групп меню", "Групп: " + registry.length, "🔄 В ПРОЦЕССЕ");

    registry.forEach(function(group, index) {
      if (!group || !group.title || !group.items || group.items.length === 0) {
        return;
      }

      const menu = ui.createMenu(group.title);
      addMenuItems(ui, menu, group.items);
      menu.addToUi();

      _logMenuStep_("Группа меню добавлена", (index + 1) + ". " + group.title + " (" + group.items.length + " пунктов)", "✅ OK");
    });

    console.log('Dynamic menu created for project:', config.project, 'groups:', registry.length);
    return;
  }

  // Backward-compatibility: single menu with items
  if (config && config.menu_title && config.items && config.items.length) {
    const menu = ui.createMenu(config.menu_title);
    config.items.forEach(function(item) {
      if (item.separator) {
        menu.addSeparator();
        return;
      }

      menu.addItem(item.label, item.function_name);

      if (item.separator_after) {
        menu.addSeparator();
      }
    });
    menu.addToUi();
    _logMenuStep_("Legacy меню создано", config.menu_title, "✅ OK");
    console.log('Dynamic menu (legacy format) created for project:', config.project);
    return;
  }

  // If config exists but empty, fallback
  buildFallbackMenu(ui);
}

/**
 * Recursively add menu items and submenus.
 */
function addMenuItems(ui, menu, items) {
  items.forEach(function(item) {
    if (item.separator) {
      menu.addSeparator();
      return;
    }

    if (item.submenu && item.items && item.items.length) {
      const sub = ui.createMenu(item.submenu);
      addMenuItems(ui, sub, item.items);
      menu.addSubMenu(sub);
      if (item.separator_after) {
        menu.addSeparator();
      }
      return;
    }

    if (item.label && item.function_name) {
      menu.addItem(item.label, item.function_name);
    }

    if (item.separator_after) {
      menu.addSeparator();
    }
  });
}

/**
 * Build fallback static menu when server is unavailable.
 * @param {GoogleAppsScript.Base.Ui} ui - UI service
 */
function buildFallbackMenu(ui) {
  const menu = ui.createMenu('⚠️ Ecosystem (Offline)');

  menu.addItem('🔄 Обновить меню', 'refreshMenu')
      .addItem('🔴 Проверить статус сервера', 'checkServerStatus')
      .addSeparator()
      .addItem('🧪 Тест AI (быстрый анализ)', 'menuSimpleAnalyze')
      .addItem('🤖 Анализ строки (полный)', 'menuAnalyzeSelected')
      .addSeparator()
      .addItem('🏭 Сортировать по производителю', 'sortByManufacturer')
      .addItem('💰 Сортировать по цене', 'sortByPrice')
      .addSeparator()
      .addItem('🔄 Обновить данные', 'callServerLoadFunctions')
      .addItem('📑 Упорядочить листы', 'reorderSheets')
      .addSeparator()
      .addItem('📋 Архивировать логи', 'manualArchiveLogs_proxy')
      .addItem('📊 Статус логов', 'showArchiveStatus_proxy')
      .addSeparator()
      .addItem('🐛 Debug: Show Spreadsheet ID', 'debugShowSpreadsheetId')
      .addToUi();

  console.log('Fallback menu created (server unavailable)');
}

/**
 * Force refresh menu from server.
 * Call this manually if menu shows "Offline" but server is running.
 */
function refreshMenu() {
  _logMenuStep_("Обновление меню", "Пользователь запросил обновление", "🔄 В ПРОЦЕССЕ");
  createAgentMenu();
  SpreadsheetApp.getActiveSpreadsheet().toast('Меню обновлено!', 'Ecosystem', 2);
  _logMenuStep_("Обновление меню завершено", "Меню успешно обновлено", "✅ ГОТОВО");
}

// =======================================================================================
// LEGACY MENU FUNCTIONS (for backwards compatibility)
// =======================================================================================

/**
 * @deprecated Use sortByManufacturer() from Client.gs instead
 */
function callSortManufacturer() {
  if (typeof sortByManufacturer === 'function') {
    sortByManufacturer();
  } else {
    SpreadsheetApp.getUi().alert('Function sortByManufacturer not found');
  }
}

/**
 * @deprecated Use sortByPrice() from Client.gs instead
 */
function callSortPrice() {
  if (typeof sortByPrice === 'function') {
    sortByPrice();
  } else {
    SpreadsheetApp.getUi().alert('Function sortByPrice not found');
  }
}
