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
 * Creates the main Ecosystem menu.
 * Loads configuration from server to build project-specific menu.
 */
function createAgentMenu(options) {
  try {
    let ui;
    try {
      ui = SpreadsheetApp.getUi();
    } catch (err) {
      console.warn('Cannot access UI (likely running from editor): ' + err.toString());
      return; // Return gracefully, we don't need UI in editor
    }

    const allowNetwork = !options || options.allowNetwork !== false;

    // Try to load menu config from server
    const config = getMenuConfig({ useCacheOnFail: true, skipFetch: !allowNetwork });

    if (config) {
      // Build dynamic menu from server config
      buildDynamicMenu(ui, config);
      console.log('Menu created for project:', config.project);
    } else {
      // Fallback to static menu if server unavailable
      buildFallbackMenu(ui);
      console.log('Fallback menu created (server unavailable)');
    }
  } catch (e) {
    console.error('Error creating Agent menu: ' + e.toString());

    // Fallback menu on any error
    try {
      const ui = SpreadsheetApp.getUi();
      buildFallbackMenu(ui);
    } catch (e2) {
      console.warn('Failed to create fallback menu (UI unavailable):', e2);
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
    registry.forEach(function(group, index) {
      if (!group || !group.title || !group.items || group.items.length === 0) {
        return;
      }

      const menu = ui.createMenu(group.title);
      addMenuItems(ui, menu, group.items);
      menu.addToUi();
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

  // Submenu 1: Synchronization
  const syncSub = ui.createMenu('🔄 Синхронизация')
      .addItem('🔄 Синхронизировать всё', 'runFullSync')
      .addItem('🔄 Синхронизировать строку', 'syncSelectedRow')
      .addSeparator()
      .addItem('📝 Правила синхронизации', 'showSyncRulesManagerDialog');

  // Submenu 2: Server Settings
  const serverSub = ui.createMenu('🛠️ Настройки сервера')
      .addItem('🔧 Обновить триггеры', 'setupTriggers')
      .addItem('📋 Журнал синхронизации', 'showSyncLogDialog')
      .addSeparator()
      .addItem('🔴 Проверить статус сервера', 'checkServerStatus')
      .addItem('🔗 Установить URL (ngrok)', 'serverSetLocalTunnel')
      .addItem('🚀 Reset to Production', 'serverResetToProduction');

  // Submenu 3: Ecosystem
  const ecosystemSub = ui.createMenu('🟢 Ecosystem')
      .addItem('🏠 Главная страница', 'openServerMainPage')
      .addItem('📑 Упорядочить листы', 'reorderSheets')
      .addSeparator()
      .addItem('🧪 Тест AI (быстрый)', 'menuSimpleAnalyze');

  // Submenu 4: Supabase
  const supabaseSub = ui.createMenu('🗄️ Supabase')
      .addItem('🔗 Открыть Console', 'openSupabaseConsole')
      .addItem('📊 Просмотр данных', 'showSupabaseTablesView');

  menu.addSubMenu(syncSub)
      .addSubMenu(serverSub)
      .addSubMenu(ecosystemSub)
      .addSubMenu(supabaseSub)
      .addSeparator()
      .addItem('🔄 Обновить меню', 'refreshMenu')
      .addToUi();

  console.log('Fallback menu created (server unavailable)');
}

/**
 * Clear menu config cache.
 * Call this if menu is not updating after server changes.
 */
function clearMenuCache() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = ss.getId();
    const cacheKey = 'MENU_CONFIG_CACHE_' + spreadsheetId;
    const props = PropertiesService.getUserProperties();
    props.deleteProperty(cacheKey);
    console.log('Menu cache cleared for spreadsheet:', spreadsheetId);
    
    try {
      ss.toast('Кэш меню очищен!', 'Cache', 2);
    } catch (toastErr) {
      // Ignore toast error in editor
    }
  } catch (e) {
    console.error('Error clearing menu cache:', e);
    try {
      SpreadsheetApp.getUi().alert('Ошибка: ' + e.toString());
    } catch (uiErr) {}
  }
}

/**
 * Force refresh menu from server.
 * Call this manually if menu shows "Offline" but server is running.
 * This now clears the cache before reloading.
 */
function refreshMenu() {
  clearMenuCache();
  createAgentMenu();
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Меню обновлено!', 'Ecosystem', 2);
  } catch (e) {}
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
