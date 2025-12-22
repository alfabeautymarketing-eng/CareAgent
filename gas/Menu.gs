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
function createAgentMenu() {
  try {
    const ui = SpreadsheetApp.getUi();

    // Try to load menu config from server
    const config = getMenuConfig({ useCacheOnFail: true });

    if (config) {
      // Build dynamic menu from server config
      buildDynamicMenu(ui, config);
    } else {
      // Fallback to static menu if server unavailable
      buildFallbackMenu(ui);
    }
  } catch (e) {
    Logger.log('Error creating Agent menu: ' + e.toString());
    console.error('Error creating Agent menu: ' + e.toString());

    // Fallback menu on any error
    try {
      buildFallbackMenu(SpreadsheetApp.getUi());
    } catch (e2) {
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
    registry.forEach(function(group) {
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

  menu.addItem('🔄 Обновить меню', 'refreshMenu')
      .addItem('🔴 Проверить статус сервера', 'checkServerStatus')
      .addSeparator()
      .addItem('🏭 Сортировать по производителю', 'sortByManufacturer')
      .addItem('💰 Сортировать по цене', 'sortByPrice')
      .addSeparator()
      .addItem('🔄 Обновить данные', 'callServerLoadFunctions')
      .addItem('📑 Упорядочить листы', 'reorderSheets')
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
  createAgentMenu();
  SpreadsheetApp.getActiveSpreadsheet().toast('Меню обновлено!', 'Ecosystem', 2);
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
