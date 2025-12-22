// =======================================================================================
// Client.gs — HTTP CLIENT FOR PYTHON SERVER
// =======================================================================================
// DESCRIPTION:
//   This file handles all communication with the Python backend server.
//   It provides functions for menu configuration, sorting, and data operations.
// =======================================================================================

const SERVER_URL = 'http://46.226.167.153:8000';

// Cache for sort config (avoid repeated API calls)
var _sortConfigCache = null;

// =======================================================================================
// MENU CONFIGURATION
// =======================================================================================

/**
 * Build cache key for menu config.
 */
function getMenuConfigCacheKey(spreadsheetId) {
  return 'MENU_CONFIG_CACHE_' + spreadsheetId;
}

/**
 * Save menu config to user properties cache.
 */
function saveMenuConfigToCache(spreadsheetId, config) {
  try {
    const props = PropertiesService.getUserProperties();
    props.setProperty(getMenuConfigCacheKey(spreadsheetId), JSON.stringify(config));
  } catch (e) {
    console.error('Failed to cache menu config:', e);
  }
}

/**
 * Load cached menu config (if any).
 */
function getCachedMenuConfig(spreadsheetId) {
  try {
    const props = PropertiesService.getUserProperties();
    const raw = props.getProperty(getMenuConfigCacheKey(spreadsheetId));
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    console.error('Failed to read cached menu config:', e);
    return null;
  }
}

/**
 * Get menu configuration from server for current spreadsheet.
 * Falls back to cached config to avoid "Offline" menu when server briefly недоступен.
 * @param {{useCacheOnFail?: boolean}} options
 * @returns {Object|null} Menu config or null on error
 */
function getMenuConfig(options) {
  const useCacheOnFail = !options || options.useCacheOnFail !== false;
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const url = SERVER_URL + '/api/v1/menu/config?spreadsheet_id=' + encodeURIComponent(spreadsheetId);

  console.log('Fetching menu config from:', url);
  console.log('Spreadsheet ID:', spreadsheetId);

  try {
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      timeout: 10 // 10 seconds timeout
    });

    const code = response.getResponseCode();
    const text = response.getContentText();

    console.log('Menu config response code:', code);

    if (code === 200) {
      const config = JSON.parse(text);
      console.log('Menu config loaded for project:', config.project);
      saveMenuConfigToCache(spreadsheetId, config);
      return config;
    } else {
      console.error('Menu config error (code ' + code + '):', text);
    }
  } catch (e) {
    console.error('Menu config fetch error:', e.toString());
    console.error('Server URL:', SERVER_URL);
  }

  if (useCacheOnFail) {
    const cached = getCachedMenuConfig(spreadsheetId);
    if (cached) {
      console.log('Using cached menu config for spreadsheet:', spreadsheetId);
      return cached;
    }
  }

  return null;
}

/**
 * Get sort configuration for current spreadsheet.
 * Cached to avoid repeated API calls during menu operations.
 * @param {boolean} forceRefresh - Force cache refresh
 * @returns {Object|null} Sort config or null on error
 */
function getSortConfig(forceRefresh) {
  // By default always refresh from server to avoid stale headers.
  const useCache = _sortConfigCache && forceRefresh === false;
  if (useCache) {
    return _sortConfigCache;
  }

  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

  try {
    const response = UrlFetchApp.fetch(
      SERVER_URL + '/api/v1/menu/sort-config?spreadsheet_id=' + encodeURIComponent(spreadsheetId),
      { muteHttpExceptions: true }
    );

    if (response.getResponseCode() === 200) {
      _sortConfigCache = JSON.parse(response.getContentText());
      return _sortConfigCache;
    } else {
      console.error('Sort config error:', response.getContentText());
      return null;
    }
  } catch (e) {
    console.error('Sort config fetch error:', e);
    return null;
  }
}

// =======================================================================================
// SORT FUNCTIONS (Project-aware)
// =======================================================================================

/**
 * Sort by Manufacturer column using project-specific configuration.
 */
function sortByManufacturer() {
  const config = getSortConfig();

  if (!config) {
    SpreadsheetApp.getUi().alert('❌ Не удалось получить конфигурацию сортировки с сервера.');
    return;
  }

  const columnName = config.sort_columns.manufacturer;
  const sheetName = config.order_sheet;

  callServerSortSheet(sheetName, columnName, true);
}

/**
 * Sort by Price column using project-specific configuration.
 */
function sortByPrice() {
  const config = getSortConfig();

  if (!config) {
    SpreadsheetApp.getUi().alert('❌ Не удалось получить конфигурацию сортировки с сервера.');
    return;
  }

  const columnName = config.sort_columns.price;
  const sheetName = config.order_sheet;

  callServerSortSheet(sheetName, columnName, true);
}

/**
 * Sort a specific sheet by column name.
 * @param {string} sheetName - Name of the sheet to sort
 * @param {string} columnName - Column header to sort by
 * @param {boolean} ascending - Sort order
 */
function callServerSortSheet(sheetName, columnName, ascending) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const payload = {
    spreadsheet_id: ss.getId(),
    sheet_name: sheetName,
    column_name: columnName,
    ascending: ascending
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/sort', options);
    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code === 200) {
      ss.toast('✅ Сортировка выполнена: ' + sheetName + ' по ' + columnName);
    } else {
      SpreadsheetApp.getUi().alert('❌ Ошибка сервера (' + code + '):\n' + text);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Ошибка сети:\n' + e.message);
  }
}

/**
 * Generic sort function (for backwards compatibility).
 * Sorts the ACTIVE sheet by column name.
 */
function callServerSort(columnName, ascending) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  callServerSortSheet(sheet.getName(), columnName, ascending);
}

/**
 * High-performance structure sorting via Python server.
 * Replaces the slow GAS structureMultipleSheets function.
 * Sorts 3 sheets (Заказ, Динамика цены, Расчет цены) with group headers.
 * @param {string} mode - 'byManufacturer' or 'byPrice'
 */
function callServerStructureSort(mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const payload = {
    spreadsheet_id: ss.getId(),
    mode: mode
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    ss.toast('⏳ Сортировка запущена...', 'Python Server', 30);
    
    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/sort/structure', options);
    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code === 200) {
      const result = JSON.parse(text);
      const sheets = result.details.sheets_processed.map(function(s) { return s.name; }).join(', ');
      const time = result.details.total_time_ms;
      
      ss.toast(
        '✅ Сортировка завершена за ' + (time / 1000).toFixed(1) + ' сек\n' +
        'Листы: ' + sheets,
        'Python Server',
        5
      );
    } else {
      SpreadsheetApp.getUi().alert('❌ Ошибка сервера (' + code + '):\n' + text);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Ошибка сети:\n' + e.message + '\n\nПроверьте, запущен ли сервер.');
  }
}

// =======================================================================================
// SERVER STATUS
// =======================================================================================

/**
 * Check server health and display detailed status.
 */
function checkServerStatus() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();

  let report = '=== ДИАГНОСТИКА СЕРВЕРА ===\n\n';
  report += 'Spreadsheet ID: ' + spreadsheetId + '\n';
  report += 'Server URL: ' + SERVER_URL + '\n\n';

  // 1. Check /health endpoint
  try {
    const healthResponse = UrlFetchApp.fetch(SERVER_URL + '/health', {
      muteHttpExceptions: true,
      timeout: 10
    });
    const healthCode = healthResponse.getResponseCode();

    if (healthCode === 200) {
      const json = JSON.parse(healthResponse.getContentText());
      report += '✅ /health: OK (версия ' + json.version + ')\n';
    } else {
      report += '❌ /health: Код ' + healthCode + '\n';
    }
  } catch (e) {
    report += '❌ /health: ' + e.message + '\n';
  }

  // 2. Check /api/v1/menu/config endpoint
  try {
    const menuUrl = SERVER_URL + '/api/v1/menu/config?spreadsheet_id=' + encodeURIComponent(spreadsheetId);
    const menuResponse = UrlFetchApp.fetch(menuUrl, {
      muteHttpExceptions: true,
      timeout: 10
    });
    const menuCode = menuResponse.getResponseCode();

    if (menuCode === 200) {
      const config = JSON.parse(menuResponse.getContentText());
      const registry = config.menus || config.menu_registry || [];
      const registryCount = registry.length;
      const topLevelItems = registry.reduce(function(acc, g) {
        return acc + (g.items ? g.items.length : 0);
      }, 0);

      report += '✅ /menu/config: OK (проект ' + config.project + ')\n';
      if (config.menu_title) {
        report += '   Название меню: ' + config.menu_title + '\n';
      }
      if (registryCount > 0) {
        report += '   Групп меню: ' + registryCount + ', верхний уровень: ' + topLevelItems + ' пунктов\n';
      } else if (config.items) {
        report += '   Пунктов меню: ' + config.items.length + '\n';
      }
    } else {
      report += '❌ /menu/config: Код ' + menuCode + '\n';
      report += '   Ответ: ' + menuResponse.getContentText().substring(0, 200) + '\n';
    }
  } catch (e) {
    report += '❌ /menu/config: ' + e.message + '\n';
  }

  // 3. Check /api/v1/sheets/reorder endpoint (just check it responds)
  try {
    const reorderUrl = SERVER_URL + '/api/v1/sheets/reorder';
    const reorderResponse = UrlFetchApp.fetch(reorderUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ spreadsheet_id: spreadsheetId }),
      muteHttpExceptions: true,
      timeout: 15
    });
    const reorderCode = reorderResponse.getResponseCode();

    if (reorderCode === 200) {
      const result = JSON.parse(reorderResponse.getContentText());
      report += '✅ /sheets/reorder: OK (' + result.details.sheets_moved + ' листов перемещено)\n';
    } else {
      report += '⚠️ /sheets/reorder: Код ' + reorderCode + '\n';
    }
  } catch (e) {
    report += '❌ /sheets/reorder: ' + e.message + '\n';
  }

  report += '\n=== РЕКОМЕНДАЦИИ ===\n';
  report += 'Если сервер недоступен:\n';
  report += '1. Проверьте что сервер запущен\n';
  report += '2. Выполните: cd AgentCare && poetry run uvicorn src.main:app --host 0.0.0.0 --port 8000\n';

  ui.alert('Диагностика сервера', report, ui.ButtonSet.OK);
}

// =======================================================================================
// DATA OPERATIONS
// =======================================================================================

/**
 * Run heavy load functions on server (formulas, calculations).
 */
function callServerLoadFunctions() {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const config = getSortConfig();
  const project = config ? config.project : 'Common';

  const payload = {
    spreadsheet_id: spreadsheetId,
    project: project
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/load-functions', options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const result = JSON.parse(responseBody);
      SpreadsheetApp.getActiveSpreadsheet().toast('✅ ' + result.message);
    } else {
      console.error('Server error:', responseBody);
      SpreadsheetApp.getActiveSpreadsheet().toast('❌ Ошибка сервера: ' + responseCode);
    }
  } catch (e) {
    console.error('Fetch error:', e);
    SpreadsheetApp.getActiveSpreadsheet().toast('❌ Ошибка соединения: ' + e.message);
  }
}

/**
 * Send edit event to Python server for synchronization.
 * @param {Object} payload - Event data (sheet, row, col, value, etc.)
 */
function callServerSyncEvent(payload) {
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/sync/event', options);
    const code = response.getResponseCode();

    if (code !== 200) {
      console.error('Sync server error:', response.getContentText());
    } else {
      console.log('Sync success:', response.getContentText());
    }
  } catch (e) {
    console.error('Sync network error:', e);
  }
}

// =======================================================================================
// STRUCTURE SORTING (Multi-sheet sorting with group headers)
// =======================================================================================

/**
 * Structure sort by manufacturer via Python server.
 */
function structureSortByManufacturer() {
  callServerStructureSort('byManufacturer');
}

/**
 * Structure sort by price via Python server.
 */
function structureSortByPrice() {
  callServerStructureSort('byPrice');
}

// =======================================================================================
// SHEET ORDERING
// =======================================================================================

/**
 * Reorder sheets according to standard order via Python server.
 * Called automatically on document open or manually from menu.
 */
function reorderSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const payload = {
    spreadsheet_id: ss.getId()
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/sheets/reorder', options);
    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code === 200) {
      const result = JSON.parse(text);
      if (result.details.sheets_moved > 0) {
        ss.toast('📑 ' + result.message, 'Python Server', 3);
      }
      return true;
    } else {
      console.error('Reorder error:', text);
      return false;
    }
  } catch (e) {
    console.error('Reorder network error:', e);
    return false;
  }
}

/**
 * Silent reorder - no toast, used for automatic onOpen.
 */
function reorderSheetsSilent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const payload = {
    spreadsheet_id: ss.getId()
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    UrlFetchApp.fetch(SERVER_URL + '/api/v1/sheets/reorder', options);
  } catch (e) {
    console.error('Silent reorder failed:', e);
  }
}

// =======================================================================================
// DEBUG FUNCTIONS
// =======================================================================================

/**
 * Show current spreadsheet ID (for debugging).
 */
function debugShowSpreadsheetId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const id = ss.getId();
  const config = getSortConfig();
  const project = config ? config.project : 'Unknown';

  SpreadsheetApp.getUi().alert(
    'Debug Info',
    'Spreadsheet ID:\n' + id + '\n\nПроект: ' + project,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  console.log('DEBUG: Spreadsheet ID:', id, 'Project:', project);
}
