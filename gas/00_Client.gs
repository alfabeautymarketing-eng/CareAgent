// =======================================================================================
// Client.gs — HTTP CLIENT FOR PYTHON SERVER
// =======================================================================================
// DESCRIPTION:
//   This file handles all communication with the Python backend server.
//   It provides functions for menu configuration, sorting, and data operations.
// =======================================================================================

const SERVER_URL = 'http://46.226.167.153:8000';

/**
 * Proxy for API requests from the native GAS UI.
 * Bypasses Mixed Content restrictions by fetching from server-side GAS.
 * @param {string} path - Relative API path (e.g. "rules/123")
 * @param {string} method - HTTP method
 * @param {string} payload - JSON string payload
 */
function apiProxy(path, method, payload) {
  let url = SERVER_URL + '/api/v1' + (path.startsWith('/') ? '' : '/') + path;
  const httpMethod = (method || 'GET').toUpperCase();
  
  const options = {
    method: httpMethod,
    muteHttpExceptions: true,
    contentType: 'application/json'
  };
  
  if (payload) {
    if (httpMethod === 'GET') {
      const qs = Object.entries(payload)
        .filter(([_, v]) => v !== null && v !== undefined)
        .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
        .join('&');
      if (qs) url += (url.includes('?') ? '&' : '?') + qs;
    } else {
      options.payload = typeof payload === 'string' ? payload : JSON.stringify(payload);
    }
  }

  console.log(`[apiProxy] ${httpMethod} ${url}`);
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const content = response.getContentText();
    let json = null;
    try { json = JSON.parse(content); } catch(e) {}
    
    return {
      status: response.getResponseCode(),
      content: content,
      json: json
    };
  } catch (e) {
    console.error(`[apiProxy] Error: ${e.toString()}`);
    return {
      status: 500,
      content: JSON.stringify({ error: e.toString() }),
      json: { error: e.toString() }
    };
  }
}

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
  const skipFetch = options && options.skipFetch === true;
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const url = SERVER_URL + '/api/v1/menu/config?spreadsheet_id=' + encodeURIComponent(spreadsheetId);

  if (skipFetch) {
    const cached = getCachedMenuConfig(spreadsheetId);
    if (cached) {
      console.log('Using cached menu config (network skipped) for spreadsheet:', spreadsheetId);
      return cached;
    }

    if (!useCacheOnFail) {
      return null;
    }

    console.warn('Menu config: network disabled and cache is empty, falling back to offline menu');
    return null;
  }

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
// RULES MANAGEMENT (Server CRUD)
// =======================================================================================

/**
 * Получить правила синхронизации с сервера.
 * @param {boolean} forceReload — обновить кэш на сервере
 */
function fetchServerRules(forceReload) {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const res = Lib.callServer(`/rules/${ssId}`, { force_reload: !!forceReload }, "GET");
  return res && res.rules ? res.rules : [];
}

/**
 * Перезагрузить правила на сервере (invalidate cache).
 */
function reloadServerRules() {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return Lib.callServer(`/rules/${ssId}/reload`, {}, "POST");
}

/**
 * Создать правило на сервере.
 */
function createServerRule(ruleData) {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return Lib.callServer(`/rules/${ssId}/create`, ruleData || {}, "POST");
}

/**
 * Обновить правило на сервере.
 */
function updateServerRule(ruleId, updates) {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return Lib.callServer(`/rules/${ssId}/${ruleId}`, updates || {}, "PATCH");
}

/**
 * Удалить правило на сервере.
 */
function deleteServerRule(ruleId) {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return Lib.callServer(`/rules/${ssId}/${ruleId}`, {}, "DELETE");
}

/**
 * Включить/выключить правило на сервере.
 */
function toggleServerRule(ruleId, enabled) {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return Lib.callServer(`/rules/${ssId}/${ruleId}/toggle`, { enabled: !!enabled }, "PATCH");
}

function deleteLegacyRulesSheet() {
  if (Lib.deleteLegacyRulesSheet) {
    return Lib.deleteLegacyRulesSheet();
  }
  return "Функция Lib.deleteLegacyRulesSheet не найдена.";
}

function migrateLegacyRules() {
  if (Lib.migrateLegacyRules) {
    return Lib.migrateLegacyRules();
  }
  return "Функция Lib.migrateLegacyRules не найдена.";
}

/**
 * Открыть UI управления правилами.
 */
function showSyncRulesManagerDialog() {
  if (Lib.showSyncRulesManagerDialog) {
    return Lib.showSyncRulesManagerDialog();
  }
  throw new Error('Lib.showSyncRulesManagerDialog не определена');
}


/**
 * Список рабочих листов (без служебных).
 */
function getSheetsList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const service = new Set();
  if (Lib.CONFIG && Lib.CONFIG.SHEETS) {
    ["RULES", "LOG", "LOG_DEBUG", "EXTERNAL_DOCS", "SESSION_LOG"].forEach((k) => {
      if (Lib.CONFIG.SHEETS[k]) service.add(Lib.CONFIG.SHEETS[k]);
    });
  }
  return ss
    .getSheets()
    .map((s) => s.getName())
    .filter((n) => !service.has(n));
}

/**
 * Заголовки указанного листа (первая строка).
 */
function getSheetColumns(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || [];
  return headers.map((h) => String(h || "").trim()).filter(Boolean);
}

/**
 * Список внешних документов для диалога правил.
 */
function getExternalDocsList() {
  if (Lib.getExternalDocsList) {
    return Lib.getExternalDocsList();
  }
  return [];
}

/**
 * Список листов во внешнем документе.
 */
function getExternalSheetsList(docId) {
  if (Lib.getExternalSheetsList) {
    return Lib.getExternalSheetsList(docId);
  }
  return [];
}

/**
 * Заголовки листа во внешнем документе.
 */
function getExternalSheetColumns(docId, sheetName) {
  if (Lib.getExternalSheetColumns) {
    return Lib.getExternalSheetColumns(docId, sheetName);
  }
  return [];
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
      try {
        console.log('Reorder completed:', result.details);
      } catch (e) {
        console.error('Reorder post-processing failed:', e);
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
    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/sheets/reorder', options);
    console.log('Silent reorder completed with code:', response.getResponseCode());
  } catch (e) {
    console.error('Silent reorder failed:', e);
  }
}

// =======================================================================================
// LOGS & DOCUMENTS (Python Server)
// =======================================================================================

/**
 * Call server to recreate or clean log sheet.
 * @param {boolean} forceClean - If true, clears the sheet.
 */
// ОТКЛЮЧЕНО: Лист логов в таблице больше не используется.

/**
 * Call server to collect documents.
 */
function callServerCollectDocuments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const payload = {
    spreadsheet_id: ss.getId(),
    target_sheet: "Для инвойса" // Default
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    timeout: 300 // 5 minutes timeout for large copy jobs
  };

  try {
    ss.toast('⏳ Сбор документов начат...', 'Python Server', 60);

    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/documents/collect', options);
    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code === 200) {
      const result = JSON.parse(text);
      const data = result.data;
      
      let msg = '✅ Документы собраны!\n\n';
      msg += 'Файлов скопировано: ' + data.count + '\n';
      msg += 'Ошибок: ' + data.errors + '\n';
      if (data.folder_id) {
        msg += 'ID папки: ' + data.folder_id + '\n';
      }
      
      ui.alert('Сбор документов', msg, ui.ButtonSet.OK);
      return true;
    } else {
      ui.alert('❌ Ошибка сервера (' + code + '):\n' + text);
      return false;
    }
  } catch (e) {
    ui.alert('❌ Ошибка сети:\n' + e.message);
    return false;
  }
}

/**
 * Call server to archive logs manually.
 */
function callServerArchiveLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    "Архивирование логов",
    "Скопировать все текущие логи в месячный архив?\n\n" +
    "Логи будут скопированы в папку архивов, лист 'Логи' будет очищен.",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const payload = {
    spreadsheet_id: ss.getId()
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    timeout: 300 // Long timeout for archive
  };

  try {
    ss.toast('⏳ Архивирование...', 'Python Server', 30);

    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/logs/archive', options);
    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code === 200) {
      const result = JSON.parse(text);
      const data = result.data;
      
      let msg = '✅ Архивирование завершено!\n\n';
      msg += 'Всего строк: ' + data.total_rows + '\n';
      msg += 'Имя архива: ' + data.archive_name + '\n';
      
      ui.alert('Архивирование', msg, ui.ButtonSet.OK);
      return true;
    } else {
      ui.alert('❌ Ошибка сервера (' + code + '):\n' + text);
      return false;
    }
  } catch (e) {
    ui.alert('❌ Ошибка сети:\n' + e.message);
    return false;
  }
}

/**
 * Call server to show archive status.
 */
function callServerShowArchiveStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const url = SERVER_URL + '/api/v1/logs/archive/status?spreadsheet_id=' + encodeURIComponent(ss.getId());

  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code === 200) {
      const result = JSON.parse(text);
      const status = result.data;
      
      let msg = "📊 СТАТУС АРХИВИРОВАНИЯ (Server)\n\n";
      msg += `📁 Текущий архив: ${status.archive_name}\n`;
      msg += `📂 ID Папки: ${status.folder_id}\n`;
      msg += `✅ Существует: ${status.exists ? "Да" : "Нет"}\n`;
      if (status.last_modified) {
          msg += `🕒 Изменен: ${status.last_modified}\n`;
      }
      
      ui.alert("Статус архивирования", msg, ui.ButtonSet.OK);
      return true;
    } else {
      ui.alert('❌ Ошибка сервера (' + code + '):\n' + text);
      return false;
    }
  } catch (e) {
    ui.alert('❌ Ошибка сети:\n' + e.message);
    return false;
  }
}

/**
 * Call server to structure documents for 353pp (Certification).
 */
function callServerStructureDocuments353pp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    "Сбор документов (353пп)",
    "Запустить сбор документов для заявки (353пп)?\n\n" +
    "Будут обработаны строки листа 'New sert' с видом документа '353'.\n" +
    "Документы будут скопированы в папку 'Заявка 353пп'.",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const payload = {
    spreadsheet_id: ss.getId()
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    timeout: 300 // Long timeout
  };

  try {
    ss.toast('⏳ Обработка документов...', 'Python Server', 60);

    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/documents/structure-353pp', options);
    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code === 200) {
      const result = JSON.parse(text);
      const data = result.data;
      
      let msg = '✅ Обработка завершена!\n\n';
      msg += 'Обработано строк: ' + data.processed + '\n';
      if (data.folder_url) {
        msg += '📁 Ссылка на папку: ' + data.folder_url + '\n';
      }
      
      if (data.details && data.details.length > 0) {
        msg += '\n\n⚠️ Есть ошибки (см. логи): ' + data.details.length;
      }

      ui.alert('Сбор документов 353пп', msg, ui.ButtonSet.OK);
      return true;
    } else {
      ui.alert('❌ Ошибка сервера (' + code + '):\n' + text);
      return false;
    }
  } catch (e) {
    ui.alert('❌ Ошибка сети:\n' + e.message);
    return false;
  }
}

// =======================================================================================
// AI / AGENT FUNCTIONS (Python Server)
// =======================================================================================

/**
 * Check Gemini AI service status via Python server.
 * Shows connection status, current model, and available models.
 */
function menuCheckService() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    ss.toast('⏳ Проверка Gemini API...', 'AI Agent', 10);

    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/ai/status', {
      muteHttpExceptions: true,
      timeout: 30
    });

    const code = response.getResponseCode();

    if (code === 200) {
      const result = JSON.parse(response.getContentText());

      if (result.status === 'online') {
        let msg = '✅ Gemini API онлайн!\n\n';
        msg += '🤖 Модель: ' + result.model + '\n';
        msg += '📝 Ответ теста: ' + (result.response || 'OK') + '\n\n';

        if (result.available_models && result.available_models.length > 0) {
          msg += '📋 Доступные модели:\n';
          result.available_models.slice(0, 5).forEach(function(m) {
            msg += '   • ' + m + '\n';
          });
        }

        ui.alert('Статус AI сервиса', msg, ui.ButtonSet.OK);
      } else {
        ui.alert('❌ Ошибка', 'Gemini API недоступен:\n' + (result.error || 'Unknown error'), ui.ButtonSet.OK);
      }
    } else {
      ui.alert('❌ Ошибка сервера', 'Код: ' + code + '\n' + response.getContentText(), ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('❌ Ошибка сети', e.message + '\n\nПроверьте, запущен ли Python сервер.', ui.ButtonSet.OK);
  }
}

/**
 * Show current Gemini settings from Python server.
 */
function showGeminiSettings() {
  const ui = SpreadsheetApp.getUi();

  try {
    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/ai/settings', {
      muteHttpExceptions: true,
      timeout: 10
    });

    if (response.getResponseCode() === 200) {
      const settings = JSON.parse(response.getContentText());

      let msg = '⚙️ Настройки Gemini AI\n\n';
      msg += '🤖 Модель: ' + settings.model + '\n';
      msg += '🔑 API ключ: ' + settings.api_key_masked + '\n';
      msg += '   Статус: ' + (settings.api_key_status === 'configured' ? '✅ Настроен' : '❌ Не настроен') + '\n';
      msg += '🔄 Max retries: ' + settings.max_retries + '\n';

      ui.alert('Настройки Gemini', msg, ui.ButtonSet.OK);
    } else {
      ui.alert('❌ Ошибка', 'Не удалось получить настройки', ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('❌ Ошибка', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Analyze the currently selected row via Python server.
 * Reads INCI link from column H, sends to AI, writes results to L-Y.
 */
function menuAnalyzeSelected() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const row = sheet.getActiveCell().getRow();

  if (row < 2) {
    ui.alert('⚠️ Внимание', 'Выберите строку данных (не заголовок)', ui.ButtonSet.OK);
    return;
  }

  const sheetName = sheet.getName();

  // Confirm
  const confirm = ui.alert(
    'Анализ строки ' + row,
    'Анализировать строку ' + row + ' в листе "' + sheetName + '"?\n\n' +
    'AI извлечет состав из PDF и заполнит колонки L-Y.',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    return;
  }

  try {
    ss.toast('⏳ Анализ строки ' + row + '... (может занять 10-30 сек)', 'AI Agent', 60);

    const payload = {
      spreadsheet_id: ss.getId(),
      sheet_name: sheetName,
      row_number: row
    };

    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/ai/analyze/row', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 120
    });

    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code === 200) {
      const result = JSON.parse(text);
      ss.toast('✅ Анализ завершен! Результаты записаны в строку ' + row, 'AI Agent', 5);

      // Show summary
      const data = result.data;
      ui.alert(
        '✅ Анализ завершен',
        'Строка: ' + row + '\n' +
        'Время: ' + result.duration + '\n\n' +
        '📦 Категория: ' + (data.category_code || 'N/A') + ' - ' + (data.category || '') + '\n' +
        '🏷️ Вид: ' + (data.product_type || 'N/A') + '\n' +
        '📄 Документ: ' + (data.reg_doc || 'N/A'),
        ui.ButtonSet.OK
      );
    } else {
      ss.toast('❌ Ошибка анализа', 'AI Agent', 5);
      ui.alert('❌ Ошибка', 'Код: ' + code + '\n\n' + text, ui.ButtonSet.OK);
    }
  } catch (e) {
    ss.toast('❌ Ошибка', 'AI Agent', 3);
    ui.alert('❌ Ошибка сети', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Simple AI analysis - quickly classify product without PDF.
 * Just fills "Вид продукции" and "Почему такой вид" columns.
 * Use this to test if Gemini API is working.
 */
function menuSimpleAnalyze() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const row = sheet.getActiveCell().getRow();

  if (row < 2) {
    ui.alert('⚠️ Внимание', 'Выберите строку данных (не заголовок)', ui.ButtonSet.OK);
    return;
  }

  // Read data from row (E - product name, I - purpose, J - application)
  const productName = sheet.getRange(row, 5).getValue();  // Column E
  const purpose = sheet.getRange(row, 9).getValue();       // Column I
  const application = sheet.getRange(row, 10).getValue();  // Column J

  if (!productName) {
    ui.alert('⚠️ Внимание', 'Наименование продукта (колонка E) пустое', ui.ButtonSet.OK);
    return;
  }

  try {
    ss.toast('⏳ Быстрый анализ: ' + productName.substring(0, 30) + '...', 'AI Agent', 30);

    const payload = {
      product_name: productName,
      purpose: purpose || null,
      application: application || null
    };

    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/ai/analyze/simple', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 60
    });

    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code === 200) {
      const result = JSON.parse(text);
      const data = result.data;

      // Write results to columns N and O
      sheet.getRange(row, 14).setValue(data.product_type || '');   // Column N - Вид продукции
      sheet.getRange(row, 15).setValue(data.product_reason || ''); // Column O - Почему

      ss.toast('✅ Готово! Вид: ' + (data.product_type || 'N/A'), 'AI Agent', 5);
      
      ui.alert(
        '✅ Анализ завершен',
        'Строка: ' + row + '\n\n' +
        '🏷️ Вид продукции: ' + (data.product_type || 'N/A') + '\n\n' +
        '📝 Обоснование: ' + (data.product_reason || 'N/A'),
        ui.ButtonSet.OK
      );
    } else {
      ss.toast('❌ Ошибка анализа', 'AI Agent', 5);
      ui.alert('❌ Ошибка', 'Код: ' + code + '\n\n' + text, ui.ButtonSet.OK);
    }
  } catch (e) {
    ss.toast('❌ Ошибка', 'AI Agent', 3);
    ui.alert('❌ Ошибка сети', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Call the Smart Match API on the Python server.
 * Returns the match result or null.
 * @param {string} productName - Name of the product to match
 * @returns {Object|null} Match result
 */
function callServerSmartMatch(productName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
    
  // Default to "Сертификация" sheet, or we could pass it as an argument?
  // For now, let's hardcode "Сертификация" as per the API default, but explicit is better.
  const sheetName = "Сертификация"; 

  const payload = {
    product_name: productName,
    spreadsheet_id: spreadsheetId,
    sheet_name: sheetName
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/products/match', options);
    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code === 200) {
      return JSON.parse(text);
    } else {
      console.error('SmartMatch error (' + code + '): ' + text);
      return null;
    }
  } catch (e) {
    console.error('SmartMatch network error:', e);
    return null;
  }
}

/**
 * Menu function: Smart Match for selected row.
 * Reads product name from the selected row and finds matching base product.
 * Найдёт базовый продукт для выбранной строки и предложит заполнение.
 */
function menuSmartMatch() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const row = sheet.getActiveCell().getRow();

  if (row < 2) {
    ui.alert('⚠️ Внимание', 'Выберите строку данных (не заголовок)', ui.ButtonSet.OK);
    return;
  }

  // Получаем название продукта из колонки "Наименования англ по ДС" (обычно колонка C или D)
  // Используем колонку E как основное название (или можно сделать динамически)
  const productName = sheet.getRange(row, 5).getValue(); // Column E - Наименование

  if (!productName) {
    ui.alert('⚠️ Внимание', 'Название продукта (колонка E) пустое', ui.ButtonSet.OK);
    return;
  }

  try {
    ss.toast('⏳ Поиск базового продукта для: ' + productName.substring(0, 40) + '...', 'Smart Match', 30);

    const result = callServerSmartMatch(productName);

    if (!result) {
      ui.alert('❌ Ошибка', 'Не удалось выполнить Smart Match. Проверьте логи.', ui.ButtonSet.OK);
      return;
    }

    if (result.match_found) {
      const matched = result.matched_product_details || {};
      let msg = '✅ Найден базовый продукт!\\n\\n';
      msg += '🎯 Matched ID: ' + result.matched_base_id + '\\n';
      msg += '📊 Confidence: ' + (result.confidence_score || 'N/A') + '\\n\\n';
      
      if (matched.name_eng_ds) {
        msg += '🇬🇧 Англ: ' + matched.name_eng_ds + '\\n';
      }
      if (matched.name_rus_ds) {
        msg += '🇷🇺 Рус: ' + matched.name_rus_ds + '\\n';
      }
      
      msg += '\\n📝 Причина: ' + (result.reasoning || 'N/A');

      const confirm = ui.alert(
        'Smart Match - Результат',
        msg + '\\n\\nЗаполнить данные из базового продукта?',
        ui.ButtonSet.YES_NO
      );

      if (confirm === ui.Button.YES) {
        // TODO: Автозаполнение данных из matched_product_details
        ss.toast('🚧 Автозаполнение пока не реализовано', 'Smart Match', 5);
      }
    } else {
      ui.alert(
        'Smart Match - Результат',
        '❌ Совпадение не найдено\\n\\n' +
        '📝 Причина: ' + (result.reasoning || 'Не указана') + '\\n\\n' +
        'Продукт "' + productName + '" не соответствует ни одному базовому продукту.',
        ui.ButtonSet.OK
      );
    }
  } catch (e) {
    ss.toast('❌ Ошибка', 'Smart Match', 3);
    ui.alert('❌ Ошибка', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Analyze all empty rows (rows with INCI link but no results).
 * Runs in background on the server.
 */
function menuAnalyzeEmpty() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  // Confirm
  const confirm = ui.alert(
    'Пакетный анализ',
    'Запустить анализ всех пустых строк в листе "' + sheetName + '"?\n\n' +
    '⚠️ Это может занять несколько минут.\n' +
    'Будут обработаны строки с INCI ссылкой, но без заполненных результатов.',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    return;
  }

  try {
    ss.toast('⏳ Пакетный анализ запущен в фоне...', 'AI Agent', 10);

    const payload = {
      spreadsheet_id: ss.getId(),
      sheet_name: sheetName,
      delay_between: 2.0
    };

    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/ai/analyze/batch', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 30
    });

    const code = response.getResponseCode();

    if (code === 200) {
      const result = JSON.parse(response.getContentText());
      ui.alert(
        '✅ Анализ запущен',
        'Пакетный анализ запущен в фоновом режиме.\n\n' +
        'Лист: ' + result.sheet_name + '\n' +
        'Статус: ' + result.status + '\n\n' +
        'Результаты будут записаны в таблицу по мере обработки.\n' +
        'Проверьте "Журнал логов" для отслеживания прогресса.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('❌ Ошибка', 'Код: ' + code + '\n' + response.getContentText(), ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('❌ Ошибка сети', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Show available ТН ВЭД categories.
 */
function menuShowCategories() {
  const ui = SpreadsheetApp.getUi();

  try {
    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/ai/categories', {
      muteHttpExceptions: true,
      timeout: 10
    });

    if (response.getResponseCode() === 200) {
      const result = JSON.parse(response.getContentText());
      const categories = result.categories || [];

      let msg = '📦 Категории ТН ВЭД для классификации:\n\n';

      categories.forEach(function(cat) {
        msg += '• ' + cat.code + ': ' + cat.name + '\n';
        msg += '   ' + cat.description + '\n\n';
      });

      ui.alert('Категории ТН ВЭД', msg, ui.ButtonSet.OK);
    } else {
      ui.alert('❌ Ошибка', 'Не удалось получить категории', ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('❌ Ошибка', e.message, ui.ButtonSet.OK);
  }
}

// =======================================================================================
// GEMINI API KEY STORAGE (Script Properties - persistent)
// =======================================================================================

const GEMINI_PROPS = {
  API_KEY: 'GEMINI_API_KEY',
  MODEL: 'GEMINI_MODEL'
};

const DEFAULT_GEMINI_MODEL = 'gemini-2.5-flash';

/**
 * Get stored Gemini API key from Script Properties.
 */
function getStoredGeminiKey() {
  return PropertiesService.getScriptProperties().getProperty(GEMINI_PROPS.API_KEY) || '';
}

/**
 * Get stored Gemini model from Script Properties.
 */
function getStoredGeminiModel() {
  return PropertiesService.getScriptProperties().getProperty(GEMINI_PROPS.MODEL) || DEFAULT_GEMINI_MODEL;
}

/**
 * Save Gemini settings to Script Properties.
 */
function saveGeminiSettings(apiKey, model) {
  const props = PropertiesService.getScriptProperties();
  if (apiKey) {
    props.setProperty(GEMINI_PROPS.API_KEY, apiKey);
  }
  if (model) {
    props.setProperty(GEMINI_PROPS.MODEL, model);
  }
}

/**
 * Initialize Gemini on server from stored Script Properties.
 * Called automatically on document open.
 */
function initGeminiFromStorage() {
  const apiKey = getStoredGeminiKey();
  const model = getStoredGeminiModel();

  if (!apiKey) {
    console.log('No Gemini API key stored in Script Properties');
    return false;
  }

  try {
    const payload = {
      api_key: apiKey,
      model: model
    };

    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/ai/configure', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 15
    });

    const code = response.getResponseCode();
    if (code === 200) {
      const result = JSON.parse(response.getContentText());
      if (result.status === 'success') {
        console.log('✅ Gemini initialized from storage, model:', result.model);
        return true;
      }
    }
    console.warn('Gemini init failed:', response.getContentText());
    return false;
  } catch (e) {
    console.error('Gemini init error:', e);
    return false;
  }
}

/**
 * Setup Gemini API - shows dialog to enter API key.
 * Key is saved to Script Properties (persistent) and sent to server.
 */
function setupGeminiComplete() {
  const ui = SpreadsheetApp.getUi();

  // Get stored settings
  const storedKey = getStoredGeminiKey();
  const storedModel = getStoredGeminiModel();
  const keyStatus = storedKey ? 'сохранён ✅' : 'не настроен ❌';

  // Ask user what to configure
  const choice = ui.alert(
    '⚙️ Настройка Gemini API',
    'Текущий статус:\n' +
    '• API ключ: ' + keyStatus + '\n' +
    '• Модель: ' + storedModel + '\n\n' +
    'Ключ хранится в Script Properties (постоянно).\n' +
    'Хотите ввести новый API ключ?',
    ui.ButtonSet.YES_NO
  );

  if (choice !== ui.Button.YES) {
    return;
  }

  // Show dialog for API key
  const keyResponse = ui.prompt(
    '🔑 Введите API ключ',
    'Введите ваш Gemini API ключ:\n\n' +
    'Получить ключ: https://aistudio.google.com/apikey\n\n' +
    'Ключ будет сохранён в Script Properties.',
    ui.ButtonSet.OK_CANCEL
  );

  if (keyResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const apiKey = keyResponse.getResponseText().trim();

  if (!apiKey) {
    ui.alert('❌ Ошибка', 'API ключ не может быть пустым', ui.ButtonSet.OK);
    return;
  }

  // Ask for model
  const modelResponse = ui.prompt(
    '🤖 Выберите модель',
    'Введите название модели (или оставьте пустым для ' + DEFAULT_GEMINI_MODEL + '):\n\n' +
    'Рекомендуемые модели:\n' +
    '• gemini-2.5-flash (быстрая, по умолчанию)\n' +
    '• gemini-2.0-flash (стабильная)\n' +
    '• gemini-2.5-pro (медленнее, но умнее)',
    ui.ButtonSet.OK_CANCEL
  );

  if (modelResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const model = modelResponse.getResponseText().trim() || DEFAULT_GEMINI_MODEL;

  // Save to Script Properties FIRST (persistent storage)
  saveGeminiSettings(apiKey, model);

  // Then send to server
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('⏳ Настройка Gemini API...', 'AI Agent', 10);

  try {
    const payload = {
      api_key: apiKey,
      model: model
    };

    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/ai/configure', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 30
    });

    const code = response.getResponseCode();
    const result = JSON.parse(response.getContentText());

    if (code === 200 && result.status === 'success') {
      ss.toast('✅ API ключ сохранён!', 'AI Agent', 3);
      ui.alert(
        '✅ Настройка завершена',
        'Gemini API успешно настроен!\n\n' +
        'Модель: ' + result.model + '\n' +
        'Тест: ' + (result.test_response || 'OK') + '\n\n' +
        'Ключ сохранён в Script Properties.\n' +
        'Теперь вы можете использовать AI анализ.',
        ui.ButtonSet.OK
      );
    } else {
      ss.toast('❌ Ошибка настройки', 'AI Agent', 3);
      ui.alert(
        '❌ Ошибка настройки',
        'Не удалось настроить Gemini API:\n\n' +
        (result.error || result.message || 'Unknown error') + '\n\n' +
        'Проверьте правильность API ключа.',
        ui.ButtonSet.OK
      );
    }
  } catch (e) {
    ss.toast('❌ Ошибка сети', 'AI Agent', 3);
    ui.alert('❌ Ошибка', 'Ошибка соединения с сервером:\n' + e.message, ui.ButtonSet.OK);
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

// =======================================================================================
// PRICE PROCESSING
// =======================================================================================

/**
 * Process supplier price list via Python server.
 * Replaces heavy GAS logic with server-side processing.
 *
 * @param {string} project - Project code (mt, sk, ss)
 * @param {string} mode - Processing mode (main, tester, samples, probes)
 * @param {Object} options - Additional options
 * @param {boolean} options.dryRun - If true, return preview without writing
 * @param {string} options.sourceDocId - Optional source document ID
 * @returns {Object} Processing result
 */
function callServerProcessPrice(project, mode, options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  options = options || {};

  const payload = {
    spreadsheet_id: spreadsheetId,
    mode: mode || 'main',
    dry_run: options.dryRun || false,
    source_doc_id: options.sourceDocId || null
  };

  const fetchOptions = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    timeout: 300 // 5 minutes for long operations
  };

  try {
    ss.toast('⏳ Обработка прайса...', project.toUpperCase() + ' ' + mode, 300);

    const response = UrlFetchApp.fetch(
      SERVER_URL + '/api/v1/price/process/' + project.toLowerCase(),
      fetchOptions
    );
    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code === 200) {
      const result = JSON.parse(text);

      if (result.status === 'preview') {
        ss.toast('📋 Превью готово', project.toUpperCase(), 5);
        return result;
      }

      if (result.status === 'success') {
        const message = result.message || 'Обработка завершена';
        const details = result.total_rows
          ? 'Строк: ' + result.total_rows + ', новых: ' + (result.total_new || 0)
          : result.processed_rows
            ? 'Строк: ' + result.processed_rows + ', групп: ' + (result.groups_found || 0)
            : '';
        ss.toast('✅ ' + message + (details ? '\n' + details : ''), project.toUpperCase(), 10);

        if (typeof Lib !== 'undefined' && typeof Lib.logInfo === 'function') {
          Lib.logInfo('[' + project.toUpperCase() + '] ' + message + '. ' + details);
        }

        return result;
      }

      // Error from server (includes partial errors from process_all)
      var errorMsg = result.message || result.error || 'Unknown error';
      if (result.errors && result.errors.length > 0) {
        errorMsg += '\n' + result.errors.join('\n');
      }
      ss.toast('❌ Ошибка: ' + errorMsg, project.toUpperCase(), 15);
      if (typeof Lib !== 'undefined' && typeof Lib.logError === 'function') {
        Lib.logError('[' + project.toUpperCase() + '] Ошибка обработки: ' + errorMsg);
      }
      return { status: 'error', message: errorMsg };

    } else {
      const errorMsg = 'HTTP ' + code + ': ' + text.substring(0, 200);
      ss.toast('❌ Ошибка сервера', project.toUpperCase(), 10);
      return { status: 'error', message: errorMsg };
    }

  } catch (e) {
    const errorMsg = e.message || String(e);
    ss.toast('❌ Ошибка: ' + errorMsg, project.toUpperCase(), 10);
    console.error('[' + project.toUpperCase() + '] callServerProcessPrice failed:', e);
    return { status: 'error', message: errorMsg };
  }
}

/**
 * Загрузить остатки через сервер
 * @param {string} sourceDocId - Опциональный ID документа-источника
 * @returns {Object} Результат обработки
 */
function callServerLoadStocks(sourceDocId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();

  const payload = {
    spreadsheet_id: spreadsheetId,
    source_doc_id: sourceDocId || null
  };

  const fetchOptions = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    ss.toast('⏳ Загрузка остатков через сервер...', 'Пожалуйста, подождите', 300);

    const response = UrlFetchApp.fetch(
      SERVER_URL + '/api/v1/stocks/load',
      fetchOptions
    );
    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code === 200) {
      const result = JSON.parse(text);
      ss.toast('✅ Остатки загружены: ' + result.updated_rows + ' строк', 'Готово', 5);
      
      // Вызываем пересчет на стороне GAS (если нужно)
      if (typeof Lib.recalculateOrderSheet === 'function') {
        Lib.recalculateOrderSheet();
      }
      
      return result;
    }

    const errorMsg = 'Ошибка при загрузке остатков: ' + (text.substring(0, 500) || 'Неизвестная ошибка');
    ss.toast('❌ ' + errorMsg, 'Ошибка', 10);
    return { status: 'error', message: text };
  } catch (e) {
    console.error('[callServerLoadStocks] Error:', e);
    ss.toast('❌ ' + e.toString(), 'Критическая ошибка', 10);
    return { status: 'error', message: e.toString() };
  }
}

/**
 * Process MT main price (Б/З поставщик) via server.
 */
function callServerProcessMtMain() {
  return callServerProcessPrice('mt', 'main');
}

/**
 * Process MT tester price via server.
 */
function callServerProcessMtTester() {
  return callServerProcessPrice('mt', 'tester');
}

/**
 * Process MT samples price via server.
 */
function callServerProcessMtSamples() {
  return callServerProcessPrice('mt', 'samples');
}

/**
 * Process SK main price via server.
 */
function callServerProcessSkMain() {
  return callServerProcessPrice('sk', 'main');
}

/**
 * Process SK probes price via server.
 */
function callServerProcessSkProbes() {
  return callServerProcessPrice('sk', 'probes');
}

/**
 * Process SS main price via server.
 */
function callServerProcessSsMain() {
  return callServerProcessPrice('ss', 'main');
}
/**
 * Вспомогательная функция для открытия веб-страниц сервера.
 * @param {string} path - Путь страницы (например, "/docs")
 * @param {string} title - Заголовок окна
 */
function openServerUrl(path, title) {
  const baseUrl = SERVER_URL; 
  const url = baseUrl + path;
  
  const html = `
    <html>
      <body style="font-family: sans-serif; text-align: center; padding: 20px; background: #1e293b; color: white;">
        <h3 style="margin-bottom: 20px;">${title}</h3>
        <p>Для просмотра журнала нажмите кнопку ниже:</p>
        <p>
          <a href="${url}" target="_blank" style="
            display: inline-block;
            background: #0d9488;
            color: white;
            padding: 12px 24px;
            text-decoration: none;
            border-radius: 6px;
            font-weight: bold;
            font-size: 16px;">
            Открыть журнал
          </a>
        </p>
        <p style="margin-top: 20px; font-size: 12px; color: #94a3b8;">
          ID таблицы: <br>
          <span style="font-family: monospace; background: #0f172a; padding: 4px; border-radius: 4px;">${path.split('id=')[1] || 'Не указан'}</span>
        </p>
        <p style="margin-top: 10px; font-size: 10px; color: #64748b;">
          Ссылка: <br>
          ${url}
        </p>
      </body>
    </html>
  `;
  
  const userInterface = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(300);
    
  SpreadsheetApp.getUi().showModalDialog(userInterface, '🚀 ' + title);
}

// Функции-обертки для меню
function openServerMainPage() { openServerUrl('/', 'Главная страница'); }
function openServerRulesPage() { openServerUrl('/api/v1/rules-ui', 'Правила синхронизации'); }
function openServerDocsPage() { openServerUrl('/docs', 'Swagger UI'); }
function openLogDashboard_proxy() { openServerUrl('/api/v1/logs-ui', 'Журнал логов'); }
function openLogsManager() { openServerUrl('/api/v1/logs-ui', 'Журнал логов'); }
function showSyncLogDialog() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const id = ss.getId();
  openServerUrl('/api/v1/logs-ui?id=' + id, 'Журнал синхронизации'); 
}
