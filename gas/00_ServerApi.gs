/**
 * Server API Wrapper for GAS
 *
 * This module provides utilities to call AgentCare server endpoints from Google Apps Script.
 * It replaces local heavy-lifting operations with efficient server API calls.
 *
 * Usage:
 *   var result = ServerApi.processPriceList('mt', 'main', spreadsheetId, sourceDocId);
 *   var status = ServerApi.getPriceStatus(taskId);
 */

var ServerApi = ServerApi || {};

(function (exports) {
  // ============================================
  // Configuration
  // ============================================

  /**
   * Get server base URL from project config or environment
   * Reads from: Lib.CONFIG.SERVER_URL or falls back to default
   */
  function _getServerUrl() {
    if (Lib && Lib.CONFIG && Lib.CONFIG.SERVER_URL) {
      return Lib.CONFIG.SERVER_URL;
    }
    // Fallback to environment or default
    var scriptProperties = PropertiesService.getScriptProperties();
    var serverUrl = scriptProperties.getProperty('SERVER_URL');
    if (!serverUrl) {
      // Default - should be overridden in production
      serverUrl = 'https://agentcare.example.com';
    }
    return serverUrl;
  }

  /**
   * Set server URL in script properties
   */
  exports.setServerUrl = function (url) {
    PropertiesService.getScriptProperties().setProperty('SERVER_URL', url);
  };

  // ============================================
  // Price Processing
  // ============================================

  /**
   * Process supplier price list on the server
   *
   * @param {string} project - Project code: 'mt', 'sk', or 'ss'
   * @param {string} mode - Processing mode: 'main', 'tester', 'samples', or 'auto'
   * @param {string} spreadsheetId - Target spreadsheet ID (main project sheet)
   * @param {string} sourceDocId - Source document ID (with supplier prices)
   * @param {Object} options - Optional parameters: { dryRun: false, showUI: true }
   * @returns {Object} { status, message, taskId, processedRows, createdArticles, etc. }
   */
  exports.processPriceList = function (project, mode, spreadsheetId, sourceDocId, options) {
    options = options || {};
    var showUI = options.showUI !== false;
    var dryRun = options.dryRun || false;

    if (!project || !mode || !spreadsheetId || !sourceDocId) {
      throw new Error('Missing required parameters: project, mode, spreadsheetId, sourceDocId');
    }

    var ui = SpreadsheetApp.getUi();
    var serverUrl = _getServerUrl();
    var url = serverUrl + '/api/v1/price/process/' + project.toLowerCase();

    var payload = {
      spreadsheet_id: spreadsheetId,
      source_doc_id: sourceDocId,
      mode: mode.toLowerCase(),
      dry_run: dryRun
    };

    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 300 // 5 minutes
    };

    if (showUI) {
      ui.showModelessDialog(
        HtmlService.createHtmlOutput(
          '<div style="text-align: center; padding: 20px; font-family: Arial;">' +
          '<p style="font-size: 16px;">⏳ Обработка данных на сервере...</p>' +
          '<p style="font-size: 12px; color: #666;">Пожалуйста, подождите</p>' +
          '</div>'
        ),
        '🔄 Обработка прайс-листа'
      );
    }

    try {
      var response = UrlFetchApp.fetch(url, options);
      var httpCode = response.getResponseCode();
      var resultText = response.getContentText();
      var result = JSON.parse(resultText);

      if (httpCode >= 400) {
        throw new Error('Server error (' + httpCode + '): ' + result.detail);
      }

      if (showUI) {
        ui.alert(
          '✅ Обработка завершена!',
          'Статус: ' + (result.message || 'OK'),
          ui.ButtonSet.OK
        );
      }

      return result;

    } catch (error) {
      if (showUI) {
        ui.alert(
          '❌ Ошибка при обработке',
          'Деталь: ' + error.toString(),
          ui.ButtonSet.OK
        );
      }
      throw error;
    }
  };

  /**
   * Get price processing status by task ID
   *
   * @param {string} taskId - Task ID returned from processPriceList()
   * @returns {Object} { status, progress, message, result }
   */
  exports.getPriceStatus = function (taskId) {
    if (!taskId) {
      throw new Error('taskId is required');
    }

    var serverUrl = _getServerUrl();
    var url = serverUrl + '/api/v1/price/status/' + taskId;

    var options = {
      method: 'get',
      muteHttpExceptions: true,
      timeout: 30
    };

    try {
      var response = UrlFetchApp.fetch(url, options);
      var result = JSON.parse(response.getContentText());
      return result;
    } catch (error) {
      Lib.logError('Failed to get price status: ' + error.toString());
      return null;
    }
  };

  /**
   * Cancel price processing task
   *
   * @param {string} taskId - Task ID to cancel
   * @returns {Object} { status, message }
   */
  exports.cancelPriceProcessing = function (taskId) {
    if (!taskId) {
      throw new Error('taskId is required');
    }

    var serverUrl = _getServerUrl();
    var url = serverUrl + '/api/v1/price/cancel/' + taskId;

    var options = {
      method: 'post',
      muteHttpExceptions: true,
      timeout: 30
    };

    try {
      var response = UrlFetchApp.fetch(url, options);
      var result = JSON.parse(response.getContentText());
      return result;
    } catch (error) {
      Lib.logError('Failed to cancel price processing: ' + error.toString());
      throw error;
    }
  };

  // ============================================
  // Sorting/Structuring
  // ============================================

  /**
   * Sort order data on the server
   *
   * @param {string} spreadsheetId - Spreadsheet ID
   * @param {string} sortType - 'by_manufacturer', 'by_price', etc.
   * @param {Object} options - Optional parameters
   * @returns {Object} { status, message, sortedRows }
   */
  exports.sortOrderData = function (spreadsheetId, sortType, options) {
    options = options || {};

    var serverUrl = _getServerUrl();
    var url = serverUrl + '/api/v1/sort/structure';

    var payload = {
      spreadsheet_id: spreadsheetId,
      sort_type: sortType,
      options: options
    };

    var fetchOptions = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 60
    };

    try {
      var response = UrlFetchApp.fetch(url, fetchOptions);
      var result = JSON.parse(response.getContentText());

      if (response.getResponseCode() >= 400) {
        throw new Error('Server error: ' + result.detail);
      }

      return result;

    } catch (error) {
      Lib.logError('Sorting failed: ' + error.toString());
      throw error;
    }
  };

  // ============================================
  // Formula Operations
  // ============================================

  /**
   * Recalculate price dynamics formulas on the server
   *
   * @param {string} spreadsheetId - Spreadsheet ID
   * @param {Object} options - Optional parameters
   * @returns {Object} { status, message, formulasUpdated }
   */
  exports.recalculatePriceDynamicsFormulas = function (spreadsheetId, options) {
    options = options || {};

    var serverUrl = _getServerUrl();
    var url = serverUrl + '/api/v1/formulas/price-dynamics';

    var payload = {
      spreadsheet_id: spreadsheetId,
      options: options
    };

    var fetchOptions = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 60
    };

    try {
      var response = UrlFetchApp.fetch(url, fetchOptions);
      var result = JSON.parse(response.getContentText());

      if (response.getResponseCode() >= 400) {
        throw new Error('Server error: ' + result.detail);
      }

      return result;

    } catch (error) {
      Lib.logError('Formula recalculation failed: ' + error.toString());
      throw error;
    }
  };

  /**
   * Update price calculation formulas on the server
   *
   * @param {string} spreadsheetId - Spreadsheet ID
   * @param {Object} options - Optional parameters
   * @returns {Object} { status, message, formulasUpdated }
   */
  exports.updatePriceCalculationFormulas = function (spreadsheetId, options) {
    options = options || {};

    var serverUrl = _getServerUrl();
    var url = serverUrl + '/api/v1/formulas/price-calculation';

    var payload = {
      spreadsheet_id: spreadsheetId,
      options: options
    };

    var fetchOptions = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 60
    };

    try {
      var response = UrlFetchApp.fetch(url, fetchOptions);
      var result = JSON.parse(response.getContentText());

      if (response.getResponseCode() >= 400) {
        throw new Error('Server error: ' + result.detail);
      }

      return result;

    } catch (error) {
      Lib.logError('Formula update failed: ' + error.toString());
      throw error;
    }
  };

  // ============================================
  // Health Check
  // ============================================

  /**
   * Check if server is reachable
   *
   * @returns {boolean} true if server is reachable, false otherwise
   */
  exports.isServerAvailable = function () {
    try {
      var serverUrl = _getServerUrl();
      var response = UrlFetchApp.fetch(serverUrl + '/health', {
        method: 'get',
        muteHttpExceptions: true,
        timeout: 5
      });
      return response.getResponseCode() === 200;
    } catch (error) {
      Lib.logWarning('Server health check failed: ' + error.toString());
      return false;
    }
  };

  /**
   * Get server status and information
   *
   * @returns {Object} { status, version, endpoints, etc. }
   */
  exports.getServerStatus = function () {
    try {
      var serverUrl = _getServerUrl();
      var url = serverUrl.includes('/api/v1') ? serverUrl + '/status' : serverUrl + '/api/v1/status';
      var response = UrlFetchApp.fetch(url, {
        method: 'get',
        muteHttpExceptions: true,
        timeout: 10
      });
      var result = JSON.parse(response.getContentText());
      return result;
    } catch (error) {
      Lib.logWarning('Failed to get server status: ' + error.toString());
      return null;
    }
  };

  /**
   * Run full synchronization on the server
   *
   * @param {string} spreadsheetId - Spreadsheet ID
   * @param {string} project - Project code (e.g. 'MT', 'SK', 'SS')
   * @param {string} sourceSheet - Name of the source sheet (optional in global mode, but good to pass if needed)
   * @returns {Object} { status, details, task_id }
   */
  exports.syncFull = function (spreadsheetId, project, sourceSheet) {
    if (!spreadsheetId || !project) {
      throw new Error('Missing required parameters: spreadsheetId, project');
    }

    var serverUrl = _getServerUrl();
    var baseUrl = serverUrl.includes('/api/v1') ? serverUrl : serverUrl + '/api/v1';
    var url = baseUrl + '/sync/full?project=' + encodeURIComponent(project) + 
              '&source_sheet=' + encodeURIComponent(sourceSheet || 'ALL') + 
              '&spreadsheet_id=' + encodeURIComponent(spreadsheetId);

    var options = {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      timeout: 300 // 5 minutes
    };

    try {
      var response = UrlFetchApp.fetch(url, options);
      var result = JSON.parse(response.getContentText());

      if (response.getResponseCode() >= 400) {
        throw new Error('Server error: ' + (result.detail || result.message || 'Unknown error'));
      }

      return result;
    } catch (error) {
      Lib.logError('Full sync failed: ' + error.toString());
      throw error;
    }
  };

})(ServerApi);

// ============================================
// Export for use in other GAS modules
// ============================================
// Use: ServerApi.processPriceList(...), ServerApi.getPriceStatus(...), etc.
