// CONFIGURATION
// Replace with your actual server URL (e.g. https://api.yourdomain.com)
// For local testing with ngrok: https://xxxx-xx-xx-xx-xx.ngrok-free.app
const SERVER_URL = 'http://46.226.167.153:8000'; 

/**
 * Triggers sort by Manufacturer.
 * Assumes column header is "Производитель".
 */
function triggerSortManufacturer() {
  callServerSort('Производитель', true);
}

/**
 * Triggers sort by Price.
 * Assumes column header is "Цена" or "Price".
 */
function triggerSortPrice() {
  // Trying 'Цена' first, you can change this to 'Price' if needed
  callServerSort('Цена', true);
}

/**
 * Check server health.
 */
function checkServerStatus() {
  try {
    const response = UrlFetchApp.fetch(SERVER_URL + '/health');
    const json = JSON.parse(response.getContentText());
    if (json.status === 'healthy') {
      SpreadsheetApp.getUi().alert('✅ Сервер онлайн\nВерсия: ' + json.version);
    } else {
      SpreadsheetApp.getUi().alert('⚠️ Сервер ответил, но статус: ' + json.status);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Ошибка подключения к серверу:\n' + e.message);
  }
}

/**
 * Generic function to call the sort endpoint.
 */
function callServerSort(columnName, ascending) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  const payload = {
    spreadsheet_id: ss.getId(),
    sheet_name: sheet.getName(),
    column_name: columnName,
    ascending: ascending
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(SERVER_URL + '/sort', options);
    const code = response.getResponseCode();
    const text = response.getContentText();
    
    if (code === 200) {
      SpreadsheetApp.getActiveSpreadsheet().toast('✅ Сортировка выполнена: ' + columnName);
    } else {
      SpreadsheetApp.getUi().alert('❌ Ошибка сервера (' + code + '):\n' + text);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Ошибка сети:\n' + e.message + '\n\nПроверьте, запущен ли сервер и доступен ли он.');
  }
}
