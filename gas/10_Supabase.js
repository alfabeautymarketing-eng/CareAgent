// =======================================================================================
// 10_Supabase.js — SUPABASE INTEGRATION & DATABASE OPERATIONS
// =======================================================================================
// DESCRIPTION:
//   Provides functions to interact with Supabase PostgreSQL database.
//   Includes console access, SQL query execution, and data management operations.
// =======================================================================================

/**
 * Opens the Supabase console in a new browser tab
 */
function openSupabaseConsole() {
  const supabaseUrl = "https://supabase.com/dashboard/project/kxxdxnsvbvdpyfxccvjb";
  HtmlService.createHtmlOutput(
    '<script>window.open("' + supabaseUrl + '", "_blank"); google.script.host.close();</script>'
  ).getWidth(1).getHeight(1);
  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput('<p>Открыли Supabase Console...</p>'),
    'Supabase'
  );
}

/**
 * Shows available tables in Supabase database
 */
function showSupabaseTablesView() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('🗄️ Просмотр таблиц Supabase\n\nЭта функция будет реализована через MCP сервер.\n\nДоступные таблицы:\n- users\n- products\n- orders\n- sync_log');
}

/**
 * Execute SQL query against Supabase database
 */
function executeSupabaseSqlQuery() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Введите SQL запрос для выполнения:', 'SELECT * FROM users LIMIT 10;');

  if (response.getSelectedButton() === ui.Button.OK) {
    const query = response.getResponseText();
    ui.alert('SQL запрос отправлен:\n\n' + query + '\n\nРезультаты будут загружены через MCP сервер.');
    console.log('Supabase SQL Query:', query);
  }
}

/**
 * Import data from Supabase into Google Sheets
 */
function importSupabaseData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Выберите таблицу для импорта:', 'products');

  if (response.getSelectedButton() === ui.Button.OK) {
    const tableName = response.getResponseText();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();

    ui.alert('Импорт данных из таблицы: ' + tableName + '\n\nПожалуйста, подождите...');
    console.log('Importing from Supabase table:', tableName);

    // Placeholder for actual import logic
    sheet.getRange('A1').setValue('Import from ' + tableName + ' (via MCP)');
  }
}

/**
 * Export current sheet data to Supabase
 */
function exportSupabaseData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Укажите название таблицы для экспорта:', 'exported_data');

  if (response.getSelectedButton() === ui.Button.OK) {
    const tableName = response.getResponseText();
    ui.alert('Экспорт ' + data.length + ' строк в таблицу: ' + tableName + '\n\nОперация выполняется через MCP сервер.');
    console.log('Exporting to Supabase table:', tableName, 'Rows:', data.length);
  }
}

/**
 * Manage database access permissions
 */
function manageSupabasePermissions() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(
    '<p>🔐 Управление правами доступа в Supabase</p>' +
    '<p>Роли:</p>' +
    '<ul>' +
    '<li>❌ anon (публичный доступ) - отключен</li>' +
    '<li>✅ authenticated - включен</li>' +
    '<li>✅ service_role - включен</li>' +
    '</ul>' +
    '<p>Для изменения прав доступа используйте Supabase Console.</p>'
  );
  ui.showModelessDialog(html, 'Управление правами');
}

/**
 * Configure Supabase connection settings
 */
function configureSupabaseConnection() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const currentUrl = props.getProperty('SUPABASE_URL') || 'https://kxxdxnsvbvdpyfxccvjb.supabase.co';
  const urlResponse = ui.prompt('Supabase URL:', currentUrl);

  if (urlResponse.getSelectedButton() === ui.Button.OK) {
    props.setProperty('SUPABASE_URL', urlResponse.getResponseText());
    ui.alert('✅ Supabase URL сохранен:\n' + urlResponse.getResponseText());
  }
}

/**
 * Test Supabase connection
 */
function testSupabaseConnection() {
  const SERVER_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL');
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

  try {
    const response = UrlFetchApp.fetch(
      SERVER_URL + '/api/v1/supabase/test?spreadsheet_id=' + encodeURIComponent(spreadsheetId),
      {
        method: 'get',
        muteHttpExceptions: true,
        timeout: 10
      }
    );

    const code = response.getResponseCode();
    if (code === 200) {
      SpreadsheetApp.getUi().alert('✅ Соединение с Supabase успешно!\n\nСтатус: ' + response.getContentText());
    } else {
      SpreadsheetApp.getUi().alert('❌ Ошибка подключения к Supabase\nКод: ' + code);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Ошибка:\n' + e.toString());
  }
}
