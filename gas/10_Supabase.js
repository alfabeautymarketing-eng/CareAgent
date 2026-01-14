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
/**
 * Shows available tables in Supabase database
 */
function showSupabaseTablesView() {
  const ui = SpreadsheetApp.getUi();
  const SERVER_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL');
  
  if (!SERVER_URL) {
    ui.alert('❌ Error: SERVER_URL not set in Script Properties');
    return;
  }

  try {
    ui.alert('⏳ Loading tables from Supabase...');
    const response = UrlFetchApp.fetch(SERVER_URL + '/api/v1/supabase/tables', {
      method: 'get',
      muteHttpExceptions: true
    });
    
    const result = JSON.parse(response.getContentText());
    if (response.getResponseCode() !== 200) {
      throw new Error(result.detail || 'Unknown error');
    }

    const tables = result.tables;
    if (!tables || tables.length === 0) {
      ui.alert('🗄️ Supabase Tables', 'No public tables found.', ui.ButtonSet.OK);
      return;
    }

    const html = HtmlService.createHtmlOutput(
      '<style>body { font-family: sans-serif; padding: 10px; } ul { list-style-type: none; padding: 0; } li { padding: 8px; border-bottom: 1px solid #eee; } li:last-child { border-bottom: none; } .table-icon { margin-right: 8px; }</style>' +
      '<h3>🗄️ Supabase Tables</h3>' +
      '<ul>' +
      tables.map(function(t) { return '<li><span class="table-icon">📄</span>' + t + '</li>'; }).join('') +
      '</ul>'
    ).setWidth(300).setHeight(400);

    ui.showModelessDialog(html, 'Supabase Tables');

  } catch (e) {
    console.error(e);
    ui.alert('❌ Error fetching tables:\n' + e.toString());
  }
}

/**
 * Execute SQL query against Supabase database
 */
function executeSupabaseSqlQuery() {
  const ui = SpreadsheetApp.getUi();
  const SERVER_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL');

  if (!SERVER_URL) {
    ui.alert('❌ Error: SERVER_URL not set');
    return;
  }

  const response = ui.prompt('SQL Query', 'SELECT * FROM users LIMIT 5', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    const query = response.getResponseText();
    ui.toast('🚀 Executing SQL...', 'Supabase');
    
    try {
      const apiResponse = UrlFetchApp.fetch(SERVER_URL + '/api/v1/supabase/query', {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ query: query }),
        muteHttpExceptions: true
      });

      const result = JSON.parse(apiResponse.getContentText());
      if (apiResponse.getResponseCode() !== 200) {
        throw new Error(result.detail || 'Unknown error');
      }

      const data = result.data;
      if (!data || data.length === 0) {
        ui.alert('✅ Query executed successfully. No rows returned.');
      } else {
        // Show results in a simplified way or specialized sheet
        const headers = Object.keys(data[0]);
        const headerRow = '<tr>' + headers.map(function(h) { return '<th>' + h + '</th>'; }).join('') + '</tr>';
        const rows = data.map(function(row) {
          return '<tr>' + headers.map(function(h) { return '<td>' + (row[h] !== null ? row[h] : '') + '</td>'; }).join('') + '</tr>';
        }).join('');

        const html = HtmlService.createHtmlOutput(
          '<style>table { border-collapse: collapse; width: 100%; font-family: sans-serif; font-size: 12px; } th, td { border: 1px solid #ddd; padding: 4px; text-align: left; } th { background-color: #f2f2f2; }</style>' +
          '<h3>Query Results (' + data.length + ' rows)</h3>' +
          '<table>' + headerRow + rows + '</table>'
        ).setWidth(800).setHeight(600);

        ui.showModelessDialog(html, 'SQL Results');
      }

    } catch (e) {
      console.error(e);
      ui.alert('❌ SQL Error:\n' + e.toString());
    }
  }
}

/**
 * Import data from Supabase into Google Sheets
 */
function importSupabaseData() {
  const ui = SpreadsheetApp.getUi();
  const SERVER_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL');
  
  if (!SERVER_URL) {
    ui.alert('❌ Error: SERVER_URL not set');
    return;
  }

  const response = ui.prompt('Import Table', 'Enter table name (e.g. products)', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    const tableName = response.getResponseText();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    ui.toast('📥 Importing data...', 'Supabase');

    try {
      const apiResponse = UrlFetchApp.fetch(SERVER_URL + '/api/v1/supabase/table/' + tableName + '?limit=100', {
        method: 'get',
        muteHttpExceptions: true
      });

      const result = JSON.parse(apiResponse.getContentText());
      if (apiResponse.getResponseCode() !== 200) {
        throw new Error(result.detail || 'Unknown error');
      }

      const data = result.data;
      if (!data || data.length === 0) {
        ui.alert('⚠️ Table exists but returned no data (or empty).');
        return;
      }

      // Write data to sheet
      sheet.clear();
      const headers = Object.keys(data[0]);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

      const values = data.map(function(row) {
        return headers.map(function(header) { return row[header]; });
      });

      if (values.length > 0) {
        sheet.getRange(2, 1, values.length, headers.length).setValues(values);
      }
      
      ui.alert('✅ Imported ' + values.length + ' rows from ' + tableName);

    } catch (e) {
      console.error(e);
      ui.alert('❌ Import Error:\n' + e.toString());
    }
  }
}

/**
 * Export current sheet data to Supabase
 */
function exportSupabaseData() {
  const ui = SpreadsheetApp.getUi();
  const SERVER_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL');

  if (!SERVER_URL) {
    ui.alert('❌ Error: SERVER_URL not set');
    return;
  }

  ui.alert('⚠️ Export feature is currently read-only in this version by safety default.\n\nPlease define target table structure first.');
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
