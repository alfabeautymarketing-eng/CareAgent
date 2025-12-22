/**
 * =======================================================================================
 * SERVER-SIDE OVERRIDES (z-server-overrides.js)
 * ---------------------------------------------------------------------------------------
 * Replaces local logic with Server API calls for consistency and performance.
 * Loaded last (z-) to override previous definitions.
 * =======================================================================================
 */

var Lib = Lib || {};

(function (Lib) {
    
    // --- HELPER: Call Server ---
    Lib.callServer = function(endpoint, payload) {
        // Find server URL in existing config or hardcode/detect
        // We assume 03Синхронизация.js has some SERVER_URL or we define it here.
        // During dev it is http://46.226.167.153:8000
        const SERVER_URL = "http://46.226.167.153:8000"; // Update if changed
        
        try {
            const options = {
                method: "post",
                contentType: "application/json",
                payload: JSON.stringify(payload),
                headers: {
                    "ngrok-skip-browser-warning": "true" 
                },
                muteHttpExceptions: true
            };
            
            const url = `${SERVER_URL}${endpoint}`;
            Lib.logInfo(`[Server] Calling ${url}...`);
            const response = UrlFetchApp.fetch(url, options);
            const code = response.getResponseCode();
            const text = response.getContentText();
            let json = {};
            try { json = JSON.parse(text); } catch(e) {}
            
            if (code >= 200 && code < 300) {
                Lib.logInfo(`[Server] Success: ${text}`);
                return json;
            } else {
                throw new Error(`Server Error (${code}): ${json.detail || text}`);
            }
        } catch (e) {
            Lib.logError(`[Server] Call Failed: ${e.message}`, e);
            throw e;
        }
    };

    // --- OVERRIDE: Add Article ---
    Lib.addArticleManually = function() {
        const ui = SpreadsheetApp.getUi();
        const response = ui.prompt("Добавить артикул", "Введите ID артикула:", ui.ButtonSet.OK_CANCEL);
        if (response.getSelectedButton() !== ui.Button.OK) return;
        
        const article = response.getResponseText().trim();
        if (!article) return;
        
        try {
            const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
            // We deduce project from SS ID on server
            const res = Lib.callServer("/sync/add-article", {
                article: article,
                spreadsheet_id: ssId,
                project: "UNKNOWN" // Server resolves it
            });
            ui.alert(`✅ Артикул создан!\n${res.message}`);
        } catch (e) {
            ui.alert(`❌ Ошибка: ${e.message}`);
        }
    };

    // --- OVERRIDE: Delete Selected Rows ---
    Lib.deleteSelectedRowsWithSync = function() {
        const ui = SpreadsheetApp.getUi();
        const sel = SpreadsheetApp.getActiveRangeList();
        if (!sel) { ui.alert("Выберите строки"); return; }
        
        const sheet = SpreadsheetApp.getActiveSheet();
        const rows = new Set();
        sel.getRanges().forEach(r => {
            for (let i = r.getRow(); i <= r.getLastRow(); i++) if(i>1) rows.add(i);
        });
        
        if (rows.size === 0) { ui.alert("Нет строк для удаления"); return; }
        
        const confirm = ui.alert("Подтверждение удаление (Сервер)", 
            `Удалить ${rows.size} строк через сервер? Это удалит артикулы из всех связанных листов.`, 
            ui.ButtonSet.YES_NO);
        if (confirm !== ui.Button.YES) return;
        
        const ids = [];
        rows.forEach(r => {
            const val = String(sheet.getRange(r, 1).getValue() || "").trim();
            if (val) ids.push(val);
        });
        
        if (ids.length === 0) { ui.alert("Не найдены ID в колонке A."); return; }
        
        try {
            const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
            const res = Lib.callServer("/sync/delete-articles", {
                articles: ids,
                spreadsheet_id: ssId,
                project: "UNKNOWN"
            });
            ui.alert(`✅ Удалено!\n${res.message}`);
            // Force refresh or delete locally too? 
            // Server deleted rows, but GAS sheet might need refresh to see changes or we delete locally to be instant.
            // Better to delete locally too to avoid confusion, but server does it. 
            // If server deletes, we should reload? Google Sheets updates automatically.
        } catch (e) {
            ui.alert(`❌ Ошибка: ${e.message}`);
        }
    };

    // --- OVERRIDE: Sync Selected Row ---
    Lib.syncSelectedRow = function() {
        const ui = SpreadsheetApp.getUi();
        const row = SpreadsheetApp.getActiveRange().getRow();
        if (row <= 1) return;
        const sheet = SpreadsheetApp.getActiveSheet();
        const article = String(sheet.getRange(row, 1).getValue() || "").trim();
        
        if (!article) { ui.alert("Нет ID в этой строке"); return; }

        try {
            ui.toast("Синхронизация строки...");
            const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
            const res = Lib.callServer("/sync/row", {
                spreadsheet_id: ssId,
                sheet_name: sheet.getName(),
                article: article,
                project: "UNKNOWN"
            });
            ui.alert(`✅ Синхронизация завершена.`);
        } catch (e) {
            ui.alert(`❌ Ошибка: ${e.message}`);
        }
    };
    
    // --- OVERRIDE: Run Full Sync ---
    Lib.runFullSync = function() {
        const ui = SpreadsheetApp.getUi();
        const sheet = SpreadsheetApp.getActiveSheet();
        const confirm = ui.alert("Полная синхронизация", 
            `Запустить полную синхронизацию для листа "${sheet.getName()}" через сервер? Это может занять время.`, 
            ui.ButtonSet.YES_NO);
        if (confirm !== ui.Button.YES) return;

        try {
            ui.toast("Запуск полной синхронизации...");
            const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
            // Start async or sync? Endpoint returns task_id but waits by default in current impl.
            const res = Lib.callServer("/sync/full", {
                spreadsheet_id: ssId,
                source_sheet: sheet.getName(),
                project: "UNKNOWN"
            });
            ui.alert(`✅ Завершено.\nСтатус: ${res.status}`);
        } catch (e) {
            ui.alert(`❌ Ошибка: ${e.message}`);
        }
    };

})(Lib);
