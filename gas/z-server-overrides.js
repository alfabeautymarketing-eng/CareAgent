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
        const SERVER_URL = "http://46.226.167.153:8000";
        
        // Log Entry
        if (Lib.logStep) {
            Lib.logStep("Network", `>>> START: ${endpoint}`, "DEBUG");
        }
        
        try {
            const options = {
                method: "post",
                contentType: "application/json",
                payload: JSON.stringify(payload),
                headers: { "ngrok-skip-browser-warning": "true" },
                muteHttpExceptions: true
            };
            
            let finalEndpoint = endpoint;
            if (!finalEndpoint.startsWith("/api/v1")) {
                if (!finalEndpoint.startsWith("/")) finalEndpoint = "/" + finalEndpoint;
                finalEndpoint = "/api/v1" + finalEndpoint;
            }

            const url = `${SERVER_URL}${finalEndpoint}`;
            const response = UrlFetchApp.fetch(url, options);
            const code = response.getResponseCode();
            const text = response.getContentText();
            let json = {};
            try { json = JSON.parse(text); } catch(e) {}
            
            // Log Exit
            if (Lib.logStep) {
                const statusIcon = (code >= 200 && code < 300) ? "✅" : "❌";
                Lib.logStep("Network", `<<< END: ${endpoint} [${code}] ${statusIcon}`, "DEBUG");
            }

            if (code >= 200 && code < 300) {
                return json;
            } else {
                throw new Error(`Server Error (${code}): ${json.detail || text}`);
            }
        } catch (e) {
            if (Lib.logWarn) {
                Lib.logWarn(`Network: ${endpoint} FAILED`, e);
            }
            throw e;
        }
    };

    // --- HELPER: Init Session Logs ---
    Lib.initSessionLogs = function() {
        try {
            const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
            Lib.callServer("/logs/init", { spreadsheet_id: ssId });
        } catch(e) {
            console.error("Failed to init session logs:", e);
        }
    };

    // --- OVERRIDE: Add Article ---
    Lib.addArticleManually = function() {
        const ui = SpreadsheetApp.getUi();
        
        try {
            const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
            ui.toast("🔄 Создание нового артикула...", "Агент", 5);
            if (Lib.logStep) {
              Lib.logStep("Article", "Запрос создания артикула на сервере");
            }
            
            // We deduce project from SS ID on server
            const res = Lib.callServer("/sync/add-article", {
                article: "", // Server will generate automatically
                spreadsheet_id: ssId,
                project: "UNKNOWN" // Server resolves it
            });
            if (Lib.logStep) {
              const status = res && res.status ? res.status : "unknown";
              Lib.logStep("Article", "Артикул создан на сервере, статус: " + status);
            }
            ui.alert(`✅ Артикул создан!\nСтатус: ${res.status}`);
        } catch (e) {
            if (Lib.logWarn) {
              Lib.logWarn("Article: ошибка создания артикула", e);
            }
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
        if (Lib.logStep) {
          Lib.logStep("Delete", "Подтверждено удаление " + rows.size + " строк");
        }
        
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
            if (Lib.logStep) {
              const message = res && res.message ? res.message : "Удаление выполнено";
              Lib.logStep("Delete", message);
            }
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
            if (Lib.logStep) {
              Lib.logStep("Sync", `Синхронизация строки ${sheet.getName()}#${row} (ID=${article})`);
            }
            const res = Lib.callServer("/sync/row", {
                spreadsheet_id: ssId,
                sheet_name: sheet.getName(),
                article: article,
                project: "UNKNOWN"
            });
            if (Lib.logStep) {
              Lib.logStep("Sync", "Синхронизация завершена: " + (res && res.status ? res.status : "ok"));
            }
            ui.alert(`✅ Синхронизация завершена.`);
        } catch (e) {
            ui.alert(`❌ Ошибка: ${e.message}`);
        }
    };
    
    // --- OVERRIDE: Run Full Sync ---
    Lib.runFullSync = function() {
        // ... (existing content) ...
        const ui = SpreadsheetApp.getUi();
        const sheet = SpreadsheetApp.getActiveSheet();
        const confirm = ui.alert("Полная синхронизация", 
            `Запустить полную синхронизацию для листа "${sheet.getName()}" через сервер? Это может занять время.`, 
            ui.ButtonSet.YES_NO);
        if (confirm !== ui.Button.YES) return;

        try {
            ui.toast("Запуск полной синхронизации...");
            const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
            if (Lib.logStep) {
              Lib.logStep("Sync", "Полная синхронизация запущена для листа " + sheet.getName());
            }
            const res = Lib.callServer("/sync/full", {
                spreadsheet_id: ssId,
                source_sheet: sheet.getName(),
                project: "UNKNOWN"
            });
            if (Lib.logStep) {
              Lib.logStep("Sync", "Полная синхронизация завершена: " + (res && res.status ? res.status : "ok"));
            }
            ui.alert(`✅ Завершено.\nСтатус: ${res.status}`);
        } catch (e) {
            ui.alert(`❌ Ошибка: ${e.message}`);
        }
    };
    
    // --- BATCH PROCESS EVENT ---
    Lib.processEditEvent_Batch_ = function(e) {
        if (!e || !e.range) return;
        const range = e.range;
        const sheet = range.getSheet();
        const sheetName = sheet.getName();
        const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
        const userEmail = (e.user && e.user.getEmail()) ? e.user.getEmail() : "";
        
        const events = [];
        const startRow = range.getRow();
        const startCol = range.getColumn();
        const numRows = range.getNumRows();
        const numCols = range.getNumColumns();
        
        // Fetch data for the whole range for efficiency
        const values = range.getValues(); // 2D array
        // Fetch IDs (Col A) for rows involved
        // Optimization: fetch Col A range once
        let rowKeys = [];
        try {
            rowKeys = sheet.getRange(startRow, 1, numRows, 1).getValues().map(r => String(r[0] || "").trim());
        } catch(err) {
            // fallback if range is invalid or header row
        }
        
        // Fetch headers for columns involved
        // Optimization: fetch row 1 for columns
        let headerNames = [];
        try {
            headerNames = sheet.getRange(1, startCol, 1, numCols).getValues()[0].map(h => String(h || "").trim());
        } catch(err) {
        }
        
        for (let i = 0; i < numRows; i++) {
            const rowIndex = startRow + i;
            // Skip header row if selected
            if (rowIndex <= 1) continue;
            
            for (let j = 0; j < numCols; j++) {
                const colIndex = startCol + j;
                const value = values[i][j];
                const header = headerNames[j] || "";
                const key = rowKeys[i] || "";
                
                // If single cell, we might have e.oldValue. Batch doesn't have it easily.
                // We send what we have.
                
                events.push({
                    spreadsheet_id: ssId,
                    sheet_name: sheetName,
                    row: rowIndex,
                    col: colIndex,
                    value: value,
                    old_value: null, // Unknown in batch
                    user_email: userEmail,
                    header_name: header,
                    row_key: key
                });
            }
        }
        
        if (events.length === 0) return;
        
        // Split into chunks if too large? 
        // 500 events limit?
        if (events.length > 0) {
            Lib.logInfo(`[BatchSync] Sending ${events.length} events...`);
            try {
                // Call batch endpoint
                 Lib.callServer("/sync/batch-event", {
                    spreadsheet_id: ssId,
                    events: events
                });
            } catch (err) {
                Lib.logError("[BatchSync] Error: " + err.message);
            }
        }
    };
    
    // --- OVERRIDE: onEdit Internal ---
    // Replaces the core edit handler to use Batch Processing
    Lib.onEdit_internal_ = function(e) {
        if (!e) return;
        const sheet = e.range.getSheet();
        const row = e.range.getRow();
        
        // 1. Maintain Price Auto ID Logic
        try {
            if (typeof Lib.autoAssignPriceLineIdForRow === "function") {
                Lib.autoAssignPriceLineIdForRow(sheet, row);
            }
        } catch(err) {
            Lib.logError("PriceLogic Error", err);
        }
        
        // 2. Batch Sync Process
        try {
            Lib.processEditEvent_Batch_(e);
        } catch(err) {
            Lib.logError("BatchSync Error", err);
        }
    };

})(Lib);
