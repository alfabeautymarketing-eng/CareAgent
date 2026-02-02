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
    Lib.callServer = function(endpoint, payload, method) {
        const SERVER_URL = "http://46.226.167.153:8000";
        const httpMethod = (method || "post").toUpperCase();
        
        // Log Entry
        if (Lib.logStep) {
            Lib.logStep("Network", `>>> START: ${endpoint} [${httpMethod}]`, "DEBUG");
        }
        
        try {
            const options = {
                method: httpMethod,
                contentType: "application/json",
                payload: httpMethod === "GET" ? undefined : JSON.stringify(payload || {}),
                headers: { "ngrok-skip-browser-warning": "true" },
                muteHttpExceptions: true
            };
            
            let finalEndpoint = endpoint;
            if (finalEndpoint.startsWith("//")) {
                // Если начинается с //, считаем это путем от корня (без /api/v1)
                finalEndpoint = finalEndpoint.substring(1);
            } else if (!finalEndpoint.startsWith("/api/v1")) {
                if (!finalEndpoint.startsWith("/")) finalEndpoint = "/" + finalEndpoint;
                finalEndpoint = "/api/v1" + finalEndpoint;
            }

            let url = `${SERVER_URL}${finalEndpoint}`;

            // GET: добавляем query-параметры вместо тела
            if (httpMethod === "GET" && payload && typeof payload === "object") {
                const qs = Object.entries(payload)
                  .filter(([,v]) => v !== undefined && v !== null)
                  .map(([k,v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
                  .join("&");
                if (qs) url += (url.includes("?") ? "&" : "?") + qs;
            }

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

            // --- 1. Smart Match Prompt ---
            const prompt = ui.prompt("Добавить артикул", "Введите название продукта или артикул производителя для сопоставления:", ui.ButtonSet.OK_CANCEL);
            if (prompt.getSelectedButton() !== ui.Button.OK) return;
            const productName = prompt.getResponseText().trim();

            let matchedDetails = null;
            if (productName) {
                ui.toast("🎯 Ищем совпадения в базе...", "Smart Match", 5);
                // serverSmartMatch defined in 00_GlobalApiBridge.js calls callServerSmartMatch
                const matchRes = Lib.serverSmartMatch ? Lib.serverSmartMatch(productName) : null;
                if (matchRes && matchRes.match_found && matchRes.confidence > 80) {
                    const details = matchRes.matched_product_details;
                    const confirm = ui.alert("Найдено совпадение!", 
                        `Найдено: "${details.name_rus_ds}" (${matchRes.confidence}%)\n\nЗаполнить данные сертификации автоматически?`, 
                        ui.ButtonSet.YES_NO);
                    if (confirm === ui.Button.YES) {
                        matchedDetails = details;
                    }
                }
            }

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

            const newArticle = res && res.article ? res.article : null;
            
            if (Lib.logStep) {
              const status = res && res.status ? res.status : "unknown";
              Lib.logStep("Article", `Артикул ${newArticle} создан на сервере, статус: ${status}`);
            }

            // --- 2. Apply Matched Data ---
            if (newArticle && matchedDetails) {
                ui.toast("📝 Заполняем данные сертификации...", "Smart Match", 5);
                Lib.applyMatchedDataToCertification(newArticle, matchedDetails);
            }

            ui.alert(`✅ Артикул создан: ${newArticle}\nСтатус: ${res.status}`);
        } catch (e) {
            if (Lib.logWarn) {
              Lib.logWarn("Article: ошибка создания артикула", e);
            }
            ui.alert(`❌ Ошибка: ${e.message}`);
        }
    };

    /**
     * Заполняет данные сертификации для нового артикула на основе найденного совпадения.
     */
    Lib.applyMatchedDataToCertification = function(article, details) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName("Сертификация");
        if (!sheet) return;

        const lastCol = sheet.getLastColumn();
        const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        const rowData = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues(); // ID в колонке A
        
        let targetRow = -1;
        for (let i = 0; i < rowData.length; i++) {
            if (String(rowData[i][0]).trim() === String(article).trim()) {
                targetRow = i + 1;
                break;
            }
        }

        if (targetRow === -1) {
            if (Lib.logWarn) Lib.logWarn(`ApplyMatch: Артикул ${article} не найден на листе Сертификация`);
            return;
        }

        const fieldMap = {
            "наименования англ по дс": "name_eng_ds",
            "наименования рус по дс": "name_rus_ds",
            "категория": "категория",
            "вид продукции": "Вид продукции",
            "код тн вэд": "Код ТН ВЭД",
            "дс": "ДС",
            "номер дс": "номер ДС",
            "дс до": "ДС до",
            "протокол дс": "Протокол ДС",
            "протокол доп": "Протокол доп",
            "inci": "INCI",
            "coa": "COA"
        };

        let updateCount = 0;
        headers.forEach((h, idx) => {
            const hLower = String(h).toLowerCase().trim().replace(/ё/g, "е");
            if (fieldMap[hLower]) {
                const val = details[fieldMap[hLower]];
                if (val !== undefined && val !== null && val !== "") {
                    sheet.getRange(targetRow, idx + 1).setValue(val);
                    updateCount++;
                }
            }
        });

        // Запуск каскада для обновления производных полей
        if (updateCount > 0 && (typeof Lib._runCertificationCascade === "function")) {
            Lib._runCertificationCascade(sheet, targetRow, "Наименования рус по ДС");
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

    // --- OVERRIDE: Collect Documents ---
    Lib.collectAndCopyDocuments = function() {
        const ui = SpreadsheetApp.getUi();
        const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
        const brandPrefix = (Lib.CONFIG && Lib.CONFIG.SETTINGS && Lib.CONFIG.SETTINGS.BRAND_PREFIX) || "MT";

        try {
            ui.toast("Сбор документов на сервере...", "Агент", 30);
            const res = Lib.callServer("/documents/collect", {
                spreadsheet_id: ssId,
                target_sheet: "Для инвойса",
                brand_prefix: brandPrefix
            });

            ui.alert("Операция завершена!", res.data.message, ui.ButtonSet.OK);
        } catch (e) {
            ui.alert("Ошибка сбора документов", e.message, ui.ButtonSet.OK);
        }
    };

    Lib.collectAndCopyDocuments_proxy = function() {
        return Lib.collectAndCopyDocuments();
    };

    // --- ARCHIVED: LOG SHEET FUNCTIONS (removed - all logging is server-side now) ---
    // Previously: quickCleanLogSheet(), recreateLogSheet(), recreateDebugLogSheet()
    // These tried to manage local Google Sheets logging
    // Now: All logging is server-side only

    // --- OVERRIDE: Archive Logs ---
    Lib.manualArchiveLogs_proxy = function() {
        const ui = SpreadsheetApp.getUi();
        const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
        
        try {
            ui.toast("Архивация логов на сервере...", "Агент", 30);
            const res = Lib.callServer("/logs/archive", {
                spreadsheet_id: ssId
            });
            
            ui.alert("Архивация завершена!", `Записано ${res.data.total_rows} строк в ${res.data.archive_name}`, ui.ButtonSet.OK);
        } catch (e) {
            ui.alert("Ошибка архивации", e.message, ui.ButtonSet.OK);
        }
    };

    // --- OVERRIDE: Archive Status ---
    Lib.showArchiveStatus_proxy = function() {
        const ui = SpreadsheetApp.getUi();
        const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
        
        try {
            const res = Lib.callServer("/logs/archive/status", {
                spreadsheet_id: ssId
            }, "GET");
            
            if (!res || !res.data) throw new Error("Нет данных от сервера");
            
            const data = res.data;
            let status = `Архив: ${data.archive_name}\n`;
            status += `Существует: ${data.exists ? "✅ Да" : "❌ Нет"}\n`;
            if (data.exists && data.last_modified) {
                status += `Изменен: ${data.last_modified}`;
            }
            
            ui.alert("Статус архива", status, ui.ButtonSet.OK);
        } catch (e) {
            ui.alert("Ошибка статуса", e.message, ui.ButtonSet.OK);
        }
    };

    // --- REMOVED: Local sheet navigation functions (now server-side only) ---
    // Previously: showLogJournal(), showSyncJournal() tried to navigate to local sheets
    // Now: All logging is server-side, access via server URLs

    /**
     * Проверка статуса всех внешних сервисов с детальным логированием.
     */
    Lib.showAllServicesStatus = function() {
        const ui = SpreadsheetApp.getUi();
        SpreadsheetApp.getActive().toast("Проверка статуса всех сервисов...", "Экосистема", 5);

        // START: Логируем начало процесса
        if (Lib.logWithEmoji) {
            Lib.logWithEmoji(
                "Запуск полной диагностики системы", 
                "INFO", 
                "🚀", 
                "showAllServicesStatus", 
                "Начата проверка всех внешних сервисов", 
                "System", 
                "START"
            );
        }

        try {
            const checkResults = {
                server: { status: "PENDING", message: "" },
                gemini: { status: "PENDING", message: "" },
                drive: { status: "PENDING", message: "" }
            };

            // 1. Проверка основного сервера
            if (Lib.logWithEmoji) {
                Lib.logWithEmoji("Проверка связи с основным сервером...", "DEBUG", "⏳", "showAllServicesStatus", "Вызов //health", "API", "PROGRESS");
            }

            let serverStatus = "🔴 Ошибка";
            try {
                const res = Lib.callServer("//health", {}, "GET");
                if (res && res.status === "healthy") {
                    serverStatus = "🟢 OK";
                    checkResults.server = { status: "OK", message: "Healthy", details: res };
                    if (Lib.logWithEmoji) {
                        Lib.logWithEmoji("Сервер доступен", "INFO", "✅", "showAllServicesStatus", "Ответ HEALTHY получен", "API", "SUCCESS", null, res);
                    }
                } else {
                    checkResults.server = { status: "ERROR", message: "Invalid Response", details: res };
                    if (Lib.logWithEmoji) {
                        Lib.logWithEmoji("Некорректный ответ сервера", "WARN", "⚠️", "showAllServicesStatus", "Статус не healthy", "API", "WARNING", null, res);
                    }
                }
            } catch(e) { 
                serverStatus = "🔴 " + e.message;
                checkResults.server = { status: "ERROR", message: e.message };
                if (Lib.logWithEmoji) {
                    Lib.logWithEmoji("Ошибка связи с сервером", "ERROR", "❌", "showAllServicesStatus", e.message, "API", "ERROR", null, {error: e.toString()});
                }
            }
            
            // 2. Проверка Gemini/AI
            if (Lib.logWithEmoji) {
                Lib.logWithEmoji("Проверка статуса Gemini AI...", "DEBUG", "⏳", "showAllServicesStatus", "Вызов /ai/status", "API", "PROGRESS");
            }

            let geminiStatus = "🔴 Ошибка";
            try {
                const res = Lib.callServer("/ai/status", {}, "GET");
                if (res && res.status === "online") {
                    geminiStatus = "🟢 OK";
                    checkResults.gemini = { status: "OK", message: "Online", model: res.model };
                    if (Lib.logWithEmoji) {
                        Lib.logWithEmoji("Gemini AI доступен", "INFO", "✅", "showAllServicesStatus", `Модель: ${res.model}`, "API", "SUCCESS", null, res);
                    }
                } else {
                    checkResults.gemini = { status: "ERROR", message: "Offline/Error", details: res };
                    if (Lib.logWithEmoji) {
                        Lib.logWithEmoji("Gemini AI вернул ошибку", "WARN", "⚠️", "showAllServicesStatus", `Статус: ${res.status}, Ошибка: ${res.error}`, "API", "WARNING", null, res);
                    }
                }
            } catch(e) { 
                geminiStatus = "🔴 " + e.message;
                checkResults.gemini = { status: "ERROR", message: e.message };
                if (Lib.logWithEmoji) {
                    Lib.logWithEmoji("Ошибка проверки Gemini AI", "ERROR", "❌", "showAllServicesStatus", e.message, "API", "ERROR", null, {error: e.toString()});
                }
            }
            
            // 3. Проверка Google Drive
            if (Lib.logWithEmoji) {
                Lib.logWithEmoji("Проверка доступа к Google Drive...", "DEBUG", "⏳", "showAllServicesStatus", "Получение корневой папки", "Storage", "PROGRESS");
            }

            let driveStatus = "🟢 OK";
            try {
                const root = DriveApp.getRootFolder();
                checkResults.drive = { status: "OK", message: "Access Granted", name: root.getName() };
                if (Lib.logWithEmoji) {
                    Lib.logWithEmoji("Google Drive доступен", "INFO", "✅", "showAllServicesStatus", "Корневая папка получена", "Storage", "SUCCESS");
                }
            } catch(e) { 
                driveStatus = "🔴 " + e.message;
                checkResults.drive = { status: "ERROR", message: e.message };
                if (Lib.logWithEmoji) {
                    Lib.logWithEmoji("Ошибка доступа к Google Drive", "ERROR", "❌", "showAllServicesStatus", e.message, "Storage", "ERROR", null, {error: e.toString()});
                }
            }
            
            const resultsSummary = `Основной сервер: ${serverStatus}\nGemini AI: ${geminiStatus}\nGoogle Drive: ${driveStatus}`;
            ui.alert("Статус сервисов", resultsSummary, ui.ButtonSet.OK);
            
            // SUCCESS: Итоговый лог
            if (Lib.logWithEmoji) {
                Lib.logWithEmoji(
                    "Диагностика завершена", 
                    "INFO", 
                    "🏁", 
                    "showAllServicesStatus", 
                    resultsSummary.replace(/\n/g, ", "), 
                    "System", 
                    "SUCCESS", 
                    null, 
                    checkResults
                );
            }

        } catch(e) {
            ui.alert("Ошибка при проверке", e.message, ui.ButtonSet.OK);
            // ERROR: Итоговый лог
            if (Lib.logWithEmoji) {
                Lib.logWithEmoji(
                    "Критическая ошибка диагностики", 
                    "ERROR", 
                    "📛", 
                    "showAllServicesStatus", 
                    e.message, 
                    "System", 
                    "ERROR", 
                    null, 
                    {error: e.toString()}
                );
            }
        }
    };

    // --- REMOVED: Local dashboard navigation (now server-side only) ---
    // Previously: openLogDashboard() tried to create/open local sheet
    // Now: Dashboard is accessed via server URL

    /**
     * Opens the server-hosted logging dashboard
     */
    Lib.openServerDashboard = function() {
        const ui = SpreadsheetApp.getUi();
        const serverUrl = (Lib.CONFIG && Lib.CONFIG.SERVER_URL) || "http://localhost:5000";
        const dashboardUrl = serverUrl + "/logs-ui";

        ui.alert(
            "📊 Открыть Дашборд",
            "Дашборд логов доступен по адресу:\n\n" + dashboardUrl + "\n\nСкопируйте ссылку в браузер.",
            ui.ButtonSet.OK
        );

        // Log the action
        if (Lib.logWithEmoji) {
            Lib.logWithEmoji(
                "Запрос открытия дашборда",
                "INFO",
                "",
                "openServerDashboard",
                "Пользователь открыл дашборд с сервера",
                "System",
                "SUCCESS",
                {},
                { url: dashboardUrl }
            );
        }
    };

    /**
     * Opens the server-hosted logs journal
     */
    Lib.openServerLogsJournal = function() {
        const ui = SpreadsheetApp.getUi();
        const serverUrl = (Lib.CONFIG && Lib.CONFIG.SERVER_URL) || "http://localhost:5000";
        const logsUrl = serverUrl + "/logs-ui";

        ui.alert(
            "📝 Открыть Журнал логов",
            "Журнал логов доступен по адресу:\n\n" + logsUrl + "\n\nСкопируйте ссылку в браузер.",
            ui.ButtonSet.OK
        );

        // Log the action
        if (Lib.logWithEmoji) {
            Lib.logWithEmoji(
                "Запрос открытия журнала логов",
                "INFO",
                "",
                "openServerLogsJournal",
                "Пользователь открыл журнал логов с сервера",
                "System",
                "SUCCESS",
                {},
                { url: logsUrl }
            );
        }
    };

    /**
     * Opens the server-hosted sync journal
     */
    Lib.openServerSyncJournal = function() {
        const ui = SpreadsheetApp.getUi();
        const serverUrl = (Lib.CONFIG && Lib.CONFIG.SERVER_URL) || "http://localhost:5000";
        const syncUrl = serverUrl + "/logs-ui";

        ui.alert(
            "🔄 Открыть Журнал синхронизации",
            "Журнал синхронизации доступен по адресу:\n\n" + syncUrl + "\n\nСкопируйте ссылку в браузер.",
            ui.ButtonSet.OK
        );

        // Log the action
        if (Lib.logWithEmoji) {
            Lib.logWithEmoji(
                "Запрос открытия журнала синхро",
                "INFO",
                "",
                "openServerSyncJournal",
                "Пользователь открыл журнал синхронизации с сервера",
                "System",
                "SUCCESS",
                {},
                { url: syncUrl }
            );
        }
    };

    // --- PROXY FUNCTIONS FOR MENU ---
    Lib.showAllServicesStatus_proxy = function() {
        return Lib.showAllServicesStatus();
    };

    // --- NEW: Server-based dashboard and journal accessors ---
    // These functions open server URLs instead of local sheets

    Lib.openServerDashboard_proxy = function() {
        return Lib.openServerDashboard();
    };

    Lib.openServerLogsJournal_proxy = function() {
        return Lib.openServerLogsJournal();
    };

    Lib.openServerSyncJournal_proxy = function() {
        return Lib.openServerSyncJournal();
    };

    // --- ARCHIVED: LOG_DEBUG Sheet Structure (All removed - server-side logging only) ---
    // NOTE: All local Google Sheets logging has been removed
    // Logging is now exclusively server-side
    // Dashboard and journals are accessed via server URLs

})(Lib);

// ============= ARCHIVED: GLOBAL PROXY FUNCTIONS (removed - server-side logging only) =============
// Previously: callServerLogCommand(), callServerRecreateDebugLog(), Lib.refreshLogs()
// These tried to manage local Google Sheets logging
// Now: All logging is server-side only - access via server URLs
