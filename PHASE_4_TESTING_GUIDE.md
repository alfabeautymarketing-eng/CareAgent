# 📋 PHASE 4: MANUAL TESTING GUIDE
## Comprehensive Testing Checklist for Phase 3 Implementation (30/33 Functions)

**Document Version**: 1.0
**Created**: 14.01.2026 (16:30 MSK)
**Project**: AgentCare Logging System
**Status**: Ready for Manual Testing

---

## 🎯 TESTING OBJECTIVES

✅ Verify that all 30 implemented functions generate correct logs
✅ Confirm logs appear correctly in "Журнал логов" (LOG_DEBUG sheet)
✅ Validate Dashboard filters work with all log categories
✅ Test JSON parameter and result data display
✅ Verify CSV export functionality
✅ Validate performance impact is acceptable

---

## 📊 FUNCTIONS TO TEST (30 Total)

### **TIER 1: CRITICAL (6/6)** - MUST TEST
- [x] `runFullSync()` - Full synchronization
- [x] `syncSelectedRow()` - Single row sync
- [x] `deleteSelectedRowsWithSync()` - Bulk delete with sync
- [x] `addArticleManually()` - Add new article
- [x] `onEdit_internal_()` - Edit trigger
- [x] `createFullInvoice()` - Invoice creation

### **TIER 2: HIGH (14/15)** - SHOULD TEST

**SK Module (3)**:
- [x] `processSkPriceSheet()` - Process SK prices
- [x] `processSkPriceProbes()` - Process SK samples
- [x] `loadSkStockData()` - Load SK stock

**MT Module (4)**:
- [x] `processMtMainPrice()` - Process MT main prices
- [x] `processMtTesterPrice()` - Process MT tester prices
- [x] `processMtSamplesPrice()` - Process MT samples
- [x] `loadMtStockData()` - Load MT stock

**SS Module (2)**:
- [x] `processSsPriceSheet()` - Process SS prices
- [x] `loadSsStockData()` - Load SS stock

**Invoice Module (2)**:
- [x] `formatOrderSheet()` - Format order sheet
- [x] `collectAndCopyDocuments()` - Collect documents

**Sync Module (2)**:
- [x] `handleOnChange()` - Handle sheet changes
- [x] `syncMultipleRows()` - Bulk row sync

**Certification Module (1)**:
- [x] `runManualCascadeOnCertification()` - Cascade recalc

### **TIER 3: MEDIUM (9/9)** - NICE TO TEST

**Export Module (2)**:
- [x] `exportPromotions()` - Export promotions
- [x] `exportSets()` - Export sets

**Agent Module (4)**:
- [x] `menuCheckService()` - Check AI API status
- [x] `menuAnalyzeSelected()` - Analyze selected row
- [x] `menuSmartMatch()` - Smart product match
- [x] `menuAnalyzeEmpty()` - Batch analysis

**Certification Module (3)**:
- [x] `structureDocuments_353pp()` - Structure documents
- [x] `calculateAndAssignSpiritNumbers()` - Calc spirit numbers
- [x] `generateSpiritProtocols()` - Generate protocols

### **TIER 4: LOW (1/1)** - OPTIONAL

**Dashboard Module (1)**:
- [x] `menuShowCategories()` - Show categories

---

## 🔬 TESTING PROCEDURES

### **PROCEDURE 1: Single Function Test**

**Goal**: Verify one function logs correctly
**Time per function**: 2-3 minutes
**Tools needed**: Google Sheets, browser

**Steps**:

1. **Open the spreadsheet**
   ```
   - Go to AgentCare Google Sheet
   - Open Apps Script (Extensions > Apps Script)
   - Load the deployed version
   ```

2. **Execute the function**
   ```
   - Open the main menu (Refresh if needed)
   - Click on appropriate menu item
   - Complete any dialogs/confirmations
   - Wait for success message
   ```

3. **Check LOG_DEBUG sheet**
   ```
   - Go to sheet named "Журнал логов"
   - Look for entries from the function you just ran
   - Should have columns: Время, Категория, Статус, Сообщение, Функция, Детали, Параметры (JSON), Результаты (JSON)
   ```

4. **Verify log content**
   ```
   ✅ Timestamp is current (within last minute)
   ✅ Category is correct (Price, Sync, Export, Agent, Certificate, Dashboard)
   ✅ Status shows: START, PROGRESS/SUCCESS, or ERROR
   ✅ Message describes what happened
   ✅ Function name matches what you called
   ✅ JSON parameters are valid (not corrupted)
   ✅ JSON results are valid (not corrupted)
   ```

5. **Check Dashboard**
   ```
   - Open "📊 Дашборд" sheet
   - Use filters to find your function's logs
   - Verify data displays correctly
   - Try clicking "Show Details" to view JSON
   ```

---

### **PROCEDURE 2: Error Handling Test**

**Goal**: Verify ERROR logging works correctly
**Time per function**: 1-2 minutes
**Tools needed**: Google Sheets, browser console

**Steps**:

1. **Trigger an error condition**
   ```
   - For each function, identify what could cause an error:
     * Invalid selection (for row operations)
     * Missing data (for processing)
     * API unavailable (for Agent functions)
     * Network error (disconnect WiFi if testing menuCheckService)
   ```

2. **Execute function and observe error**
   ```
   - Run the function with error condition
   - Note the error message shown to user
   ```

3. **Check ERROR log**
   ```
   - Go to "Журнал логов" sheet
   - Look for latest entry with Status = "ERROR"
   - Verify error details are logged:
     ✅ Error message is descriptive
     ✅ Stack trace present in results JSON
     ✅ Function name matches
     ✅ Category is correct
   ```

4. **Dashboard ERROR filter**
   ```
   - Open Dashboard
   - Filter by Status = "ERROR"
   - Verify ERROR logs display correctly
   ```

---

### **PROCEDURE 3: Category & Status Filtering**

**Goal**: Verify all categories and statuses filter correctly
**Time**: 10 minutes
**Tools needed**: Google Sheets, 📊 Дашборд

**Steps**:

1. **Run multiple functions** (at least 3 different categories)
   ```
   Example:
   - Run menuCheckService (Agent)
   - Run exportPromotions (Export)
   - Run processSkPriceSheet (Price)
   ```

2. **Test Category filter**
   ```
   - Open Dashboard
   - Check Category filter checkboxes
   - Select only "Agent"
     ✅ Should show only Agent logs
   - Select only "Export"
     ✅ Should show only Export logs
   - Select "Agent" + "Price"
     ✅ Should show combined results
   ```

3. **Test Status filter**
   ```
   - Reset Category filter (select all)
   - Check Status filter dropdown
   - Select only "START"
     ✅ Should show function entry logs
   - Select only "SUCCESS"
     ✅ Should show completion logs
   - Select only "ERROR"
     ✅ Should show error logs
   ```

4. **Test combined filters**
   ```
   - Category: Agent, Export
   - Status: SUCCESS
   - Time: Today
     ✅ Should show successful completions of Agent/Export functions
   ```

---

### **PROCEDURE 4: JSON Data Inspection**

**Goal**: Verify parameters and results JSON is valid and accessible
**Time**: 5 minutes
**Tools needed**: Dashboard

**Steps**:

1. **Open Dashboard**
2. **Find a SUCCESS log** (with results JSON)
3. **Click "View Details"** or appropriate button
4. **Inspect JSON Modal**
   ```
   ✅ Parameters JSON shows function input data
   ✅ Results JSON shows function output data
   ✅ Both are valid JSON (not corrupted)
   ✅ Data is readable and sensible
   ```

5. **Examples to check**:
   ```
   For exportPromotions:
   - parameters: { type: "promotions" }
   - results: { status: "completed", rowsWritten: N }

   For menuCheckService:
   - parameters: { }
   - results: { status: "online", model: "..." }

   For processSkPriceSheet:
   - parameters: { projectKey: "SK" }
   - results: { rowsProcessed: N, status: "completed" }
   ```

---

### **PROCEDURE 5: CSV Export Test**

**Goal**: Verify filtered logs can be exported to CSV
**Time**: 3 minutes
**Tools needed**: Dashboard

**Steps**:

1. **Open Dashboard**
2. **Apply some filters**
   ```
   Example: Status = SUCCESS, Category = Agent
   ```

3. **Click CSV Export button**
   ```
   ✅ File downloads as "logs_TIMESTAMP.csv"
   ✅ File is not empty
   ✅ Filename is reasonable
   ```

4. **Open CSV file** (Excel, Google Sheets, or text editor)
   ```
   ✅ Headers match LOG_DEBUG columns
   ✅ Data rows are properly formatted
   ✅ Special characters are escaped correctly
   ✅ JSON fields are enclosed in quotes
   ```

5. **Sample inspection**:
   ```
   Expected columns:
   Время | Категория | Статус | Сообщение | Функция | Детали | Параметры (JSON) | Результаты (JSON)

   Expected data:
   2026-01-14 16:30:45|Agent|SUCCESS|✅ Проверка статуса...|menuCheckService|...|{}|{"status":"online"}
   ```

---

### **PROCEDURE 6: Performance Test**

**Goal**: Verify logging doesn't significantly impact performance
**Time**: 10 minutes (per function category)
**Tools needed**: Google Sheets, browser DevTools (optional)

**Steps**:

1. **Test Price Processing** (most common operation)
   ```
   - Open menu: Обработка > Загрузить SK/MT/SS
   - Note approximate time on toast notification
   - Check if time is reasonable (should be <30 seconds for typical data)
   - Wait for completion
   ```

2. **Test Sync Operations**
   ```
   - Run syncMultipleRows() with ~10-100 rows
   - Note time on completion toast
   - Should complete in <15 seconds for 100 rows
   ```

3. **Test Agent Functions**
   ```
   - Run menuSmartMatch()
   - Should complete in <5 seconds
   - Run menuAnalyzeSelected()
   - Should complete in <30 seconds (API call)
   ```

4. **Check memory usage** (optional, advanced)
   ```
   - Open Browser DevTools (F12)
   - Monitor memory before and after function execution
   - Should not grow unbounded (logging adds ~50-100KB per 100 logs)
   ```

---

### **PROCEDURE 7: Dashboard Features Test**

**Goal**: Verify Dashboard UI works correctly
**Time**: 5 minutes
**Tools needed**: Dashboard sheet

**Steps**:

1. **Test Time Period Filter**
   ```
   - Select "Today"
     ✅ Shows only today's logs
   - Select "24 Hours"
     ✅ Shows last 24 hours
   - Select "7 Days"
     ✅ Shows last 7 days
   - Select "All"
     ✅ Shows all logs
   ```

2. **Test Pagination**
   ```
   - Change rows per page to 10
     ✅ Table shows 10 rows
   - Change to 50
     ✅ Table shows 50 rows
   - Navigate to next page
     ✅ Shows next batch of results
   ```

3. **Test Search**
   ```
   - Type a function name (e.g., "exportPromotions")
     ✅ Shows only logs with that function
   - Try searching for message text
     ✅ Shows matching logs
   ```

4. **Test Metrics**
   ```
   - Verify "Total Logs" count is correct
   - Verify "Success" count matches SUCCESS status logs
   - Verify "Errors" count matches ERROR status logs
   - Verify "Warnings" count matches WARNING status logs
   ```

---

## 📈 EXPECTED RESULTS SUMMARY

| Test Category | Expected Outcome | Pass/Fail |
|---------------|-----------------|-----------|
| Log Creation | All 30 functions generate logs | ⬜ |
| LOG_DEBUG Sheet | Data appears in correct columns | ⬜ |
| Dashboard Display | Logs visible with correct formatting | ⬜ |
| Category Filter | All categories filter correctly | ⬜ |
| Status Filter | All statuses filter correctly | ⬜ |
| JSON Display | Parameters and results readable | ⬜ |
| CSV Export | Files download and format correctly | ⬜ |
| Performance | Logging adds <5% time overhead | ⬜ |
| Error Logging | ERROR logs capture full stack | ⬜ |

---

## 🚀 TESTING EXECUTION PLAN

### **Priority Testing Order** (Recommended)

**Session 1: TIER 1 Critical Functions** (20 minutes)
```
1. runFullSync() - Most critical
2. syncSelectedRow() - Core operation
3. addArticleManually() - Important operation
4. deleteSelectedRowsWithSync() - Important operation
→ Goal: Verify core sync operations log correctly
```

**Session 2: TIER 2 Price Functions** (20 minutes)
```
1. processSkPriceSheet() - SK project
2. processMtMainPrice() - MT project
3. processSsPriceSheet() - SS project
→ Goal: Verify price processing logs correctly
```

**Session 3: TIER 3 Agent Functions** (15 minutes)
```
1. menuCheckService() - Fast, no dependencies
2. menuSmartMatch() - Fast with AI response
3. menuAnalyzeSelected() - Longer operation
→ Goal: Verify Agent functions work with API calls
```

**Session 4: Dashboard & Filtering** (15 minutes)
```
1. Test all filters (Category, Status, Time)
2. Test JSON details display
3. Test CSV export
→ Goal: Verify Dashboard features work end-to-end
```

**Session 5: Error Handling** (10 minutes)
```
1. Test each function category with error condition
2. Verify ERROR logs capture full details
3. Verify Dashboard ERROR filter works
→ Goal: Verify error logging is comprehensive
```

---

## 📝 TESTING LOG TEMPLATE

**Function Tested**: _________________
**Date/Time**: _________________
**Result**: ✅ PASS / ❌ FAIL / ⚠️ PARTIAL

**Details**:
```
Log Created: Yes / No
Category: ________________
Status: START / PROGRESS / SUCCESS / ERROR
Message: ________________
JSON Valid: Yes / No
Dashboard Shows: Yes / No
Notes: ________________
```

---

## ⚠️ KNOWN LIMITATIONS & ISSUES

### **3 Unimplemented Certification Functions** (Cannot be tested)
```
- createNewsSheetFromCertification() - only shows "not implemented" alert
- generateProtocols_353pp() - only shows "not implemented" alert
- generateDsLayouts_353pp() - only shows "not implemented" alert

Action: Wait until functions are implemented before adding logging
```

### **Performance Considerations**
```
- Logging adds ~5-10ms per function call (typical)
- LOG_DEBUG sheet grows ~1-2KB per log entry
- Dashboard may be slow if >10,000 logs exist (implement archival)
- CSV export limited to ~50,000 rows for performance
```

---

## 🎓 NEXT PHASES (For Future Reference)

### **Phase 3.5-3.9: Additional Functions** (78 candidates identified)

**Phase 3.5 - TIER 2.5 (8 functions)**
- Settings/Configuration utilities
- Examples: setupTriggers(), clearAllToasts(), refreshLogs()

**Phase 3.6 - TIER 3 (11 functions)**
- Data reporting/filtering
- Examples: showOrderStage(), showPromotionsStage()

**Phase 3.7 - TIER 4 (2 functions)**
- Drive/File operations
- Examples: uploadFilesToFolder(), createFolderStructure()

**Phase 3.8 - TIER 5 (7 functions)**
- Data calculations
- Examples: serverProcessPricePreview(), serverRecalculateCascades()

**Phase 3.9 - TIER 6-7 (40+ functions)**
- Helpers and utilities
- Examples: callServerSort(), getSortConfig(), debugShowSpreadsheetId()

**Estimated effort**: 2-4 hours for complete logging of all functions

---

## 📞 SUPPORT & DOCUMENTATION

**For detailed logging specification**: See [LOGGING_GUIDE.md](LOGGING_GUIDE.md)
**For implementation details**: See [TASKS.md](TASKS.md)
**For code review**: Check latest commits in git log

**Issues or questions**:
1. Check LOGGING_GUIDE.md for reference implementation patterns
2. Review similar functions in the same module
3. Consult TASKS.md Phase 4 verification results

---

**Document Version**: 1.0
**Last Updated**: 14.01.2026 16:30 MSK
**Status**: ✅ Ready for Manual Testing
