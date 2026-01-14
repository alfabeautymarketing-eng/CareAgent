# 🛣️ ROADMAP: Phase 3.5+ Function Logging Expansion
## Comprehensive Plan for 78 Additional Medium/Low Priority Functions

**Document Version**: 1.0
**Created**: 14.01.2026 (16:40 MSK)
**Project**: AgentCare Comprehensive Logging System
**Status**: Planning & Recommendations

---

## 📊 EXECUTIVE SUMMARY

**Current Status**: 30/108 functions logged (28%)
- Phase 3.1-3.4: 30 HIGH/CRITICAL functions ✅ COMPLETE
- Phase 3.5-3.9: 78 MEDIUM/LOW functions (Planned)

**Additional Logging Candidates**: 78 functions across all modules
- Tier 2.5 (Medium-High): 8 functions (Settings/Configuration)
- Tier 3 (Medium Data): 11 functions (Reporting/Filtering)
- Tier 4 (File Ops): 2 functions (Drive operations)
- Tier 5 (Calculations): 7 functions (Data transformations)
- Tier 6-7 (Low Helpers): 40+ functions (Utilities/HTTP/Caching)

**Estimated Effort**: 2-4 hours for complete implementation
**Expected Token Cost**: 15,000-25,000 tokens

---

## 🎯 PHASE 3.5: TIER 2.5 (Settings & Configuration) - 8 Functions

### Overview
Functions that manage system configuration, triggers, and settings. These are critical for debugging configuration issues but don't directly process business data.

**Priority**: HIGH (do soon)
**Functions**: 8
**Effort**: 1-1.5 hours
**Files affected**: 00_GlobalApiBridge.js (7), z-server-overrides.js (1)

---

### **Functions to Log**

#### 1. `clearAllToasts()` [00_GlobalApiBridge.js:260]
**Purpose**: Force-close all active toast notifications
**Severity**: Utility
**When called**: Manually via menu when toast stuck

```javascript
// Implementation pattern:
function clearAllToasts() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Запуск очистки всех уведомлений", "INFO", "", "clearAllToasts",
      "Закрытие активных toast-уведомлений", "System", "START", {}, null);
  }
  try {
    // ... existing code ...
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Toast-уведомления очищены", "INFO", "", "clearAllToasts",
        "Все активные уведомления закрыты", "System", "SUCCESS", {}, { status: "cleared" });
    }
  } catch(e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при очистке уведомлений", "ERROR", "", "clearAllToasts",
        e && e.message ? e.message : String(e), "System", "ERROR", {}, { error: e ? e.toString() : "Unknown" });
    }
  }
}
```

---

#### 2. `setupTriggers()` [00_GlobalApiBridge.js:307]
**Purpose**: Install installable triggers for onOpen/onChange
**Severity**: Configuration
**When called**: Once during initial setup

**Logging Category**: Settings
**Expected Status Pattern**: START → SUCCESS/ERROR

---

#### 3. `refreshLogs()` [00_GlobalApiBridge.js:352]
**Purpose**: Refresh log sheet structure from server
**Severity**: Maintenance
**When called**: Manually or on initialization

**Logging Category**: System
**Expected Status Pattern**: START → SUCCESS/ERROR

---

#### 4. `showSyncConfigDialog()` [00_GlobalApiBridge.js:315]
**Purpose**: Shows deprecated sync rules manager (wrapper)
**Severity**: UI
**When called**: User clicks menu item

**Logging Category**: Menu
**Expected Status Pattern**: START → SUCCESS

---

#### 5. `showSyncRulesManagerDialog()` [00_GlobalApiBridge.js:320]
**Purpose**: Opens modern sync rules CRUD manager
**Severity**: UI
**When called**: User clicks menu item (modern version)

**Logging Category**: Menu
**Expected Status Pattern**: START → SUCCESS

---

#### 6. `showExternalDocManagerDialog()` [00_GlobalApiBridge.js:327]
**Purpose**: Opens external documents manager
**Severity**: UI
**When called**: User clicks menu item

**Logging Category**: Menu
**Expected Status Pattern**: START → SUCCESS

---

#### 7. `runAllTests()` [00_GlobalApiBridge.js:336]
**Purpose**: Runs all validation/test suites
**Severity**: Testing
**When called**: Developer/Admin command

**Logging Category**: System
**Expected Status Pattern**: START → PROGRESS (multiple) → SUCCESS/ERROR

---

#### 8. `clearTestResults()` [00_GlobalApiBridge.js:343]
**Purpose**: Clears test result data
**Severity**: Maintenance
**When called**: Manual cleanup

**Logging Category**: System
**Expected Status Pattern**: START → SUCCESS/ERROR

---

## 🎯 PHASE 3.6: TIER 3 (Data Display/Reporting) - 11 Functions

### Overview
Functions that fetch and display data for reporting. These don't modify data but provide transparency into what users are viewing.

**Priority**: MEDIUM (do after 3.5)
**Functions**: 11
**Effort**: 1.5-2 hours
**Files affected**: 00_GlobalApiBridge.js

---

### **Functions to Log**

#### Client Functions (5)
1. `showAllOrderData()` [547] - Display all orders
2. `showOrderStage()` [554] - Display orders at stage
3. `showPromotionsStage()` [561] - Display promotions at stage
4. `showSetStage()` [568] - Display sets at stage
5. `showPriceStage()` [575] - Display price data at stage

#### Server Functions (5)
6. `serverShowAllOrderData()` [980]
7. `serverShowOrderStage()` [990]
8. `serverShowPromotionsStage()` [1000]
9. `serverShowSetStage()` [1010]
10. `serverShowPriceStage()` [1020]

#### Helper (1)
11. `_callServerOrderFilter(stage)` [1031]

**Logging Pattern**:
```javascript
Lib.logWithEmoji("Запрос данных по стадии заказа", "INFO", "", "showOrderStage",
  "Загрузка заказов для текущей стадии", "Dashboard", "START", { stage: stage }, null);
// ... fetch data ...
Lib.logWithEmoji("Данные стадии заказа получены", "INFO", "", "showOrderStage",
  "Получено " + orders.length + " заказов", "Dashboard", "SUCCESS", { stage: stage },
  { count: orders.length, status: "loaded" });
```

---

## 🎯 PHASE 3.7: TIER 4 (File & Drive Operations) - 2 Functions

### Overview
Functions that manage files and folders in Google Drive.

**Priority**: MEDIUM
**Functions**: 2
**Effort**: 0.5 hours
**Files affected**: 00_GlobalApiBridge.js

---

### **Functions to Log**

1. **`uploadFilesToFolder()` [743]**
   - Purpose: Upload files to Drive folder
   - Category: File Operations
   - Expected logs: START → SUCCESS/ERROR (one per file)

2. **`createFolderStructure()` [750]**
   - Purpose: Create standard folder hierarchy
   - Category: File Operations
   - Expected logs: START → SUCCESS/ERROR

---

## 🎯 PHASE 3.8: TIER 5 (Data Transformations) - 7 Functions

### Overview
Functions that perform calculations, previews, and data transformations.

**Priority**: MEDIUM
**Functions**: 7
**Effort**: 1.5-2 hours
**Files affected**: 00_GlobalApiBridge.js

---

### **Functions to Log**

1. **`addNewYearColumnsToPriceDynamics()` [495]**
   - Category: Price
   - Status: START → SUCCESS/ERROR

2. **`reorderAuxiliarySheets()` [758]**
   - Category: System
   - Status: START → SUCCESS/ERROR

3. **`serverProcessPricePreview(project, mode)` [845]**
   - Category: Price
   - Status: START → PROGRESS → SUCCESS/ERROR

4. **`serverSmartMatch(productName)` [854]**
   - Category: Agent
   - Status: START → SUCCESS/WARNING/ERROR

5. **`serverRecalculateCascades()` [870]**
   - Category: Certificate
   - Status: START → PROGRESS → SUCCESS/ERROR

6. **`serverProcessCascadeRow(row, column, value)` [917]**
   - Category: Certificate
   - Status: START → SUCCESS/ERROR

7. **`callServerCascade(params)` [938]**
   - Category: Certificate
   - Status: START → SUCCESS/ERROR

---

## 🎯 PHASE 3.9: TIER 6-7 (Helpers & Utilities) - 40+ Functions

### Overview
HTTP, caching, configuration, and navigation helpers. These provide detailed tracing for debugging.

**Priority**: LOW (nice to have)
**Functions**: 40+
**Effort**: 2-3 hours
**Files affected**: Client.gs (30+), z-server-overrides.js (10+)

---

### **Subcategories**

#### **Sorting Functions (5)**
- `sortByManufacturer()`
- `sortByPrice()`
- `structureSortByManufacturer()`
- `structureSortByPrice()`
- `callServerStructureSort(mode)`

#### **Configuration Functions (10)**
- `getSortConfig(forceRefresh)`
- `getStoredGeminiKey()`
- `getStoredGeminiModel()`
- `saveGeminiSettings(apiKey, model)`
- `initGeminiFromStorage()`
- `setupGeminiComplete()`
- `getMenuConfigCacheKey(spreadsheetId)`
- `saveMenuConfigToCache(spreadsheetId, config)`
- `getCachedMenuConfig(spreadsheetId)`
- `showGeminiSettings()`

#### **Utility Functions (10)**
- `getSheetsList()`
- `getSheetColumns(sheetName)`
- `getExternalDocsList()`
- `getExternalSheetsList(docId)`
- `getExternalSheetColumns(docId, sheetName)`
- `reorderSheets()`
- `reorderSheetsSilent()`
- `checkServerStatus()`
- `callServerLoadFunctions()`
- `callServerSyncEvent(payload)`

#### **Debug Functions (3)**
- `debugShowSpreadsheetId()`

#### **Internal Helpers (10+)**
- `Lib.callServer()`
- `Lib.initSessionLogs()`
- `Lib.applyMatchedDataToCertification()`
- `Lib.processEditEvent_Batch_()`
- `Lib.ensureLogDebugSheetStructure()`
- `Lib.showLogJournal()`
- `Lib.showSyncJournal()`
- `Lib.openLogDashboard()`
- Plus additional helpers...

**Recommended Approach**:
- Group by functionality (Sorting, Configuration, etc.)
- Add in batches of 5-10 similar functions per commit
- Use simplified logging pattern (less detailed than critical functions)

---

## 📈 IMPLEMENTATION ROADMAP

### **Week 1: Phase 3.5 (Settings)**
```
Day 1: clearAllToasts(), setupTriggers()
Day 2: refreshLogs(), Dialog functions (4)
Day 3: Testing & fixes
```

### **Week 2: Phase 3.6 (Reporting)**
```
Day 1-2: Client functions (5)
Day 3: Server functions (5) + helper
Day 4: Testing & validation
```

### **Week 3: Phase 3.7-3.8 (Files & Calculations)**
```
Day 1: uploadFilesToFolder(), createFolderStructure()
Day 2: Price/Cascade functions (7)
Day 3: Testing & integration
```

### **Week 4+: Phase 3.9 (Helpers - Optional)**
```
As time permits: Groups of 10 similar functions
Focus on most frequently called functions first
```

---

## 💾 DATABASE OF UNLOGGED FUNCTIONS

### **Quick Reference Table**

| Phase | Category | Functions | Files | Effort | Priority |
|-------|----------|-----------|-------|--------|----------|
| 3.5 | Settings | 8 | 1 | 1-1.5h | HIGH |
| 3.6 | Reporting | 11 | 1 | 1.5-2h | MEDIUM |
| 3.7 | File Ops | 2 | 1 | 0.5h | MEDIUM |
| 3.8 | Calculations | 7 | 1 | 1.5-2h | MEDIUM |
| 3.9 | Helpers | 40+ | 2+ | 2-3h | LOW |
| **TOTAL** | - | **78** | **6** | **6-10h** | - |

---

## 🎯 RECOMMENDATION: WHICH TO IMPLEMENT FIRST

### **Quick Wins (Do First)**
1. **Phase 3.5** - Settings functions (8 functions, 1 hour)
   - Easy to implement
   - Improve debugging of config issues
   - Only 1 file to modify

2. **Phase 3.6** - Reporting functions (11 functions, 1.5 hours)
   - Straight-forward data loading logs
   - Nice-to-have transparency
   - Good for understanding user behavior

### **Medium Complexity**
3. **Phase 3.7** - File operations (2 functions, 0.5 hours)
   - Important for document tracking
   - Simple error handling

4. **Phase 3.8** - Calculations (7 functions, 1.5 hours)
   - More complex error scenarios
   - Important for price/cascade debugging

### **Optional/Low Priority**
5. **Phase 3.9** - Helpers (40+ functions, 2-3 hours)
   - Very granular logging
   - Only needed for deep debugging
   - Implement as needed for specific debugging issues

---

## 🚀 ESTIMATED TIMELINE

| Phase | Hours | Token Cost | Commits | Status |
|-------|-------|-----------|---------|--------|
| 3.1-3.4 | 2-3 | 8,000-12,000 | 8 | ✅ DONE |
| 3.5 | 1-1.5 | 3,000-4,000 | 2-3 | ⏳ PLANNED |
| 3.6 | 1.5-2 | 4,000-5,000 | 2-3 | ⏳ PLANNED |
| 3.7 | 0.5 | 1,000-1,500 | 1 | ⏳ PLANNED |
| 3.8 | 1.5-2 | 4,000-5,000 | 2-3 | ⏳ PLANNED |
| 3.9 | 2-3 | 5,000-7,000 | 5-8 | ⏳ OPTIONAL |
| **TOTAL** | **8-12** | **19,000-34,000** | **12-20** | - |

---

## ✨ BENEFITS OF IMPLEMENTING 3.5-3.8

### **Immediate Benefits**
- Complete visibility into all non-trivial operations
- Better debugging of configuration issues
- Understanding of user data access patterns
- Ability to audit file operations

### **Long-term Benefits**
- Comprehensive audit trail for compliance
- Performance optimization opportunities
- User behavior analytics
- System health monitoring

---

## 📝 IMPLEMENTATION CHECKLIST

### **Before Starting Phase 3.5**
- [ ] Review LOGGING_GUIDE.md for pattern reference
- [ ] Understand 9-parameter logWithEmoji() signature
- [ ] Test one new function logging locally
- [ ] Set up testing checklist

### **For Each Phase**
- [ ] Implement logging for all functions in tier
- [ ] Verify syntax with manual test
- [ ] Test filters in Dashboard
- [ ] Create commit with clear message
- [ ] Update TASKS.md with progress

### **After Each Phase**
- [ ] Run PHASE_4_TESTING_GUIDE.md procedures
- [ ] Verify Dashboard displays logs correctly
- [ ] Check for any performance degradation
- [ ] Document any issues or improvements

---

## 📞 NOTES & CONSIDERATIONS

### **Code Quality**
- Reuse existing patterns from Phase 3.1-3.4
- Copy-paste from similar functions where applicable
- Maintain consistent error handling patterns
- Test on actual Google Sheets before committing

### **Performance**
- Logging adds ~5-10ms per function call
- Monitor LOG_DEBUG sheet size (consider archival at >50,000 rows)
- Test with realistic data volumes
- Monitor Dashboard performance with large datasets

### **Testing Strategy**
- Use PHASE_4_TESTING_GUIDE.md for all testing
- Test success AND error paths
- Verify Dashboard filters work correctly
- Check CSV export functionality

### **Future Enhancements**
- Consider log archival strategy (move old logs to separate sheet)
- Implement log retention policies (30/90/365 day cleanup)
- Add performance metrics per function
- Create function call dependency visualization

---

## 🎓 LEARNING RESOURCES

**For Implementation Reference**:
1. Review commits from Phase 3.1-3.4:
   - f082943 (SK Module pattern)
   - 7719341 (Export pattern)
   - 56d785f (Agent pattern with errors)
   - 20785cb (Certification wrapper pattern)

2. Study these functions for patterns:
   - exportPromotions() - Simple wrapper
   - menuCheckService() - External API calls
   - structureDocuments_353pp() - Try-catch with both paths
   - syncMultipleRows() - Complex with multiple logs

**Documentation**:
- LOGGING_GUIDE.md - Complete specification
- PHASE_4_TESTING_GUIDE.md - Testing procedures
- TASKS.md - Progress tracking

---

**Document Version**: 1.0
**Last Updated**: 14.01.2026 16:40 MSK
**Status**: ✅ Ready for Implementation Planning
