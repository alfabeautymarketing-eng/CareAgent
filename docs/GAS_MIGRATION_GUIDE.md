# GAS to Server Migration Guide

## Overview

This guide explains how to migrate heavy processing functions from Google Apps Script (GAS) to server API calls, improving performance and maintainability.

## Phase 3: GAS → Server Migration

### Files Involved

1. **ServerApi Module** (`gas/99ServerApi.js`)
   - New utility module for calling server endpoints
   - Provides functions like `ServerApi.processPriceList()`, `ServerApi.sortOrderData()`, etc.
   - Must be included in the GAS project

2. **Price Processing Files** (to be updated)
   - `gas/02Обработка бл.зак Mt.js` - MT project
   - `gas/02Обработка бл.зак Sk.js` - SK project
   - `gas/02Обработка бл.зак Ss.js` - SS project

### Migration Strategy

#### Step 1: Add ServerApi Module to GAS Project

1. Copy file `99ServerApi.js` to your Google Apps Script project
2. Ensure `Lib` object is available (for logging functions)
3. Configure server URL:
   ```javascript
   ServerApi.setServerUrl('https://your-agentcare-server.com/api/v1');
   ```

#### Step 2: Replace Existing Functions

For each price processing file, replace:

**BEFORE (Local Processing):**
```javascript
Lib.processMtMainPrice = function () {
  var ui = SpreadsheetApp.getUi();
  var config = _getPrimaryDataConfig_();

  // ... 50+ lines of local processing ...
  // - Read from source document
  // - Parse data
  // - Sync IDs with main sheet
  // - Handle group changes
  // - Handle new articles
  // - Apply formulas
  // - Show dialogs
};
```

**AFTER (Server Call):**
```javascript
Lib.processMtMainPrice = function () {
  var ui = SpreadsheetApp.getUi();
  var config = _getPrimaryDataConfig_();

  if (!ServerApi.isServerAvailable()) {
    ui.alert(
      'Ошибка подключения',
      'Сервер недоступен. Пожалуйста, проверьте соединение.',
      ui.ButtonSet.OK
    );
    return;
  }

  try {
    // Get active spreadsheet ID
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheetId = ss.getId();
    var sourceDocId = config.SOURCE.DOC_ID;

    // Call server to process price list
    var result = ServerApi.processPriceList(
      'mt',           // project
      'main',         // mode
      spreadsheetId,  // target spreadsheet
      sourceDocId,    // source with supplier prices
      { showUI: true, dryRun: false }
    );

    // Handle result
    if (result && result.status === 'queued') {
      Lib.logInfo('[MT] Обработка запущена на сервере: ' + result.task_id);
    } else if (result && result.status === 'success') {
      Lib.logInfo('[MT] Обработка завершена успешно');
      ui.alert('✅ Обработка завершена!', 'Данные загружены с сервера', ui.ButtonSet.OK);
    }

  } catch (error) {
    Lib.logError('[MT] Ошибка: ' + error.toString());
    ui.alert('❌ Ошибка при обработке', error.toString(), ui.ButtonSet.OK);
  }
};
```

### Benefits of Migration

| Aspect | Before (Local) | After (Server) |
|--------|----------------|----------------|
| **Execution Time** | 30-60 seconds | 5-10 seconds (GAS returns immediately) |
| **Server Load** | High (all processing in GAS) | Low (distributed processing) |
| **Reliability** | Timeout risk for large files | Reliable with background processing |
| **Maintenance** | Update GAS code | Update server code once, all GAS clients benefit |
| **Scalability** | Limited by GAS limits | Scalable server infrastructure |
| **User Experience** | Long UI freeze | Quick response, async processing |

### Implementation Details

#### ServerApi.processPriceList(project, mode, spreadsheetId, sourceDocId, options)

**Parameters:**
- `project` (string): 'mt', 'sk', or 'ss'
- `mode` (string): 'main', 'tester', 'samples', or 'auto'
- `spreadsheetId` (string): Target spreadsheet ID
- `sourceDocId` (string): Source document ID
- `options` (object): Optional
  - `dryRun` (boolean): Preview without saving
  - `showUI` (boolean): Show UI dialogs

**Returns:**
```javascript
{
  status: "queued" | "success" | "error",
  message: "Processing started" | "Processing complete" | error details,
  task_id: "price_mt_main_20260108110000",
  processed_rows: 150,
  created_articles: 23,
  group_changes: [...],
  barcode_mismatches: [...]
}
```

**Example Usage:**
```javascript
var result = ServerApi.processPriceList(
  'mt',
  'main',
  SpreadsheetApp.getActiveSpreadsheet().getId(),
  config.SOURCE.DOC_ID,
  { showUI: true }
);

if (result.task_id) {
  Lib.logInfo('Task started: ' + result.task_id);
  // Can check status later with ServerApi.getPriceStatus(result.task_id)
}
```

#### ServerApi.getPriceStatus(taskId)

**Purpose:** Check processing status of a running task

**Parameters:**
- `taskId` (string): Task ID from processPriceList response

**Returns:**
```javascript
{
  status: "running" | "completed" | "failed",
  progress: 45,  // percentage
  message: "Processing...",
  result: { /* final result if completed */ }
}
```

**Example Usage:**
```javascript
function checkPriceStatus() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var taskId = scriptProperties.getProperty('LAST_PRICE_TASK_ID');

  if (taskId) {
    var status = ServerApi.getPriceStatus(taskId);
    var ui = SpreadsheetApp.getUi();
    ui.alert('Статус: ' + status.status + ' (' + status.progress + '%)');
  }
}
```

### Function Mapping

#### Price Processing

| GAS Function | Server Endpoint | Notes |
|-------------|-----------------|-------|
| `Lib.processMtMainPrice()` | `POST /price/process/mt?mode=main` | Process main supplier price list |
| `Lib.processMtTesterPrice()` | `POST /price/process/mt?mode=tester` | Process tester samples |
| `Lib.processMtSamplesPrice()` | `POST /price/process/mt?mode=samples` | Process samples |
| `Lib.processSkPriceSheet()` | `POST /price/process/sk?mode=main` | SK main price list |
| `Lib.processSkPriceProbes()` | `POST /price/process/sk?mode=probes` | SK probes/samples |
| `Lib.processSsPriceSheet()` | `POST /price/process/ss?mode=main` | SS main price list |

#### Sorting/Structuring

| GAS Function | Server Endpoint |
|-------------|-----------------|
| `Lib.sortOrderByManufacturer()` | `POST /sort/structure?sort_type=manufacturer` |
| `Lib.sortOrderByPrice()` | `POST /sort/structure?sort_type=price` |

#### Formula Operations

| GAS Function | Server Endpoint |
|-------------|-----------------|
| `Lib.recalculatePriceDynamicsFormulas()` | `POST /formulas/price-dynamics` |
| `Lib.updatePriceCalculationFormulas()` | `POST /formulas/price-calculation` |

### Handling UI and Dialogs

**UI elements to keep in GAS:**
- ✅ Loading dialogs/toasts
- ✅ Error alerts
- ✅ User confirmations (YES/NO dialogs)
- ✅ Input prompts (volume, category selections)
- ✅ Progress indicators

**Processing logic to move to server:**
- ❌ Data parsing
- ❌ ID matching/generation
- ❌ Sheet manipulation
- ❌ Formula calculations
- ❌ Cross-sheet updates

### Configuration

**Set server URL before using:**
```javascript
// Option 1: In Lib.CONFIG (if available)
Lib.CONFIG = Lib.CONFIG || {};
Lib.CONFIG.SERVER_URL = 'https://your-server.com/api/v1';

// Option 2: Using setServerUrl function
ServerApi.setServerUrl('https://your-server.com/api/v1');

// Option 3: In script properties
PropertiesService.getScriptProperties()
  .setProperty('SERVER_URL', 'https://your-server.com/api/v1');
```

### Error Handling

```javascript
try {
  var result = ServerApi.processPriceList(
    'mt', 'main', spreadsheetId, sourceDocId
  );

  if (result.status === 'error') {
    throw new Error(result.message);
  }

} catch (error) {
  Lib.logError('Price processing failed: ' + error.toString());
  SpreadsheetApp.getUi().alert(
    '❌ Ошибка',
    error.toString(),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
```

### Testing

**1. Check server availability:**
```javascript
if (!ServerApi.isServerAvailable()) {
  throw new Error('Server is not available');
}
```

**2. Get server status:**
```javascript
var status = ServerApi.getServerStatus();
Lib.logInfo('Server version: ' + status.version);
```

**3. Test with dry run:**
```javascript
var result = ServerApi.processPriceList(
  'mt', 'main', spreadsheetId, sourceDocId,
  { dryRun: true, showUI: false }
);
```

## Deployment Checklist

- [ ] Add `99ServerApi.js` to GAS project
- [ ] Configure server URL in GAS
- [ ] Test server connectivity with `ServerApi.isServerAvailable()`
- [ ] Update one price processing function (e.g., `processMtMainPrice`)
- [ ] Test price processing end-to-end
- [ ] Test error handling (server down, invalid input, etc.)
- [ ] Update remaining processing functions (SK, SS, tester, samples)
- [ ] Update sorting/structuring functions
- [ ] Update formula operations
- [ ] Deploy to production
- [ ] Monitor server logs and GAS execution
- [ ] Remove old processing logic once verified working

## Rollback Plan

If server migration causes issues:

1. **Temporary Rollback:** Keep old processing functions as fallback
   ```javascript
   if (ServerApi.isServerAvailable()) {
     // Use server
     return ServerApi.processPriceList(...);
   } else {
     // Fallback to local processing
     return _processPriceListLocal(...);
   }
   ```

2. **Full Rollback:** Replace ServerApi calls with original functions
3. **Hybrid Mode:** Use server for large batches, local for small operations

## Performance Expectations

### Before Migration (Local Processing)
- MT main price: 45-60 seconds
- MT tester price: 30-45 seconds
- Sorting: 20-30 seconds
- Formula updates: 15-20 seconds
- **Total: 2-3 minutes**
- User waits entire time

### After Migration (Server)
- API call + UI response: 1-2 seconds
- Server processing (background): 10-20 seconds
- User can continue working immediately
- Can check status anytime with `ServerApi.getPriceStatus()`

## Next Steps

1. Review this guide
2. Add `ServerApi.js` to GAS project
3. Start with one function migration
4. Test thoroughly
5. Proceed with remaining functions
6. Delete old processing code once verified

---

**Questions or Issues?** Check server logs:
```
logs/server.log
```

**Support:** Contact development team with task ID from ServerApi responses
