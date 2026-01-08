# GAS Wrapper Examples - Ready to Use

This file contains ready-to-use replacement functions for the price processing GAS files.
Copy these examples into your GAS files to replace the heavy local processing with server calls.

---

## MT Price Processing Examples

### Example 1: Replace `Lib.processMtMainPrice()`

**Location:** `gas/02Обработка бл.зак Mt.js` (around line 39)

**BEFORE:** ~150 lines of local processing
**AFTER:** ~30 lines of wrapper code

```javascript
/**
 * Process MT main supplier price list (Б/З поставщик)
 * Calls server for actual processing, keeps only UI logic in GAS
 */
Lib.processMtMainPrice = function () {
  var ui = SpreadsheetApp.getUi();
  var config = _getPrimaryDataConfig_();
  var menuTitle = _getMenuTitle_(config);

  if (!_isActiveProject_()) {
    ui.alert(
      menuTitle,
      "Эта функция доступна только в проекте MT.",
      ui.ButtonSet.OK
    );
    return;
  }

  // Check server availability
  if (!ServerApi.isServerAvailable()) {
    ui.alert(
      menuTitle,
      "❌ Сервер недоступен. Пожалуйста, проверьте соединение.",
      ui.ButtonSet.OK
    );
    return;
  }

  try {
    Lib.logInfo("[MT] Обработка Б/З поставщик: старт");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheetId = ss.getId();
    var sourceDocId = config.SOURCE.DOC_ID;

    // Call server for price list processing
    var result = ServerApi.processPriceList(
      'mt',           // project
      'main',         // mode
      spreadsheetId,  // target spreadsheet
      sourceDocId,    // source document
      {
        showUI: true,    // Show progress dialog
        dryRun: false    // Actually process
      }
    );

    // Log result
    if (result.status === 'queued') {
      Lib.logInfo(
        "[MT] Обработка запущена на сервере",
        "Task ID: " + result.task_id
      );
      PropertiesService.getScriptProperties()
        .setProperty('LAST_PRICE_TASK_MT_MAIN', result.task_id);
    }

    if (result.status === 'success') {
      Lib.logInfo(
        "[MT] Обработка завершена",
        "Обработано: " + result.processed_rows + " строк, " +
        "создано: " + result.created_articles + " артикулов"
      );
    }

    return result;

  } catch (error) {
    Lib.logError("[MT] Ошибка при обработке: " + error.toString());
    ui.alert(
      menuTitle,
      "❌ Ошибка при обработке:\n" + error.toString(),
      ui.ButtonSet.OK
    );
    throw error;
  }
};
```

### Example 2: Replace `Lib.processMtTesterPrice()`

**Location:** `gas/02Обработка бл.зак Mt.js` (around line 236)

```javascript
/**
 * Process MT tester samples price (Пробники)
 */
Lib.processMtTesterPrice = function () {
  var ui = SpreadsheetApp.getUi();
  var config = _getPrimaryDataConfig_();
  var menuTitle = _getMenuTitle_(config);

  if (!_isActiveProject_()) {
    ui.alert(menuTitle, "Эта функция доступна только в проекте MT.", ui.ButtonSet.OK);
    return;
  }

  if (!ServerApi.isServerAvailable()) {
    ui.alert(menuTitle, "❌ Сервер недоступен.", ui.ButtonSet.OK);
    return;
  }

  try {
    Lib.logInfo("[MT] Обработка пробников: старт");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = ServerApi.processPriceList(
      'mt',
      'tester',
      ss.getId(),
      _getPrimaryDataConfig_().SOURCE.DOC_ID,
      { showUI: true }
    );

    Lib.logInfo("[MT] Обработка пробников: завершена", result.status);
    return result;

  } catch (error) {
    Lib.logError("[MT] Ошибка: " + error.toString());
    ui.alert(menuTitle, "❌ Ошибка: " + error.toString(), ui.ButtonSet.OK);
  }
};
```

### Example 3: Replace `Lib.processMtSamplesPrice()`

**Location:** `gas/02Обработка бл.зак Mt.js` (around line 393)

```javascript
/**
 * Process MT samples (Образцы)
 */
Lib.processMtSamplesPrice = function () {
  var ui = SpreadsheetApp.getUi();
  var config = _getPrimaryDataConfig_();
  var menuTitle = _getMenuTitle_(config);

  if (!_isActiveProject_()) {
    ui.alert(menuTitle, "Эта функция доступна только в проекте MT.", ui.ButtonSet.OK);
    return;
  }

  if (!ServerApi.isServerAvailable()) {
    ui.alert(menuTitle, "❌ Сервер недоступен.", ui.ButtonSet.OK);
    return;
  }

  try {
    Lib.logInfo("[MT] Обработка образцов: старт");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = ServerApi.processPriceList(
      'mt',
      'samples',
      ss.getId(),
      _getPrimaryDataConfig_().SOURCE.DOC_ID,
      { showUI: true }
    );

    Lib.logInfo("[MT] Обработка образцов: завершена", result.status);
    return result;

  } catch (error) {
    Lib.logError("[MT] Ошибка: " + error.toString());
    ui.alert(menuTitle, "❌ Ошибка: " + error.toString(), ui.ButtonSet.OK);
  }
};
```

---

## SK Price Processing Examples

### Example 4: Replace `Lib.processSkPriceSheet()`

**Location:** `gas/02Обработка бл.зак Sk.js`

```javascript
/**
 * Process SK main price list
 * Handles Line + Group hierarchy on server
 */
Lib.processSkPriceSheet = function () {
  var ui = SpreadsheetApp.getUi();
  var config = _getPrimaryDataConfig_();
  var menuTitle = _getMenuTitle_(config);

  if (!_isActiveProject_()) {
    ui.alert(menuTitle, "Эта функция доступна только в проекте SK.", ui.ButtonSet.OK);
    return;
  }

  if (!ServerApi.isServerAvailable()) {
    ui.alert(menuTitle, "❌ Сервер недоступен.", ui.ButtonSet.OK);
    return;
  }

  try {
    Lib.logInfo("[SK] Обработка главного прайс-листа: старт");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = ServerApi.processPriceList(
      'sk',
      'main',
      ss.getId(),
      config.SOURCE.DOC_ID,
      { showUI: true }
    );

    Lib.logInfo(
      "[SK] Обработка завершена",
      "Статус: " + result.status
    );
    return result;

  } catch (error) {
    Lib.logError("[SK] Ошибка: " + error.toString());
    ui.alert(menuTitle, "❌ Ошибка: " + error.toString(), ui.ButtonSet.OK);
  }
};
```

### Example 5: Replace `Lib.processSkPriceProbes()`

**Location:** `gas/02Обработка бл.зак Sk.js`

```javascript
/**
 * Process SK probes/samples
 * Handles special probe structure on server
 */
Lib.processSkPriceProbes = function () {
  var ui = SpreadsheetApp.getUi();
  var config = _getPrimaryDataConfig_();
  var menuTitle = _getMenuTitle_(config);

  if (!_isActiveProject_()) {
    ui.alert(menuTitle, "Эта функция доступна только в проекте SK.", ui.ButtonSet.OK);
    return;
  }

  if (!ServerApi.isServerAvailable()) {
    ui.alert(menuTitle, "❌ Сервер недоступен.", ui.ButtonSet.OK);
    return;
  }

  try {
    Lib.logInfo("[SK] Обработка пробников: старт");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = ServerApi.processPriceList(
      'sk',
      'probes',
      ss.getId(),
      config.SOURCE.DOC_ID,
      { showUI: true }
    );

    Lib.logInfo("[SK] Обработка пробников: завершена");
    return result;

  } catch (error) {
    Lib.logError("[SK] Ошибка: " + error.toString());
    ui.alert(menuTitle, "❌ Ошибка: " + error.toString(), ui.ButtonSet.OK);
  }
};
```

---

## SS Price Processing Examples

### Example 6: Replace `Lib.processSsPriceSheet()`

**Location:** `gas/02Обработка бл.зак Ss.js`

```javascript
/**
 * Process SS main price list
 * Handles professional mode detection on server
 */
Lib.processSsPriceSheet = function () {
  var ui = SpreadsheetApp.getUi();
  var config = _getPrimaryDataConfig_();
  var menuTitle = _getMenuTitle_(config);

  if (!_isActiveProject_()) {
    ui.alert(menuTitle, "Эта функция доступна только в проекте SS.", ui.ButtonSet.OK);
    return;
  }

  if (!ServerApi.isServerAvailable()) {
    ui.alert(menuTitle, "❌ Сервер недоступен.", ui.ButtonSet.OK);
    return;
  }

  try {
    Lib.logInfo("[SS] Обработка главного прайс-листа: старт");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = ServerApi.processPriceList(
      'ss',
      'main',
      ss.getId(),
      config.SOURCE.DOC_ID,
      { showUI: true }
    );

    Lib.logInfo(
      "[SS] Обработка завершена",
      "Статус: " + result.status
    );
    return result;

  } catch (error) {
    Lib.logError("[SS] Ошибка: " + error.toString());
    ui.alert(menuTitle, "❌ Ошибка: " + error.toString(), ui.ButtonSet.OK);
  }
};
```

---

## Sorting/Structuring Examples

### Example 7: Replace `Lib.sortOrderByManufacturer()`

**Location:** `gas/06OrderForm.js` or similar

```javascript
/**
 * Sort orders by manufacturer
 */
Lib.sortOrderByManufacturer = function () {
  var ui = SpreadsheetApp.getUi();

  if (!ServerApi.isServerAvailable()) {
    ui.alert("❌ Ошибка", "Сервер недоступен.", ui.ButtonSet.OK);
    return;
  }

  try {
    Lib.logInfo("Сортировка заказов по производителям: старт");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = ServerApi.sortOrderData(
      ss.getId(),
      'by_manufacturer',
      { showUI: true }
    );

    Lib.logInfo("Сортировка завершена", result.status);
    ui.alert("✅ Готово!", "Заказы отсортированы по производителям", ui.ButtonSet.OK);

  } catch (error) {
    Lib.logError("Ошибка при сортировке: " + error.toString());
    ui.alert("❌ Ошибка", error.toString(), ui.ButtonSet.OK);
  }
};
```

### Example 8: Replace `Lib.sortOrderByPrice()`

**Location:** `gas/06OrderForm.js` or similar

```javascript
/**
 * Sort orders by price
 */
Lib.sortOrderByPrice = function () {
  var ui = SpreadsheetApp.getUi();

  if (!ServerApi.isServerAvailable()) {
    ui.alert("❌ Ошибка", "Сервер недоступен.", ui.ButtonSet.OK);
    return;
  }

  try {
    Lib.logInfo("Сортировка заказов по цене: старт");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = ServerApi.sortOrderData(
      ss.getId(),
      'by_price',
      { showUI: true }
    );

    Lib.logInfo("Сортировка завершена", result.status);
    ui.alert("✅ Готово!", "Заказы отсортированы по цене", ui.ButtonSet.OK);

  } catch (error) {
    Lib.logError("Ошибка при сортировке: " + error.toString());
    ui.alert("❌ Ошибка", error.toString(), ui.ButtonSet.OK);
  }
};
```

---

## Formula Operations Examples

### Example 9: Recalculate Price Dynamics Formulas

**Location:** Price processing files

```javascript
/**
 * Recalculate price dynamics formulas
 */
function recalculatePriceDynamicsFormulas() {
  var ui = SpreadsheetApp.getUi();

  if (!ServerApi.isServerAvailable()) {
    ui.alert("❌ Ошибка", "Сервер недоступен.", ui.ButtonSet.OK);
    return;
  }

  try {
    Lib.logInfo("Пересчёт формул динамики цены: старт");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = ServerApi.recalculatePriceDynamicsFormulas(
      ss.getId(),
      { showUI: true }
    );

    Lib.logInfo("Пересчёт завершен", "Обновлено: " + result.formulas_updated);
    ui.alert("✅ Готово!", "Формулы пересчитаны", ui.ButtonSet.OK);

  } catch (error) {
    Lib.logError("Ошибка при пересчёте: " + error.toString());
    ui.alert("❌ Ошибка", error.toString(), ui.ButtonSet.OK);
  }
}
```

### Example 10: Update Price Calculation Formulas

**Location:** Price processing files

```javascript
/**
 * Update price calculation formulas
 */
function updatePriceCalculationFormulas() {
  var ui = SpreadsheetApp.getUi();

  if (!ServerApi.isServerAvailable()) {
    ui.alert("❌ Ошибка", "Сервер недоступен.", ui.ButtonSet.OK);
    return;
  }

  try {
    Lib.logInfo("Обновление формул расчёта цены: старт");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = ServerApi.updatePriceCalculationFormulas(
      ss.getId(),
      { showUI: true }
    );

    Lib.logInfo("Обновление завершено", "Обновлено: " + result.formulas_updated);
    ui.alert("✅ Готово!", "Формулы обновлены", ui.ButtonSet.OK);

  } catch (error) {
    Lib.logError("Ошибка при обновлении: " + error.toString());
    ui.alert("❌ Ошибка", error.toString(), ui.ButtonSet.OK);
  }
}
```

---

## Common Patterns

### Check Server Before Processing

```javascript
if (!ServerApi.isServerAvailable()) {
  ui.alert("❌ Ошибка", "Сервер недоступен. Пожалуйста, проверьте соединение.", ui.ButtonSet.OK);
  return;
}
```

### Save Task ID for Later Status Check

```javascript
var result = ServerApi.processPriceList(...);
if (result.task_id) {
  PropertiesService.getScriptProperties()
    .setProperty('LAST_TASK_ID_' + project.toUpperCase(), result.task_id);
}
```

### Check Processing Status

```javascript
function checkPriceStatus() {
  var scriptProps = PropertiesService.getScriptProperties();
  var taskId = scriptProps.getProperty('LAST_TASK_ID_MT');

  if (!taskId) {
    SpreadsheetApp.getUi().alert("Нет активных задач");
    return;
  }

  var status = ServerApi.getPriceStatus(taskId);
  SpreadsheetApp.getUi().alert(
    "Статус: " + status.status + "\nПрогресс: " + status.progress + "%"
  );
}
```

### Error Handling

```javascript
try {
  var result = ServerApi.processPriceList(...);
  if (result.status === 'error') {
    throw new Error(result.message);
  }
} catch (error) {
  Lib.logError("Processing failed: " + error.toString());
  ui.alert("❌ Ошибка", error.toString(), ui.ButtonSet.OK);
  return;
}
```

---

## Migration Checklist for Each Function

- [ ] Copy example code
- [ ] Update project code if different ('mt', 'sk', 'ss')
- [ ] Update mode if needed ('main', 'tester', 'samples', 'probes')
- [ ] Add ServerApi.js to GAS project first
- [ ] Test function locally
- [ ] Verify server returns expected result
- [ ] Test error handling (server down, invalid input)
- [ ] Verify UI messages display correctly
- [ ] Commit changes to GAS
- [ ] Deploy via clasp

---

## Questions?

Refer to `GAS_MIGRATION_GUIDE.md` for detailed information about each API function.
