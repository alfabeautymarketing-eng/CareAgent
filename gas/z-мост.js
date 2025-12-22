
/**
 * =======================================================================================
 * Z_ApiStubs.gs — ЗАГЛУШКИ ДЛЯ НЕРАЕЛИЗОВАННОГО API
 * ---------------------------------------------------------------------------------------
 * ОПИСАНИЕ: 
 *   Этот файл содержит временные функции-заглушки для всех пунктов меню,
 *   которые еще не были реализованы. Это предотвращает ошибки "is not a function".
 *   По мере реализации функционала, функции нужно УДАЛЯТЬ из этого файла
 *   и переносить их РЕАЛИЗАЦИЮ в соответствующие логические модули
 *   (D_PriceLogic.gs, E_InvoiceLogic.gs и т.д.).
 * =======================================================================================
 */
var Lib = Lib || {};

(function(Lib, global) {

  /** @private Функция-заглушка для нереализованного функционала. */
  function _showNotImplementedAlert() {
    try {
      SpreadsheetApp.getUi().alert('Функция в разработке', 'Эта функция еще не реализована в новой библиотеке.', SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      console.warn("Функция в разработке (Headless/API)");
    }
  }

  // --- Синхронизация (диалоги и системные) ---

  // --- Поставка ---
  if (!Lib.collectAndCopyDocuments) Lib.collectAndCopyDocuments = _showNotImplementedAlert;

  // --- Сертификация ---
  if (!Lib.createNewsSheetFromCertification) Lib.createNewsSheetFromCertification = _showNotImplementedAlert;
  if (!Lib.generateProtocols_353pp) Lib.generateProtocols_353pp = _showNotImplementedAlert;
  if (!Lib.generateDsLayouts_353pp) Lib.generateDsLayouts_353pp = _showNotImplementedAlert;
  if (!Lib.structureDocuments_353pp) Lib.structureDocuments_353pp = _showNotImplementedAlert;
  if (!Lib.calculateAndAssignSpiritNumbers) Lib.calculateAndAssignSpiritNumbers = _showNotImplementedAlert;
  if (!Lib.generateSpiritProtocols) Lib.generateSpiritProtocols = _showNotImplementedAlert;

})(Lib, this);

// =======================================================================================
// PROXY FUNCTIONS FOR LEGACY BUTTON BINDINGS
// ---------------------------------------------------------------------------------------
// These functions are called by drawing buttons on sheets.
// They call the Python server for high-performance batch sorting.
// =======================================================================================

/**
 * MT: Sort by Manufacturer (proxy for buttons)
 */
function sortMtOrderByManufacturer_proxy() {
  if (typeof callServerStructureSort === 'function') {
    callServerStructureSort('byManufacturer');
  } else {
    SpreadsheetApp.getUi().alert('❌ Функция callServerStructureSort не найдена.');
  }
}

/**
 * MT: Sort by Price (proxy for buttons)
 */
function sortMtOrderByPrice_proxy() {
  if (typeof callServerStructureSort === 'function') {
    callServerStructureSort('byPrice');
  } else {
    SpreadsheetApp.getUi().alert('❌ Функция callServerStructureSort не найдена.');
  }
}

/**
 * SK: Sort by Manufacturer (proxy for buttons)
 */
function sortSkOrderByManufacturer_proxy() {
  if (typeof callServerStructureSort === 'function') {
    callServerStructureSort('byManufacturer');
  } else {
    SpreadsheetApp.getUi().alert('❌ Функция callServerStructureSort не найдена.');
  }
}

/**
 * SK: Sort by Price (proxy for buttons)
 */
function sortSkOrderByPrice_proxy() {
  if (typeof callServerStructureSort === 'function') {
    callServerStructureSort('byPrice');
  } else {
    SpreadsheetApp.getUi().alert('❌ Функция callServerStructureSort не найдена.');
  }
}

/**
 * SS: Sort by Manufacturer (proxy for buttons)
 */
function sortSsOrderByManufacturer_proxy() {
  if (typeof callServerStructureSort === 'function') {
    callServerStructureSort('byManufacturer');
  } else {
    SpreadsheetApp.getUi().alert('❌ Функция callServerStructureSort не найдена.');
  }
}

/**
 * SS: Sort by Price (proxy for buttons)
 */
function sortSsOrderByPrice_proxy() {
  if (typeof callServerStructureSort === 'function') {
    callServerStructureSort('byPrice');
  } else {
    SpreadsheetApp.getUi().alert('❌ Функция callServerStructureSort не найдена.');
  }
}

// =======================================================================================
// TRIGGER SETUP
// =======================================================================================

/**
 * Setup installable triggers for onEdit and onChange.
 * Required for UrlFetchApp to work (simple triggers don't have permissions).
 */
function setupTriggers_proxy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();
  
  // Check existing triggers
  let hasOnEdit = false;
  let hasOnChange = false;
  
  triggers.forEach(function(trigger) {
    const handler = trigger.getHandlerFunction();
    if (handler === 'handleOnEdit') hasOnEdit = true;
    if (handler === 'handleOnChange') hasOnChange = true;
  });
  
  let created = [];
  
  // Create onEdit trigger if not exists
  if (!hasOnEdit) {
    ScriptApp.newTrigger('handleOnEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
    created.push('handleOnEdit');
  }
  
  // Create onChange trigger if not exists
  if (!hasOnChange) {
    ScriptApp.newTrigger('handleOnChange')
      .forSpreadsheet(ss)
      .onChange()
      .create();
    created.push('handleOnChange');
  }
  
  if (created.length > 0) {
    SpreadsheetApp.getUi().alert(
      '✅ Триггеры созданы',
      'Созданы триггеры: ' + created.join(', ') + '\n\nТеперь синхронизация через Python сервер будет работать.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert(
      'ℹ️ Триггеры уже существуют',
      'handleOnEdit и handleOnChange уже настроены.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
