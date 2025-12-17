
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
