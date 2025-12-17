/**
 * @OnlyCurrentDoc
 */

function createAgentMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('⚙️ Ecosystem v2');
    menu.addItem('Сортировать по производителю 🏭', 'triggerSortManufacturer')
        .addItem('Сортировать по цене 💰', 'triggerSortPrice')
        .addSeparator()
        .addItem('Проверить статус сервера 🟢', 'checkServerStatus')
        .addSeparator()
        .addItem('🔍 Получить ID таблицы', 'debugShowSpreadsheetId') // Added debug tool
        .addToUi();
  } catch (e) {
    Logger.log("Error creating Agent menu: " + e.toString());
    console.error("Error creating Agent menu: " + e.toString());
  }
}

function debugShowSpreadsheetId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const id = ss.getId();
  const ui = SpreadsheetApp.getUi();
  ui.alert("ID этой таблицы:\n" + id);
  console.log("DEBUG: Document ID: " + id);
}
