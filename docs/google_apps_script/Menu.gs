/**
 * @OnlyCurrentDoc
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Ecosystem')
      .addItem('Сортировать по производителю 🏭', 'triggerSortManufacturer')
      .addItem('Сортировать по цене 💰', 'triggerSortPrice')
      .addSeparator()
      .addItem('Проверить статус сервера 🟢', 'checkServerStatus')
      .addToUi();
}
