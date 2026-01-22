/**
 * gas/00_DevTools.gs — Инструменты разработчика
 * ---------------------------------------------------------------------------------------
 * ОПИСАНИЕ:
 *   Функции для управления настройками сервера и отладки прямо из Google Таблицы.
 *   Позволяет быстро переключаться между локальным туннелем (ngrok) и VPS.
 * =======================================================================================
 */

/**
 * Устанавливает кастомный URL сервера (например, ngrok туннель).
 */
function serverSetLocalTunnel() {
  const ui = SpreadsheetApp.getUi();
  const scriptProps = PropertiesService.getScriptProperties();
  const currentUrl = scriptProps.getProperty('SERVER_URL') || 'http://46.226.167.153:8000';

  const result = ui.prompt(
    '🔗 Установка URL туннеля',
    'Введите URL вашего ngrok туннеля (например, https://xxxx.ngrok-free.app):\n\nТекущий URL: ' + currentUrl,
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    let newUrl = result.getResponseText().trim();
    
    // Базовая валидация
    if (!newUrl.startsWith('http')) {
      ui.alert('❌ Ошибка: URL должен начинаться с http:// или https://');
      return;
    }

    // Убираем слеш в конце если есть
    if (newUrl.endsWith('/')) {
      newUrl = newUrl.slice(0, -1);
    }

    // Сохраняем
    scriptProps.setProperty('SERVER_URL', newUrl);
    
    // Очищаем кэш меню чтобы обновить ссылки
    if (typeof refreshMenu === 'function') {
      refreshMenu();
    }

    ui.alert('✅ Готово!', 'SERVER_URL изменен на: ' + newUrl + '\n\nВсе запросы теперь идут на ваш локальный компьютер через туннель.', ui.ButtonSet.OK);
  }
}

/**
 * Сбрасывает URL сервера на стандартный продакшн (VPS).
 */
function serverResetToProduction() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '🚀 Сброс на Production',
    'Вы уверены, что хотите вернуть SERVER_URL на стандартный VPS (46.226.167.153)?',
    ui.ButtonSet.YES_NO
  );

  if (confirm == ui.Button.YES) {
    if (typeof initProductionServerUrl === 'function') {
      initProductionServerUrl();
      
      // Обновляем меню
      if (typeof refreshMenu === 'function') {
        refreshMenu();
      }
      
      ui.alert('✅ Готово!', 'SERVER_URL возвращен на продакшн VPS.', ui.ButtonSet.OK);
    } else {
      ui.alert('❌ Ошибка: Функция initProductionServerUrl не найдена.');
    }
  }
}

/**
 * Показывает текущий рабочий SERVER_URL.
 */
function serverShowCurrentUrl() {
  const ui = SpreadsheetApp.getUi();
  const url = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'Не установлен (используется fallback)';
  
  ui.alert('ℹ️ Текущие настройки сервера', 'SERVER_URL: ' + url, ui.ButtonSet.OK);
}

/**
 * Показывает Spreadsheet ID (полезно для отладки конфигов).
 */
function debugShowSpreadsheetId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const id = ss.getId();
  const name = ss.getName();
  
  const htmlOutput = HtmlService
    .createHtmlOutput('<p>Название: <b>' + name + '</b></p><p>ID: <code style="background: #eee; padding: 2px 5px;">' + id + '</code></p><p style="color: #666; font-size: 12px;">Используйте этот ID в файлах конфигурации проекта (mt.yaml, sk.yaml, ss.yaml).</p>')
    .setWidth(450)
    .setHeight(150);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🆔 Данные таблицы');
}
