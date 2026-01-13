/**
 * 05LogDashboard.gs
 * Task 1.4-1.5: Логирование - Дашборд для просмотра логов
 *
 * Содержит серверные функции для работы с дашбордом логов:
 * - Загрузка логов из LOG_DEBUG и LOG
 * - Фильтрация, поиск, пагинация
 * - Экспорт в CSV
 * - Расчет метрик
 *
 * Версия: 1.0
 * Статус: Ready for Task 1.4
 */

(function() {
  // Экспортируем в глобальный контекст
  const Lib = global.Lib = global.Lib || {};

  /**
   * Загружает логи из LOG_DEBUG и LOG листов
   *
   * @param {object} options - Опции загрузки
   *   - maxRows: максимальное кол-во строк (по умолчанию 500)
   *   - hoursBack: сколько часов назад загружать (по умолчанию 24)
   * @returns {array} Массив объектов логов
   */
  Lib.loadLogs = function(options) {
    options = options || {};
    const maxRows = options.maxRows || 500;
    const hoursBack = options.hoursBack || 24;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logs = [];

    // Определяем временной порог
    const now = new Date();
    const cutoffTime = new Date(now.getTime() - hoursBack * 60 * 60 * 1000);

    try {
      // Загружаем из LOG_DEBUG
      const debugSheet = ss.getSheetByName(Lib.CONFIG.SHEETS.LOG_DEBUG);
      if (debugSheet) {
        const data = debugSheet.getDataRange().getValues();

        // Пропускаем заголовок (строка 1)
        for (let i = 1; i < data.length && logs.length < maxRows; i++) {
          const row = data[i];
          if (!row[0]) continue; // Пропускаем пустые строки

          try {
            const timestamp = new Date(row[0]);
            if (timestamp < cutoffTime) continue;

            logs.push({
              time: row[0],
              category: row[1] || 'General',
              status: row[2] || 'PROGRESS',
              message: row[3] || '',
              function: row[4] || '',
              details: row[5] || '',
              variables: row[6] || '',
              result: row[7] || '',
              source: 'LOG_DEBUG'
            });
          } catch(e) {
            console.warn('Error parsing log row:', e);
          }
        }
      }

      // Загружаем из LOG (синхронизация)
      const logSheet = ss.getSheetByName(Lib.CONFIG.SHEETS.LOG);
      if (logSheet && logs.length < maxRows) {
        const data = logSheet.getDataRange().getValues();

        for (let i = 1; i < data.length && logs.length < maxRows; i++) {
          const row = data[i];
          if (!row[0]) continue;

          try {
            const timestamp = new Date(row[0]);
            if (timestamp < cutoffTime) continue;

            logs.push({
              time: row[0],
              category: 'Sync',
              status: row[2] || 'SUCCESS',
              message: row[3] || '',
              function: row[4] || 'syncSelectedRow',
              details: row[1] || '',
              variables: '',
              result: '',
              source: 'LOG'
            });
          } catch(e) {
            console.warn('Error parsing sync log row:', e);
          }
        }
      }

      // Сортируем по времени (новые сверху)
      logs.sort((a, b) => new Date(b.time) - new Date(a.time));

      return logs.slice(0, maxRows);
    } catch(e) {
      console.error('Error loading logs:', e);
      return [];
    }
  };

  /**
   * Фильтрует логи по заданным критериям
   *
   * @param {array} logs - Массив логов
   * @param {object} filters - Объект фильтров
   *   - period: 'today' | '24h' | '7d' | {start: Date, end: Date}
   *   - categories: array of strings
   *   - statuses: array of strings
   *   - levels: array of strings (DEBUG, INFO, WARN, ERROR)
   *   - search: string для поиска
   * @returns {array} Отфильтрованные логи
   */
  Lib.filterLogs = function(logs, filters) {
    filters = filters || {};
    let filtered = logs;

    // Фильтр по периоду
    if (filters.period) {
      const now = new Date();
      let startDate;

      if (typeof filters.period === 'string') {
        if (filters.period === 'today') {
          startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        } else if (filters.period === '24h') {
          startDate = new Date(now.getTime() - 24 * 60 * 60 * 1000);
        } else if (filters.period === '7d') {
          startDate = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
        }
      } else if (filters.period.start && filters.period.end) {
        startDate = filters.period.start;
        const endDate = filters.period.end;
        filtered = filtered.filter(log => {
          const logTime = new Date(log.time);
          return logTime >= startDate && logTime <= endDate;
        });
      }

      if (startDate) {
        filtered = filtered.filter(log => new Date(log.time) >= startDate);
      }
    }

    // Фильтр по категориям
    if (filters.categories && filters.categories.length > 0) {
      filtered = filtered.filter(log => filters.categories.includes(log.category));
    }

    // Фильтр по статусам
    if (filters.statuses && filters.statuses.length > 0) {
      filtered = filtered.filter(log => filters.statuses.includes(log.status));
    }

    // Фильтр по уровням
    if (filters.levels && filters.levels.length > 0) {
      filtered = filtered.filter(log => {
        const logLevel = _getLogLevel(log.message);
        return filters.levels.includes(logLevel);
      });
    }

    // Текстовый поиск
    if (filters.search && filters.search.trim()) {
      const searchLower = filters.search.toLowerCase();
      filtered = filtered.filter(log => {
        return (log.message || '').toLowerCase().includes(searchLower) ||
               (log.function || '').toLowerCase().includes(searchLower) ||
               (log.details || '').toLowerCase().includes(searchLower);
      });
    }

    return filtered;
  };

  /**
   * Вычисляет метрики для логов
   *
   * @param {array} logs - Массив логов
   * @returns {object} Объект с метриками
   */
  Lib.calculateMetrics = function(logs) {
    const metrics = {
      total: logs.length,
      success: 0,
      warning: 0,
      error: 0,
      start: 0,
      progress: 0
    };

    logs.forEach(log => {
      if (log.status === 'SUCCESS') metrics.success++;
      else if (log.status === 'WARNING') metrics.warning++;
      else if (log.status === 'ERROR') metrics.error++;
      else if (log.status === 'START') metrics.start++;
      else if (log.status === 'PROGRESS') metrics.progress++;
    });

    return metrics;
  };

  /**
   * Экспортирует логи в CSV формат
   *
   * @param {array} logs - Массив логов
   * @returns {string} CSV строка
   */
  Lib.exportLogsToCSV = function(logs) {
    if (!logs || logs.length === 0) {
      return 'No logs to export';
    }

    const headers = ['Время', 'Категория', 'Статус', 'Сообщение', 'Функция', 'Детали', 'Параметры', 'Результаты'];
    const rows = [headers.join(',')];

    logs.forEach(log => {
      const row = [
        _escapeCSV(log.time),
        _escapeCSV(log.category),
        _escapeCSV(log.status),
        _escapeCSV(log.message),
        _escapeCSV(log.function),
        _escapeCSV(log.details),
        _escapeCSV(log.variables),
        _escapeCSV(log.result)
      ];
      rows.push(row.join(','));
    });

    return rows.join('\n');
  };

  /**
   * Пагинирует логи
   *
   * @param {array} logs - Массив логов
   * @param {number} pageSize - Размер страницы (по умолчанию 50)
   * @param {number} pageNumber - Номер страницы (начиная с 1)
   * @returns {object} {items: array, totalPages: number, currentPage: number, totalItems: number}
   */
  Lib.paginateLogs = function(logs, pageSize, pageNumber) {
    pageSize = pageSize || 50;
    pageNumber = Math.max(1, pageNumber || 1);

    const totalItems = logs.length;
    const totalPages = Math.ceil(totalItems / pageSize);
    const startIndex = (pageNumber - 1) * pageSize;
    const endIndex = Math.min(startIndex + pageSize, totalItems);

    return {
      items: logs.slice(startIndex, endIndex),
      totalPages: totalPages,
      currentPage: pageNumber,
      totalItems: totalItems
    };
  };

  /**
   * Получает все категории с эмодзи
   *
   * @returns {object} {categoryName: emoji}
   */
  Lib.getCategoriesWithEmojis = function() {
    if (!Lib.__LOG_CATEGORIES__) {
      return {};
    }

    const categories = {};
    Lib.__LOG_CATEGORIES__.forEach((cat, index) => {
      if (Lib._getEmojiForCategory) {
        categories[cat] = Lib._getEmojiForCategory(cat);
      } else {
        categories[cat] = '❓';
      }
    });

    return categories;
  };

  /**
   * Получает все статусы с эмодзи
   *
   * @returns {object} {statusName: emoji}
   */
  Lib.getStatusesWithEmojis = function() {
    if (!Lib.__LOG_STATUSES__) {
      return {};
    }

    const statuses = {};
    Lib.__LOG_STATUSES__.forEach((status) => {
      if (Lib._getEmojiForStatus) {
        statuses[status] = Lib._getEmojiForStatus(status);
      } else {
        statuses[status] = '❓';
      }
    });

    return statuses;
  };

  /**
   * Открывает дашборд логов в модальном окне
   *
   * Task 1.4: Основная функция для открытия дашборда
   */
  Lib.openLogDashboard = function() {
    try {
      const ui = SpreadsheetApp.getUi();

      try {
        const template = HtmlService.createTemplateFromFile('LogDashboard');
        template.spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

        const html = template.evaluate()
          .setWidth(1400)
          .setHeight(900);

        ui.showModalDialog(html, '📊 ДАШБОРД ЛОГОВ');
      } catch(e) {
        if (e.message.includes('LogDashboard')) {
          ui.alert('Ошибка', 'HTML файл LogDashboard.html не найден. Попробуйте позже.', ui.ButtonSet.OK);
        } else {
          throw e;
        }
      }
    } catch(e) {
      console.error('Error opening dashboard:', e);
      SpreadsheetApp.getUi().alert('Ошибка при открытии дашборда: ' + e.message);
    }
  };

  // ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =====

  /**
   * Определяет уровень логирования из сообщения
   */
  function _getLogLevel(message) {
    const msg = (message || '').toUpperCase();
    if (msg.includes('❌') || msg.includes('ERROR')) return 'ERROR';
    if (msg.includes('⚠️') || msg.includes('WARNING')) return 'WARN';
    if (msg.includes('🐛') || msg.includes('DEBUG')) return 'DEBUG';
    return 'INFO';
  }

  /**
   * Экранирует строку для CSV
   */
  function _escapeCSV(value) {
    if (!value) return '""';
    const stringValue = String(value);
    if (stringValue.includes(',') || stringValue.includes('"') || stringValue.includes('\n')) {
      return '"' + stringValue.replace(/"/g, '""') + '"';
    }
    return stringValue;
  }

})();

// ===== ГЛОБАЛЬНЫЕ ФУНКЦИИ-ОБЕРТКИ ДЛЯ google.script.run =====

/**
 * Загружает логи (для HTML dashboard)
 */
function loadLogs(options) {
  return Lib.loadLogs(options);
}

/**
 * Фильтрует логи (для HTML dashboard)
 */
function filterLogs(logs, filters) {
  return Lib.filterLogs(logs, filters);
}

/**
 * Вычисляет метрики (для HTML dashboard)
 */
function calculateMetrics(logs) {
  return Lib.calculateMetrics(logs);
}

/**
 * Экспортирует логи в CSV (для HTML dashboard)
 */
function exportLogsToCSV(logs) {
  return Lib.exportLogsToCSV(logs);
}

/**
 * Получает категории с эмодзи (для HTML dashboard)
 */
function getCategoriesWithEmojis() {
  const categories = Lib.getCategoriesWithEmojis();
  const statuses = Lib.getStatusesWithEmojis();
  return { categories: categories, statuses: statuses };
}

/**
 * Открывает дашборд (для меню)
 */
function openLogDashboard() {
  return Lib.openLogDashboard();
}
