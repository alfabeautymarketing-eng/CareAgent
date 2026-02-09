# 📑 Индекс описаний кнопок AgentCare v2.0

**Систематическая документация всех кнопок в жизненном цикле товара**

Последняя обновление: 2026-02-03
Статус: Актуализация под структуру v2.0

---

## 📌 Как использовать этот индекс

1. Каждая кнопка соответствует функции на сервере.
2. Статус "⏳" означает, что детальная документация по кнопке еще не создана.
3. Статус "✅" означает, что документация готова (в данный момент идет процесс обновления).

---

## 🧾 ГРУППА 1: ЗАКАЗ (Order)

| Кнопка | Функция сервера | Статус | Документ |
|--------|-----------------|--------|----------|
| 📥 Обработка | `serverProcessPrimaryData` | ✅ | [Обзор](BUTTON_01_OBRABOTKA_OVERVIEW.md) / [MT](BUTTON_01_OBRABOTKA_MT.md) / [SK](BUTTON_01_OBRABOTKA_SK.md) / [SS](BUTTON_01_OBRABOTKA_SS.md) |
| 📊 Загрузить остатки | `serverLoadStockData` | ✅ | [BUTTON_02_LOAD_STOCK_DATA.md](BUTTON_02_LOAD_STOCK_DATA.md) |

---

## 🔄 ГРУППА 2: СОРТИРОВКА (Sort)

| Кнопка | Функция сервера | Статус | Документ |
|--------|-----------------|--------|----------|
| Сортировать по производителю | `sortByManufacturer` | ⏳ | - |
| Сортировать по прайсу | `sortByPrice` | ⏳ | - |
| 1. Все данные | `serverShowAllOrderData` | ⏳ | - |
| 2. Заказ | `serverShowOrderStage` | ⏳ | - |
| 3. Акции | `serverShowPromotionsStage` | ⏳ | - |
| 4. Набор | `serverShowSetStage` | ⏳ | - |
| 5. Прайс | `serverShowPriceStage` | ⏳ | - |

---

## 🏷️ ГРУППА 3: ПРАЙС-ЛИСТ (Price List)

| Кнопка | Функция сервера | Статус | Документ |
|--------|-----------------|--------|----------|
| 💰 Анализировать цены | `menuAnalyzePrices` | ⏳ | - |
| ✅ Сформировать свежий прайс | `serverGenerateFreshPriceList` | ⏳ | - |
| 📅 New год для динамика | `serverAddNewYearColumns` | ⏳ | - |

---

## 🛒 ГРУППА 4: ЗАКАЗ & ДОКУМЕНТАЦИЯ (Order & Docs)

| Кнопка | Функция сервера | Статус | Документ |
|--------|-----------------|--------|----------|
| 📋 Подготовить заказ | `serverPrepareOrder` | ⏳ | - |
| 📦 Подготовить документы для таможни | `serverPrepareCustomsDocuments` | ⏳ | - |
| ✏️ Заполнить разрешительную документацию | `serverFillPermitDocumentation` | ⏳ | - |
| 🔬 Сертификация | `serverCertificationWorkbench` | ⏳ | - |
| 📄 Лист новинки | `serverCreateNewsSheet` | ⏳ | - |
| 📋 Протоколы (353пп) | `serverGenerateProtocols353pp` | ⏳ | - |
| 📋 ДС Макеты (353пп) | `serverGenerateDsLayouts` | ⏳ | - |
| 📦 Собрать документы для заявки | `structureDocuments_353pp` | ⏳ | - |
| 🥃 Спирты - Посчитать | `serverCalculateSpiritNumbers` | ⏳ | - |
| 🥃 Спирты - Создать Макеты | `generateSpiritProtocols` | ⏳ | - |
| 🔄 Спирты - Пересчитать каскады | `serverRecalculateCascades` | ⏳ | - |

---

## 📦 ГРУППА 5: ЭКСПОРТ & ВЫГРУЗКА (Export)

| Кнопка | Функция сервера | Статус | Документ |
|--------|-----------------|--------|----------|
| 📤 Выгрузить Акции | `serverExportPromotions` | ⏳ | - |
| 📤 Выгрузить Наборы | `serverExportSets` | ⏳ | - |
| 📄 Форматировать лист 'Ордер' | `serverFormatOrderSheet` | ⏳ | - |
| 📄 Создать лист 'Для инвойса' | `serverCreateFullInvoice` | ⏳ | - |
| 📄 Собрать документы для инвойса | `collectAndCopyDocuments` | ⏳ | - |

---

## 🤖 ГРУППА 6: AI АГЕНТ (AI Agent)

| Кнопка | Функция сервера | Статус | Документ |
|--------|-----------------|--------|----------|
| 🔑 Ввести API ключ Gemini | `setupGeminiComplete` | ⏳ | - |
| 📋 Показать настройки AI | `showGeminiSettings` | ⏳ | - |

---

## ⚙️ ГРУППА 7: НАСТРОЙКА (Settings)

### 🔄 Подменю: Синхронизация (Synchronization)
*Управление обменом данными между листами и сервером.*

| Кнопка | Функция сервера | Статус | Документ |
|--------|-----------------|--------|----------|
| 🔄 Синхронизировать всё | `runFullSync` | ✅ | [BUTTON_SYNC_FULL.md](BUTTON_SYNC_FULL.md) |
| 🔄 Синхронизировать строку | `syncSelectedRow` | ✅ | [BUTTON_SYNC_ROW.md](BUTTON_SYNC_ROW.md) |
| ➕ Добавить артикул | `addArticleManually` | ✅ | [BUTTON_ADD_ARTICLE.md](BUTTON_ADD_ARTICLE.md) |
| ❌ Удалить строку с синхронизацией | `deleteSelectedRowsWithSync` | ✅ | [BUTTON_DELETE_ROW_SYNC.md](BUTTON_DELETE_ROW_SYNC.md) |
| 📝 Правила синхронизации | `showSyncRulesManagerDialog` | ✅ | [BUTTON_SYNC_RULES.md](BUTTON_SYNC_RULES.md) |

### 🛠️ Подменю: Настройки сервера (Server Settings)
*Администрирование сервера, триггеры и диагностика.*

| Кнопка | Функция сервера | Статус | Документ |
|--------|-----------------|--------|----------|
| 🔧 Обновить триггеры | `setupTriggers` | ✅ | [BUTTON_SETUP_TRIGGERS.md](BUTTON_SETUP_TRIGGERS.md) |
| 📋 Журнал синхронизации | `showSyncLogDialog` | ✅ | [BUTTON_SYNC_LOG.md](BUTTON_SYNC_LOG.md) |
| 🔍 Проверить статус сервера | `showAllServicesStatus_proxy` | ✅ | [BUTTON_SERVER_STATUS.md](BUTTON_SERVER_STATUS.md) |
| 🔗 Установить URL (ngrok) | `serverSetLocalTunnel` | ⏳ | - |
| 🚀 Reset to Production | `serverResetToProduction` | ⏳ | - |
| 🔄 Обновить меню | `refreshMenu` | ⏳ | - |

### 🟢 Подменю: Ecosystem
*Общие системные инструменты и интеграция с AI.*

| Кнопка | Функция сервера | Статус | Документ |
|--------|-----------------|--------|----------|
| 🏠 Главная страница | `openServerMainPage` | ⏳ | - |
| 📝 Правила (UI) | `openServerRulesPage` | ⏳ | - |
| 📜 Журнал (UI) | `openLogDashboard_proxy` | ✅ | [BUTTON_LOG_DASHBOARD.md](BUTTON_LOG_DASHBOARD.md) |
| 📚 API Docs (Swagger) | `openServerDocsPage` | ⏳ | - |
| 📑 Упорядочить листы | `reorderSheets` | ⏳ | - |
| 🔄 Обновить данные | `callServerLoadFunctions` | ⏳ | - |

### 🗄️ Подменю: Supabase
*Интеграция с базой данных Supabase.*

| Кнопка | Функция сервера | Статус | Документ |
|--------|-----------------|--------|----------|
| 🔗 Открыть Console | `openSupabaseConsole` | ⏳ | - |
| 📊 Просмотр данных | `showSupabaseTablesView` | ⏳ | - |
| ⚙️ Настройки подключения | `configureSupabaseConnection` | ⏳ | - |
| 🔐 Управление правами | `manageSupabasePermissions` | ⏳ | - |
| 📥 Импорт данных | `importSupabaseData` | ⏳ | - |
| 📤 Экспорт данных | `exportSupabaseData` | ⏳ | - |

---

## 📊 Статистика документации

| Группа | Всего кнопок | Документировано |
|--------|-------------|-----------------|
| 🧾 Заказ | 2 | 2 |
| 🔄 Сортировка | 7 | 0 |
| 🏷️ Прайс-лист | 3 | 0 |
| 🛒 Заказ & Документация | 11 | 0 |
| 📦 Экспорт & Выгрузка | 5 | 0 |
| 🤖 AI Агент | 2 | 0 |
| ⚙️ Настройка | 9 | 9 |
| 🛠️ Инструменты разработчика | 4 | 0 |
| **ИТОГО** | **43** | **11** |
