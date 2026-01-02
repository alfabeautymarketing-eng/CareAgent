# Правила синхронизации (server-side)

## Ключевая идея
Правила синхронизации хранятся **на сервере** в YAML. Лист «Правила синхро» больше не используется и может быть удалён из таблиц. Источник правды — `config/rules/<spreadsheet_id>.yaml`.

## Формат `config/rules/<spreadsheet_id>.yaml`

Массив правил:
```yaml
- id: 001-Г-С(Примечание/записная книжка)
  mode: unidirectional
  enabled: true
  category: Сертификация
  hashtags: "#sync"
  source_sheet: Главная
  source_header: Примечание/записная книжка
  target_sheet: Сертификация
  target_header: Примечание/записная книжка
  is_external: false
  target_doc_id: ""   # опционально, только если is_external=true
- id: 002-Р<->С(Код ТН ВЭД)
  mode: bidirectional
  enabled: true
  category: Сертификация
  sheet_a: Расчет цены
  header_a: Код ТН ВЭД
  sheet_b: Сертификация
  header_b: Код ТН ВЭД
```

Имена правил генерируются автоматически при сохранении:
- **unidirectional**: `<NNN>-<SRC>-<TGT>(<Header>)`
- **bidirectional**: `<NNN>-<A><-><B>(<HeaderA>)`
где `NNN` — порядковый номер, `SRC`/`TGT` — буквы листов (Г, С, Э, З, Д, Р, П, А и т.д.), `Header` — имя столбца.

Обязательные поля:
- **unidirectional**: `source_sheet`, `source_header`, `target_sheet`, `target_header`.
- **bidirectional**: `sheet_a`, `header_a`, `sheet_b`, `header_b`.
Для внешних правил (`is_external: true`) нужен `target_doc_id` (только для unidirectional). Поле `hashtags` более не используется (при загрузке игнорируется).

## Обновление правил
- Чтобы обновить правила, измените YAML и перезапустите сервис (или подождите 5 минут до истечения кэша/вызовите `force_reload` через API).
- UI для управления правилами: `GET /api/v1/rules-ui` (файл `config/rule_manager.html`).
- Программно: `GET/POST /api/v1/rules/{spreadsheet_id}` и CRUD-эндпоинты в `src/api/endpoints.py`.

## Исключения
- События на листах `Логи` и `Журнал синхро` игнорируются, чтобы не тратить квоту и время на служебные изменения.
