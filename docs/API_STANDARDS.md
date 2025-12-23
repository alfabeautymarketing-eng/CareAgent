# API Standards & Server Integration

## URL Structure
All server endpoints MUST be prefixed with `/api/v1`.

### Incorrect
`POST /sync/batch-event`
`POST /sync/row`

### Correct
`POST /api/v1/sync/batch-event`
`POST /api/v1/sync/row`

## GAS Integration
The Google Apps Script client (`z-server-overrides.js`) has been updated to automatically prepend `/api/v1` if missing in `Lib.callServer`.

However, for clarity and debugging, prefer using full paths or rely on the helper correctly.

## Endpoints
- `/api/v1/sync/event` (Single onEdit)
- `/api/v1/sync/batch-event` (Batch onEdit)
- `/api/v1/sync/add-article`
- `/api/v1/sync/delete-articles`
