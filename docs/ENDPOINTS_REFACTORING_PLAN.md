# Endpoints.py Refactoring Plan (Phase 4)

## Executive Summary

**Current State:** 3271-line monolithic `src/api/endpoints.py`
**Target State:** Modular endpoint organization across 20+ specialized route modules
**Benefits:**
- Improved maintainability
- Parallel development capability
- Better testing isolation
- Clearer separation of concerns
- Reduced cognitive load per file

## Refactoring Strategy

### Step 1: Create Directory Structure ✅

```
src/api/
├── routes/                  # Endpoint route modules
│   ├── sync.py
│   ├── rules.py
│   ├── logs.py
│   ├── sync_logs.py
│   ├── metadata.py
│   ├── certification.py
│   ├── price.py
│   ├── formulas.py
│   ├── ai.py
│   ├── cascade.py
│   ├── order.py
│   ├── export.py
│   ├── invoice.py
│   ├── sorting.py
│   ├── documents.py
│   ├── external_docs.py
│   ├── function_logs.py
│   ├── cache.py
│   ├── task_status.py
│   ├── menu.py
│   └── sheets.py
│
├── models/                  # Pydantic request/response models
│   ├── sync.py
│   ├── rules.py
│   ├── logs.py
│   ├── price.py
│   ├── cascade.py
│   ├── order.py
│   ├── export.py
│   ├── certification.py
│   ├── invoice.py
│   ├── formulas.py
│   └── ai.py
│
├── constants/              # Configuration & constants
│   ├── projects.py         # PROJECT_IDS, PROJECT_NAMES
│   ├── menu_config.py      # Menu configuration
│   └── sheet_order.py      # Sheet ordering
│
├── router.py               # Main router registration
└── endpoints.py            # (Deprecated - can remove after migration)
```

### Step 2: Extract Models (No Breaking Changes)

**Action:** Move all Pydantic models from endpoints.py to models/ subdirectory

**Models to Extract:**

**sync.py:**
- `SyncRowRequest`
- `SyncRangeRequest`
- `SyncAddArticleRequest`
- `DeleteArticlesRequest`
- `SyncEventRequest`
- `SyncBatchEventRequest`

**rules.py:**
- `RuleItem`
- `RuleCreateRequest`
- `RuleUpdateRequest`
- `RuleToggleRequest`
- `RulesSaveRequest`

**logs.py:**
- `LogInitRequest`
- `LogArchiveRequest`
- `LogResetRequest`
- `LogRotationRequest`
- `LogEntryRequest`
- `LogArchiveResponse`
- `LogStatusResponse`

**price.py:**
- `PriceProcessRequest`
- `TaskStatusResponse`

**cascade.py:**
- `CascadeProcessRequest`
- `CascadeProcessResponse`

**order.py:**
- `OrderFilterRequest`
- `OrderFilterResponse`

**export.py:**
- `ExportRequest`
- `ExportResponse`

**certification.py:**
- `CertificationNewsRequest`
- `CertificationSpiritsRequest`
- `CertificationProtocolsRequest`
- `CertificationResponse`

**invoice.py:**
- `InvoiceFormatRequest`
- `InvoiceCreateRequest`
- `InvoiceResponse`

**formulas.py:**
- `FormulaPriceDynamicsRequest`
- `FormulaPriceCalcRequest`
- `FormulaAddYearRequest`
- `FormulaResponse`

**ai.py:**
- `AIPdfAnalyzeRequest`
- `AISimpleAnalyzeRequest`
- `AIAnalyzeRequest`
- `AIAnalyzeBatchRequest`
- `AIConfigureRequest`
- `AISettingsResponse`
- `AICategoryResponse`

### Step 3: Extract Constants

**constants/projects.py:**
```python
PROJECT_IDS = { ... }  # 11 mappings
PROJECT_NAMES = { ... }  # 3 project names
```

**constants/menu_config.py:**
```python
MENU_CONFIGS = { ... }  # Per-project menus
PRIMARY_DATA_MENU_ORDER = [ ... ]
PRIMARY_DATA_MENU_ACTIONS = { ... }
BASE_MENU_GROUPS = [ ... ]
ORDER_STAGES_MENU_ORDER = [ ... ]
ORDER_STAGES_MENU_ACTIONS = { ... }
```

**constants/sheet_order.py:**
```python
SHEET_ORDER = [ ... ]  # 19-item ordering
```

### Step 4: Create Route Modules

**Order of Extraction (from independent to dependent):**

1. **Independent Routes (No cross-dependencies)**
   - `routes/external_docs.py` - 3 endpoints
   - `routes/cache.py` - 1 endpoint
   - `routes/task_status.py` - 1 endpoint

2. **Single-Service Routes**
   - `routes/function_logs.py` - 3 endpoints (FunctionLogService)
   - `routes/documents.py` - 2 endpoints (SheetsService)
   - `routes/cascade.py` - 2 endpoints (SheetsService)
   - `routes/order.py` - 2 endpoints (SheetsService)
   - `routes/export.py` - 2 endpoints (SheetsService)
   - `routes/metadata.py` - 6 endpoints (SheetsService, MetaCacheService)
   - `routes/sorting.py` - 2 endpoints (SheetsService, SortingService, LoggingService)

3. **Specialized Service Routes**
   - `routes/certification.py` - 4 endpoints (CertificationService)
   - `routes/invoice.py` - 2 endpoints (InvoiceService)
   - `routes/formulas.py` - 3 endpoints (FormulaService)
   - `routes/ai.py` - 8 endpoints (AIService, GeminiClient)
   - `routes/price.py` - 3 endpoints (PriceProcessor)

4. **Core Complex Routes**
   - `routes/sync_logs.py` - 8 endpoints (SyncLogService)
   - `routes/logs.py` - 11 endpoints (LoggingService, SheetsService)
   - `routes/rules.py` - 8 endpoints (SyncService)
   - `routes/sync.py` - 7 endpoints (SyncService, LoggingService, SyncLogService)

5. **Configuration Routes**
   - `routes/menu.py` - 2 endpoints (Configuration)
   - `routes/sheets.py` - 1 endpoint (SheetsService)

### Step 5: Create Router

**src/api/router.py:**
```python
from fastapi import APIRouter

from .routes import (
    sync, rules, logs, sync_logs, metadata,
    certification, price, formulas, ai, cascade,
    order, export, invoice, sorting, documents,
    external_docs, function_logs, cache, task_status,
    menu, sheets
)

# Create main router
api_router = APIRouter()

# Include all sub-routers
api_router.include_router(sync.router, prefix="/sync", tags=["sync"])
api_router.include_router(rules.router, prefix="/rules", tags=["rules"])
api_router.include_router(logs.router, prefix="/logs", tags=["logs"])
# ... etc for all routers

__all__ = ["api_router"]
```

### Step 6: Update Application Setup

**main.py / server startup:**
```python
from src.api.router import api_router

app.include_router(api_router, prefix="/api/v1")
```

## Implementation Order

1. ✅ Create directory structure
2. Extract models to `src/api/models/` (batch operation)
3. Extract constants to `src/api/constants/` (batch operation)
4. Create individual route modules (parallel possible)
5. Create `src/api/router.py` (integration)
6. Update main application setup
7. Run tests to verify functionality
8. Remove/archive old `endpoints.py`

## Expected Outcomes

### Code Metrics
| Metric | Before | After |
|--------|--------|-------|
| endpoints.py size | 3271 lines | ~0 (deprecated) |
| Largest route module | 3271 | ~300 |
| Number of modules | 1 | 20+ |
| Model consolidation | Scattered | Organized by domain |

### Quality Improvements
- ✅ Reduced cognitive load (300 vs 3271 lines per file)
- ✅ Easier to locate specific endpoint
- ✅ Independent testing per route module
- ✅ Parallel development on different modules
- ✅ Clear separation of concerns
- ✅ Better IDE navigation and autocompletion

### No Breaking Changes
- Endpoint paths unchanged
- Request/response contracts unchanged
- API behavior identical
- Pure internal reorganization

## Risk Mitigation

### Risks
1. Breaking changes during migration
   - **Mitigation:** Create models/ and constants/ first (non-breaking), then routes

2. Import circular dependencies
   - **Mitigation:** Use lazy imports where needed, maintain clear dependency graph

3. Missing endpoints during migration
   - **Mitigation:** Keep old endpoints.py during transition, verify all endpoints copied

4. Service initialization issues
   - **Mitigation:** Centralize service injection in router.py

### Testing Strategy
1. Unit tests for individual route modules
2. Integration tests for endpoint groups
3. Full system tests verifying all endpoints
4. API contract tests to ensure no breaking changes

## Timeline

**Estimate:** 3-4 hours for complete refactoring

- Model extraction: 30 minutes
- Constants extraction: 20 minutes
- Route module creation: 90 minutes (parallel possible)
- Router setup: 20 minutes
- Testing & verification: 30 minutes
- Cleanup & documentation: 20 minutes

## Next Steps

1. Execute Step 2: Extract models
2. Execute Step 3: Extract constants
3. Execute Steps 4-5: Create route modules and router
4. Verify functionality
5. Clean up old code
6. Commit changes

---

## Duplicate Endpoints Identified

### Issue 1: /sync-logs/{spreadsheet_id}
- **Line 564:** Async version with detailed filters
- **Line 3118:** Path-based variant for UI
- **Action:** Consolidate into single endpoint with optional query params

### Issue 2: /logs/archive
- **Line 807:** Renamed endpoint (OLD)
- **Line 2057:** Current endpoint (NEW)
- **Action:** Keep current (line 2057), remove old

### Issue 3: Async vs Non-async
- **Sync-logs:** 4 async (564, 616, 667, 717) + 4 non-async (3040, 3083, 3118, 3152)
- **Logs:** 8 async (2057+) + 3 non-async (3040+)
- **Function-logs:** 1 async + 2 non-async
- **Action:** Consolidate to single endpoint with async pattern

## Recommendations

1. **Consolidate duplicates** before refactoring
2. **Standardize async pattern** across all endpoints
3. **Move dynamic imports to module level** for better performance
4. **Consider adding OpenAPI tags** for better documentation
5. **Implement proper error handling middleware** to reduce duplication

---

**Status:** Ready for implementation ✅
**Priority:** High (code maintainability)
**Impact:** Zero (internal reorganization only)
