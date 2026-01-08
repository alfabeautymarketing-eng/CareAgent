# Route Module Template

Use this template when creating new route modules from endpoints.py.

## Structure

```python
"""
Module docstring describing what endpoints this module contains.

Endpoints:
- POST /path1 - description
- POST /path2 - description

Dependencies:
- Service: sync_service
- Models: SyncRowRequest, SyncResponse
"""

from fastapi import APIRouter, HTTPException
from typing import Optional, List, Dict, Any

from src.utils.logger import logger
from src.api.models.sync import SyncRowRequest, SyncResponse
from src.services.sync import SyncService

# Create router for this module
router = APIRouter()

# Inject dependencies (from main endpoints.py or service factory)
# Example: sync_service: SyncService = Depends(get_sync_service)

# ============== ENDPOINTS ==============

@router.post("/path1", response_model=SyncResponse)
async def endpoint_name(request: SyncRowRequest):
    """
    Endpoint description.

    Args:
        request: Request body with parameters

    Returns:
        SyncResponse with status and details
    """
    logger.info("endpoint_name_started", param=request.param)

    try:
        # Implementation
        result = {}
        return SyncResponse(**result)

    except Exception as e:
        logger.error("endpoint_name_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/path2", response_model=SyncResponse)
async def another_endpoint(request: SyncRowRequest):
    """Another endpoint description."""
    # ...
    pass
```

## Key Points

1. **Docstring:** Include module-level docstring with:
   - What the module does
   - List of endpoints
   - Dependencies (services, models)

2. **Router Creation:**
   ```python
   router = APIRouter()
   ```
   Don't pass prefix here - prefix is specified when including router in main router

3. **Imports:**
   - Import from src.utils, src.api.models, src.services
   - Keep imports minimal and specific

4. **Logging:**
   - Log operation start with relevant parameters
   - Log errors with details
   - Use consistent log keys

5. **Error Handling:**
   - Catch exceptions and convert to HTTPException
   - Include meaningful error details
   - Log before raising

6. **Request/Response Models:**
   - Import from src.api.models.{domain}
   - Use type hints for clarity
   - Return model instances (not dicts)

## Example: sync.py

```python
"""
Synchronization endpoints.

Endpoints:
- POST /sync/row - Sync single row
- POST /sync/range - Sync range of rows
- POST /sync/full - Sync entire sheet

Dependencies:
- SyncService
- LoggingService
- SyncLogService
"""

from fastapi import APIRouter, HTTPException
from typing import Optional, List, Dict, Any

from src.utils.logger import logger
from src.api.models.sync import (
    SyncRowRequest, SyncRangeRequest, SyncResponse
)
from src.services.sync import SyncService

router = APIRouter()

@router.post("/row", response_model=SyncResponse)
async def sync_row(request: SyncRowRequest):
    """
    Sync a single row.

    Matches article with base products and applies sync rules.
    """
    logger.info(
        "sync_row_started",
        spreadsheet_id=request.spreadsheet_id,
        article=request.article
    )

    try:
        # Sync logic here
        result = {"status": "success", "rows_synced": 1}
        return SyncResponse(**result)

    except Exception as e:
        logger.error("sync_row_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))
```

## Registering in router.py

```python
from .sync import router as sync_router

api_router.include_router(
    sync_router,
    prefix="/sync",
    tags=["sync"]
)
```

## Testing

Create corresponding test file:
```
tests/test_routes_sync.py
```

```python
import pytest
from fastapi.testclient import TestClient
from main import app

client = TestClient(app)

def test_sync_row():
    response = client.post("/api/v1/sync/row", json={...})
    assert response.status_code == 200
    assert response.json()["status"] == "success"
```

## Checklist

- [ ] Create module docstring with endpoint list
- [ ] Import necessary dependencies
- [ ] Create APIRouter instance
- [ ] Implement all endpoints from endpoints.py
- [ ] Add logging to start/end/error
- [ ] Use response models for type safety
- [ ] Add docstrings to each endpoint
- [ ] Handle exceptions properly
- [ ] Test with sample requests
- [ ] Update src/api/routes/__init__.py to export router
- [ ] Register in src/api/router.py
- [ ] Run full test suite to verify

## Common Issues

**Issue:** Import CircularDependency
**Solution:** Use lazy imports or move shared code to models/constants

**Issue:** Missing Service Instances
**Solution:** Service instances should be injected from main router.py or use dependency injection

**Issue:** Endpoint Path Mismatch
**Solution:** Prefix is set at include_router() level, don't include in @router.post()

**Issue:** Model Not Found
**Solution:** Create in src/api/models/{domain}.py and import at top of route module

---

For more information, see ENDPOINTS_REFACTORING_PLAN.md
