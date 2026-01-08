"""
API route modules organized by feature/domain.

This package contains specialized router modules that handle specific groups
of endpoints. Each module is responsible for a logical domain (sync, rules, logs, etc.).

**Modules:**
- sync: Synchronization operations (7 endpoints)
- rules: Sync rules management (8 endpoints)
- logs: Log management and archiving (11 endpoints)
- sync_logs: Sync operation logging (8 endpoints)
- metadata: Sheets and metadata discovery (6 endpoints)
- certification: Product certification operations (4 endpoints)
- price: Supplier price list processing (3 endpoints)
- formulas: Formula calculations and updates (3 endpoints)
- ai: AI-powered analysis (8 endpoints)
- cascade: Cascade rule processing (2 endpoints)
- order: Order data operations (2 endpoints)
- export: Data export operations (2 endpoints)
- invoice: Invoice generation (2 endpoints)
- sorting: Data sorting operations (2 endpoints)
- documents: Document structure operations (2 endpoints)
- external_docs: External document registry (3 endpoints)
- function_logs: Function execution logging (3 endpoints)
- cache: Cache management (1 endpoint)
- task_status: Task status tracking (1 endpoint)
- menu: Menu configuration (2 endpoints)
- sheets: Sheet management (1 endpoint)

**Total:** 80 endpoints organized across 21 modules

**Integration:**
Routes are registered in src/api/router.py using FastAPI's include_router()
"""

# Routes will be imported and made available here
# from . import sync, rules, logs, sync_logs, metadata, ...

__all__ = []
