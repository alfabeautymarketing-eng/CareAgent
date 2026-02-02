from src.services.sync import SyncService
from src.services.logging_service import logging_service
from src.services.sync_log_service import SyncLogService
from src.utils.config import settings

# Initialize sync log service
sync_log_service = SyncLogService(
    data_dir=settings.sync_log_data_dir,
    retention_days=settings.sync_log_retention_days,
    max_entries=settings.sync_log_max_entries,
)

sync_service = SyncService(logging_service, sync_log_service)
