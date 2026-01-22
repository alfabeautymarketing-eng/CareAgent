"""
Server-side sync journal storage and query service.
Stores entries in JSONL per spreadsheet.
"""

from __future__ import annotations

import json
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from uuid import uuid4

from filelock import FileLock
from pydantic import BaseModel, Field

from src.utils.logger import logger


class SyncLogEntry(BaseModel):
    id: str = Field(default_factory=lambda: uuid4().hex)
    timestamp: datetime = Field(default_factory=lambda: datetime.utcnow())
    spreadsheet_id: str
    project: Optional[str] = None
    row_key: Optional[str] = None
    source_info: str
    target_info: str
    old_value: Optional[str] = None
    new_value: Optional[str] = None
    category: Optional[str] = None
    status: Optional[str] = None
    event: Optional[str] = None
    tags: Optional[List[str]] = None
    rule_id: Optional[str] = None
    extra: Dict[str, Any] = Field(default_factory=dict)


from src.services.supabase_service import get_supabase_service

class SyncLogService:
    def __init__(
        self,
        data_dir: str = "data/sync_logs",
        retention_days: int = 90,
        max_entries: int = 200000,
    ):
        self.data_dir = Path(data_dir)
        self.retention_days = retention_days
        self.max_entries = max_entries
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.supabase = get_supabase_service()

    def add_entry(
        self,
        *,
        spreadsheet_id: str,
        source_info: str,
        target_info: str,
        row_key: Optional[str] = None,
        old_value: Optional[str] = None,
        new_value: Optional[str] = None,
        category: Optional[str] = None,
        status: Optional[str] = None,
        project: Optional[str] = None,
        event: Optional[str] = None,
        tags: Optional[List[str]] = None,
        rule_id: Optional[str] = None,
        extra: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        entry = SyncLogEntry(
            spreadsheet_id=spreadsheet_id,
            project=project,
            row_key=row_key,
            source_info=source_info,
            target_info=target_info,
            old_value=old_value,
            new_value=new_value,
            category=category,
            status=status,
            event=event,
            tags=tags,
            rule_id=rule_id,
            extra=extra or {},
        )
        payload = entry.model_dump(mode="json")
        path = self._log_path(spreadsheet_id)
        lock = self._lock_path(spreadsheet_id)

        # 1. Write to local file
        try:
            with FileLock(lock):
                with open(path, "a", encoding="utf-8") as handle:
                    handle.write(json.dumps(payload, ensure_ascii=False) + "\n")
            self._enforce_retention(path, lock)
        except Exception as exc:
            logger.error("sync_log_write_failed", error=str(exc))

        # 2. Write to Supabase (if configured)
        if self.supabase.is_configured():
            try:
                # We use the raw client to insert
                self.supabase.client.table("sync_logs").insert(payload).execute()
            except Exception as exc:
                # We don't want to break the main flow if supabase logging fails
                logger.error("supabase_log_write_failed", error=str(exc))

        return payload

    def list_entries(
        self,
        *,
        spreadsheet_id: str,
        limit: int = 200,
        offset: int = 0,
        order: str = "desc",
        project: Optional[str] = None,
        category: Optional[str] = None,
        status: Optional[str] = None,
        row_key: Optional[str] = None,
        source: Optional[str] = None,
        target: Optional[str] = None,
        rule_id: Optional[str] = None,
        start: Optional[datetime] = None,
        end: Optional[datetime] = None,
    ) -> Tuple[List[Dict[str, Any]], int]:
        path = self._log_path(spreadsheet_id)
        lock = self._lock_path(spreadsheet_id)

        if not path.exists():
            return [], 0

        entries: List[Dict[str, Any]] = []

        with FileLock(lock):
            with open(path, "r", encoding="utf-8") as handle:
                for line in handle:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        entry = json.loads(line)
                    except json.JSONDecodeError:
                        continue

                    if project and entry.get("project") != project:
                        continue
                    if category and entry.get("category") != category:
                        continue
                    if status and entry.get("status") != status:
                        continue
                    if row_key and entry.get("row_key") != row_key:
                        continue
                    if rule_id and entry.get("rule_id") != rule_id:
                        continue

                    source_info = entry.get("source_info", "")
                    target_info = entry.get("target_info", "")
                    if source and source.lower() not in str(source_info).lower():
                        continue
                    if target and target.lower() not in str(target_info).lower():
                        continue

                    timestamp = self._parse_timestamp(entry.get("timestamp"))
                    if start and timestamp and timestamp < start:
                        continue
                    if end and timestamp and timestamp > end:
                        continue

                    entries.append(entry)

        total = len(entries)
        reverse = order.lower() != "asc"
        entries.sort(
            key=lambda item: self._parse_timestamp(item.get("timestamp")) or datetime.min,
            reverse=reverse,
        )

        start_idx = max(offset, 0)
        end_idx = start_idx + max(limit, 0)
        return entries[start_idx:end_idx], total

    def _log_path(self, spreadsheet_id: str) -> Path:
        safe_id = spreadsheet_id.replace("/", "_")
        return self.data_dir / f"{safe_id}.jsonl"

    def _lock_path(self, spreadsheet_id: str) -> str:
        return str(self._log_path(spreadsheet_id).with_suffix(".lock"))

    def _parse_timestamp(self, value: Any) -> Optional[datetime]:
        if not value:
            return None
        if isinstance(value, datetime):
            return value
        try:
            return datetime.fromisoformat(str(value))
        except ValueError:
            return None

    def _enforce_retention(self, path: Path, lock_path: str) -> None:
        if self.retention_days <= 0 and self.max_entries <= 0:
            return

        try:
            with FileLock(lock_path):
                with open(path, "r", encoding="utf-8") as handle:
                    lines = [line for line in handle if line.strip()]

                if not lines:
                    return

                cutoff = None
                if self.retention_days > 0:
                    cutoff = datetime.utcnow() - timedelta(days=self.retention_days)

                filtered = []
                for line in lines:
                    try:
                        entry = json.loads(line)
                    except json.JSONDecodeError:
                        continue
                    if cutoff:
                        ts = self._parse_timestamp(entry.get("timestamp"))
                        if ts and ts < cutoff:
                            continue
                    filtered.append(entry)

                if self.max_entries > 0 and len(filtered) > self.max_entries:
                    filtered = filtered[-self.max_entries:]

                with open(path, "w", encoding="utf-8") as handle:
                    for entry in filtered:
                        handle.write(json.dumps(entry, ensure_ascii=False) + "\n")
        except Exception as exc:
            logger.warning("sync_log_retention_failed", error=str(exc))
