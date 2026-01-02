"""
Base storage abstractions for server-first architecture.
"""

from abc import ABC, abstractmethod
from typing import TypeVar, Generic, List, Optional, Dict, Any
from pydantic import BaseModel
from datetime import datetime


class StoreError(Exception):
    """Base exception for storage operations."""
    pass


T = TypeVar("T", bound=BaseModel)


class BaseStore(ABC, Generic[T]):
    """
    Abstract base class for data stores.

    Provides CRUD operations for Pydantic models with:
    - Versioning support
    - Audit trail
    - Conflict detection
    """

    @abstractmethod
    def get(self, id: str) -> Optional[T]:
        """Get entity by ID."""
        pass

    @abstractmethod
    def get_all(self, project: Optional[str] = None) -> List[T]:
        """Get all entities, optionally filtered by project."""
        pass

    @abstractmethod
    def find(self, query: Dict[str, Any]) -> List[T]:
        """Find entities matching query."""
        pass

    @abstractmethod
    def save(self, entity: T) -> T:
        """Save entity (create or update)."""
        pass

    @abstractmethod
    def save_many(self, entities: List[T]) -> List[T]:
        """Save multiple entities."""
        pass

    @abstractmethod
    def delete(self, id: str) -> bool:
        """Delete entity by ID."""
        pass

    @abstractmethod
    def exists(self, id: str) -> bool:
        """Check if entity exists."""
        pass

    # Versioning support

    @abstractmethod
    def get_version(self, id: str) -> int:
        """Get current version of entity."""
        pass

    @abstractmethod
    def get_history(self, id: str, limit: int = 10) -> List[Dict[str, Any]]:
        """Get version history for entity."""
        pass

    # Bulk operations

    @abstractmethod
    def import_from_sheets(self, data: List[Dict[str, Any]], project: str) -> int:
        """Import data from Google Sheets format."""
        pass

    @abstractmethod
    def export_to_sheets(self, project: str) -> List[List[Any]]:
        """Export data in Google Sheets format (2D array)."""
        pass


class VersionedEntity(BaseModel):
    """Mixin for versioned entities."""

    _version: int = 1
    _created_at: datetime = None
    _updated_at: datetime = None
    _created_by: Optional[str] = None
    _updated_by: Optional[str] = None

    class Config:
        underscore_attrs_are_private = True


class SyncMetadata(BaseModel):
    """Metadata for sync operations."""

    last_sync: Optional[datetime] = None
    sync_direction: str = "bidirectional"  # "to_sheets", "from_sheets", "bidirectional"
    conflict_strategy: str = "server_wins"  # "server_wins", "sheets_wins", "manual"
    sheets_row_map: Dict[str, int] = {}  # entity_id -> row_number
