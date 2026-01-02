"""
JSON file-based storage implementation.

Simple, file-based storage for development and small datasets.
Each project gets its own directory with JSON files per entity type.
"""

import json
import os
from pathlib import Path
from typing import TypeVar, Type, List, Optional, Dict, Any
from datetime import datetime
from pydantic import BaseModel
from filelock import FileLock

from .base import BaseStore, StoreError
from src.utils.logger import logger


T = TypeVar("T", bound=BaseModel)


class JSONStore(BaseStore[T]):
    """
    JSON file-based store implementation.

    Directory structure:
    data/
    ├── mt/
    │   ├── products.json
    │   ├── prices.json
    │   └── certifications.json
    ├── sk/
    │   └── ...
    └── ss/
        └── ...
    """

    def __init__(
        self,
        model_class: Type[T],
        entity_name: str,
        data_dir: str = "data",
        id_field: str = "id"
    ):
        self.model_class = model_class
        self.entity_name = entity_name
        self.data_dir = Path(data_dir)
        self.id_field = id_field
        self._cache: Dict[str, Dict[str, T]] = {}  # project -> {id -> entity}
        self._versions: Dict[str, Dict[str, int]] = {}  # project -> {id -> version}

        # Ensure data directory exists
        self.data_dir.mkdir(parents=True, exist_ok=True)

    def _get_file_path(self, project: str) -> Path:
        """Get JSON file path for project."""
        project_dir = self.data_dir / project
        project_dir.mkdir(parents=True, exist_ok=True)
        return project_dir / f"{self.entity_name}.json"

    def _get_lock_path(self, project: str) -> Path:
        """Get lock file path for project."""
        return self._get_file_path(project).with_suffix(".lock")

    def _load_project(self, project: str) -> Dict[str, T]:
        """Load all entities for project from JSON file."""
        if project in self._cache:
            return self._cache[project]

        file_path = self._get_file_path(project)

        if not file_path.exists():
            self._cache[project] = {}
            self._versions[project] = {}
            return {}

        try:
            with FileLock(self._get_lock_path(project)):
                with open(file_path, "r", encoding="utf-8") as f:
                    data = json.load(f)

            entities = {}
            versions = {}

            for item in data.get("entities", []):
                entity = self.model_class(**item["data"])
                entity_id = getattr(entity, self.id_field)
                entities[entity_id] = entity
                versions[entity_id] = item.get("version", 1)

            self._cache[project] = entities
            self._versions[project] = versions
            return entities

        except Exception as e:
            logger.error(f"Failed to load {self.entity_name} for {project}: {e}")
            raise StoreError(f"Failed to load data: {e}")

    def _save_project(self, project: str):
        """Save all entities for project to JSON file."""
        if project not in self._cache:
            return

        file_path = self._get_file_path(project)

        data = {
            "entity_type": self.entity_name,
            "project": project,
            "updated_at": datetime.utcnow().isoformat(),
            "count": len(self._cache[project]),
            "entities": [
                {
                    "id": entity_id,
                    "version": self._versions.get(project, {}).get(entity_id, 1),
                    "data": entity.model_dump()
                }
                for entity_id, entity in self._cache[project].items()
            ]
        }

        try:
            with FileLock(self._get_lock_path(project)):
                with open(file_path, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2, default=str)
        except Exception as e:
            logger.error(f"Failed to save {self.entity_name} for {project}: {e}")
            raise StoreError(f"Failed to save data: {e}")

    def _extract_project(self, entity: T) -> str:
        """Extract project code from entity."""
        # Try common project field names
        for field in ["project", "project_code"]:
            if hasattr(entity, field):
                return getattr(entity, field)

        # Try to extract from ID (e.g., "MT-123" -> "mt")
        entity_id = getattr(entity, self.id_field, "")
        if "-" in str(entity_id):
            return str(entity_id).split("-")[0].lower()

        return "common"

    # CRUD operations

    def get(self, id: str, project: Optional[str] = None) -> Optional[T]:
        """Get entity by ID."""
        if project:
            entities = self._load_project(project)
            return entities.get(id)

        # Search across all projects
        for proj in ["mt", "sk", "ss", "common"]:
            entities = self._load_project(proj)
            if id in entities:
                return entities[id]
        return None

    def get_all(self, project: Optional[str] = None) -> List[T]:
        """Get all entities, optionally filtered by project."""
        if project:
            return list(self._load_project(project).values())

        # Get from all projects
        all_entities = []
        for proj in ["mt", "sk", "ss", "common"]:
            all_entities.extend(self._load_project(proj).values())
        return all_entities

    def find(self, query: Dict[str, Any], project: Optional[str] = None) -> List[T]:
        """Find entities matching query."""
        entities = self.get_all(project)
        results = []

        for entity in entities:
            match = True
            for key, value in query.items():
                if hasattr(entity, key):
                    entity_value = getattr(entity, key)
                    if entity_value != value:
                        match = False
                        break
                else:
                    match = False
                    break

            if match:
                results.append(entity)

        return results

    def save(self, entity: T) -> T:
        """Save entity (create or update)."""
        project = self._extract_project(entity)
        entities = self._load_project(project)

        entity_id = getattr(entity, self.id_field)

        # Update version
        if project not in self._versions:
            self._versions[project] = {}

        current_version = self._versions[project].get(entity_id, 0)
        self._versions[project][entity_id] = current_version + 1

        # Save to cache
        entities[entity_id] = entity
        self._cache[project] = entities

        # Persist
        self._save_project(project)

        return entity

    def save_many(self, entities: List[T]) -> List[T]:
        """Save multiple entities."""
        # Group by project
        by_project: Dict[str, List[T]] = {}
        for entity in entities:
            project = self._extract_project(entity)
            if project not in by_project:
                by_project[project] = []
            by_project[project].append(entity)

        # Save each project
        saved = []
        for project, project_entities in by_project.items():
            loaded = self._load_project(project)

            for entity in project_entities:
                entity_id = getattr(entity, self.id_field)
                loaded[entity_id] = entity

                if project not in self._versions:
                    self._versions[project] = {}
                current = self._versions[project].get(entity_id, 0)
                self._versions[project][entity_id] = current + 1

                saved.append(entity)

            self._cache[project] = loaded
            self._save_project(project)

        return saved

    def delete(self, id: str, project: Optional[str] = None) -> bool:
        """Delete entity by ID."""
        if project:
            entities = self._load_project(project)
            if id in entities:
                del entities[id]
                self._cache[project] = entities
                self._save_project(project)
                return True
            return False

        # Search across projects
        for proj in ["mt", "sk", "ss", "common"]:
            entities = self._load_project(proj)
            if id in entities:
                del entities[id]
                self._cache[proj] = entities
                self._save_project(proj)
                return True
        return False

    def exists(self, id: str, project: Optional[str] = None) -> bool:
        """Check if entity exists."""
        return self.get(id, project) is not None

    # Versioning

    def get_version(self, id: str, project: Optional[str] = None) -> int:
        """Get current version of entity."""
        if project and project in self._versions:
            return self._versions[project].get(id, 0)

        for proj in ["mt", "sk", "ss", "common"]:
            if proj in self._versions and id in self._versions[proj]:
                return self._versions[proj][id]
        return 0

    def get_history(self, id: str, limit: int = 10) -> List[Dict[str, Any]]:
        """Get version history (not implemented for JSON store)."""
        # JSON store doesn't keep history, would need separate history file
        return []

    # Bulk operations for Sheets sync

    def import_from_sheets(
        self,
        data: List[Dict[str, Any]],
        project: str,
        clear_existing: bool = False
    ) -> int:
        """Import data from Google Sheets format."""
        if clear_existing:
            self._cache[project] = {}
            self._versions[project] = {}

        count = 0
        for row in data:
            try:
                entity = self.model_class(**row)
                self.save(entity)
                count += 1
            except Exception as e:
                logger.warning(f"Failed to import row: {e}")

        return count

    def export_to_sheets(self, project: str, columns: Optional[List[str]] = None) -> List[List[Any]]:
        """Export data in Google Sheets format (2D array with headers)."""
        entities = self.get_all(project)

        if not entities:
            return []

        # Determine columns from first entity if not specified
        if not columns:
            columns = list(entities[0].model_dump().keys())

        # Build 2D array
        result = [columns]  # Header row

        for entity in entities:
            data = entity.model_dump()
            row = [data.get(col, "") for col in columns]
            result.append(row)

        return result

    def clear_cache(self, project: Optional[str] = None):
        """Clear in-memory cache."""
        if project:
            self._cache.pop(project, None)
            self._versions.pop(project, None)
        else:
            self._cache.clear()
            self._versions.clear()
