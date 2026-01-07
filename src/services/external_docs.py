import os
import yaml
from typing import List, Dict, Any, Optional
from pathlib import Path
from src.utils.logger import logger

class ExternalDoc(Dict):
    """Model for external document registry entry."""
    pass

class ExternalDocsService:
    def __init__(self, config_path: str = "config/external_docs.yaml"):
        self.config_path = Path(config_path)
        self._ensure_config_exists()

    def _ensure_config_exists(self):
        """Initializes empty config if not exists."""
        if not self.config_path.exists():
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            self._save_docs([])

    def _load_docs(self) -> List[Dict[str, Any]]:
        """Loads docs from YAML."""
        try:
            if not self.config_path.exists():
                return []
            with open(self.config_path, "r", encoding="utf-8") as f:
                return yaml.safe_load(f) or []
        except Exception as e:
            logger.error("external_docs_load_failed", error=str(e))
            return []

    def _save_docs(self, docs: List[Dict[str, Any]]):
        """Saves docs to YAML."""
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                yaml.safe_dump(docs, f, allow_unicode=True, sort_keys=False)
        except Exception as e:
            logger.error("external_docs_save_failed", error=str(e))
            raise

    def list_docs(self) -> List[Dict[str, Any]]:
        """Returns all registered external documents."""
        return self._load_docs()

    def add_doc(self, name: str, doc_id: str) -> Dict[str, Any]:
        """Adds a new external document or updates existing by doc_id."""
        docs = self._load_docs()
        
        # Check if already exists
        for doc in docs:
            if doc.get("doc_id") == doc_id:
                doc["name"] = name
                self._save_docs(docs)
                logger.info("external_doc_updated", doc_id=doc_id, name=name)
                return doc
        
        new_doc = {
            "name": name,
            "doc_id": doc_id,
            "added_at": os.popen("date +%Y-%m-%dT%H:%M:%S").read().strip() # Simple timestamp
        }
        docs.append(new_doc)
        self._save_docs(docs)
        logger.info("external_doc_added", doc_id=doc_id, name=name)
        return new_doc

    def remove_doc(self, doc_id: str) -> bool:
        """Removes an external document by its ID."""
        docs = self._load_docs()
        new_docs = [d for d in docs if d.get("doc_id") != doc_id]
        
        if len(new_docs) == len(docs):
            return False
            
        self._save_docs(new_docs)
        logger.info("external_doc_removed", doc_id=doc_id)
        return True
