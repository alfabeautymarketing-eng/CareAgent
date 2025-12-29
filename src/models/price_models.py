"""
Pydantic models for price processing
"""
from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any
from enum import Enum


class ProjectCode(str, Enum):
    MT = "mt"
    SK = "sk"
    SS = "ss"


class ProcessingMode(str, Enum):
    MAIN = "main"
    TESTER = "tester"
    SAMPLES = "samples"
    PROBES = "probes"


class PriceProcessRequest(BaseModel):
    """Request model for price processing endpoint"""
    spreadsheet_id: str = Field(..., description="Target spreadsheet ID")
    mode: ProcessingMode = Field(default=ProcessingMode.MAIN, description="Processing mode")
    dry_run: bool = Field(default=False, description="If true, returns preview without writing")
    source_doc_id: Optional[str] = Field(default=None, description="Source document ID (optional, uses config if not provided)")


class ProcessedRow(BaseModel):
    """Single processed row from price list"""
    id_p: str = ""
    article: str
    name_eng: str
    volume: str = ""
    barcode: str = ""
    units_per_pack: str = ""
    price: float = 0.0
    group: str = ""


class ProcessedData(BaseModel):
    """Result of price list parsing"""
    headers: List[str]
    rows: List[List[Any]]
    articles: List[str]
    groups: List[str]


class SyncResult(BaseModel):
    """Result of syncing with main sheet"""
    assigned_idp: Dict[str, str] = Field(default_factory=dict)
    created_rows: List[int] = Field(default_factory=list)
    updated_rows: List[int] = Field(default_factory=list)
    group_changes: List[Dict[str, Any]] = Field(default_factory=list)
    barcode_mismatches: List[Dict[str, Any]] = Field(default_factory=list)


class PriceProcessResponse(BaseModel):
    """Response model for price processing endpoint"""
    status: str
    message: str
    processed_rows: int = 0
    groups_found: int = 0
    new_articles: int = 0
    updated_articles: int = 0
    errors: List[str] = Field(default_factory=list)
    preview: Optional[List[Dict[str, Any]]] = None


class StockLoadRequest(BaseModel):
    """Request model for stock loading endpoint"""
    spreadsheet_id: str
    source_doc_id: Optional[str] = None


class StockLoadResponse(BaseModel):
    """Response model for stock loading endpoint"""
    status: str
    message: str
    updated_rows: int = 0
    source_doc: str = ""
