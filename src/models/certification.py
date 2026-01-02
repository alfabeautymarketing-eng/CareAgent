"""
Pydantic models for Certification entity - the canonical server-side representation.

Certifications link products to their regulatory documents (declarations, certificates).
This is the "source of truth" for certification data.
"""
from pydantic import BaseModel, Field, computed_field
from typing import Optional, List, Dict, Any
from datetime import datetime, date
from enum import Enum

from .price_models import ProjectCode


class CertificationType(str, Enum):
    """Type of certification document."""
    DECLARATION = "declaration"      # Декларация соответствия
    CERTIFICATE = "certificate"      # Сертификат соответствия
    REGISTRATION = "registration"    # Свидетельство о регистрации
    NOTIFICATION = "notification"    # Нотификация
    VOLUNTARY = "voluntary"          # Добровольная сертификация


class CertificationStatus(str, Enum):
    """Status of certification."""
    ACTIVE = "active"           # Valid and active
    EXPIRED = "expired"         # Past expiration date
    PENDING = "pending"         # Awaiting approval
    DRAFT = "draft"             # Not yet submitted
    REVOKED = "revoked"         # Revoked/cancelled


class CertificationDocument(BaseModel):
    """Reference to a certification document file."""
    file_id: str                    # Google Drive file ID
    filename: str
    mime_type: str = "application/pdf"
    uploaded_at: datetime = Field(default_factory=datetime.utcnow)
    page_count: Optional[int] = None


class ProductReference(BaseModel):
    """Reference to a product covered by this certification."""
    product_id: str                 # Product ID (e.g., MT-123)
    article: str
    name: str
    included_at: datetime = Field(default_factory=datetime.utcnow)


class Certification(BaseModel):
    """
    Canonical Certification entity.

    Represents a regulatory certification document that covers one or more products.
    Structure:
    - id: Unique certification ID (e.g., "CERT-MT-001")
    - project: Project code
    - document_number: Official document number
    - products: List of products covered
    - validity: Date range
    - documents: Attached files
    """

    # Primary identification
    id: str = Field(..., description="Unique certification ID")
    project: ProjectCode
    document_number: str = Field(..., description="Official document number")

    # Type and status
    cert_type: CertificationType = CertificationType.DECLARATION
    status: CertificationStatus = CertificationStatus.ACTIVE

    # Validity period
    issue_date: Optional[date] = None
    expiry_date: Optional[date] = None

    # Issuing authority
    issuing_authority: str = ""
    authority_code: str = ""  # e.g., registration number of the authority

    # Products covered
    products: List[ProductReference] = Field(default_factory=list)
    product_ids: List[str] = Field(default_factory=list)  # Quick lookup list

    # Documents
    documents: List[CertificationDocument] = Field(default_factory=list)
    main_document_id: Optional[str] = None  # Primary document file ID

    # Additional info
    notes: str = ""
    customs_codes: List[str] = Field(default_factory=list)  # ТН ВЭД codes

    # Metadata
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime = Field(default_factory=datetime.utcnow)
    version: int = 1
    source: str = "manual"

    @computed_field
    @property
    def is_valid(self) -> bool:
        """Check if certification is currently valid."""
        if self.status != CertificationStatus.ACTIVE:
            return False
        if self.expiry_date and self.expiry_date < date.today():
            return False
        return True

    @computed_field
    @property
    def days_until_expiry(self) -> Optional[int]:
        """Days until expiration, or None if no expiry date."""
        if not self.expiry_date:
            return None
        delta = self.expiry_date - date.today()
        return delta.days

    @computed_field
    @property
    def product_count(self) -> int:
        """Number of products covered."""
        return len(self.products)

    def add_product(self, product_id: str, article: str, name: str) -> None:
        """Add a product to this certification."""
        if product_id not in self.product_ids:
            self.products.append(ProductReference(
                product_id=product_id,
                article=article,
                name=name
            ))
            self.product_ids.append(product_id)
            self.updated_at = datetime.utcnow()

    def remove_product(self, product_id: str) -> bool:
        """Remove a product from this certification."""
        if product_id in self.product_ids:
            self.product_ids.remove(product_id)
            self.products = [p for p in self.products if p.product_id != product_id]
            self.updated_at = datetime.utcnow()
            return True
        return False

    def to_sheets_row(self, columns: List[str]) -> List[Any]:
        """Convert certification to a row for Google Sheets export."""
        flat = self._flatten()
        return [flat.get(col, "") for col in columns]

    def _flatten(self) -> Dict[str, Any]:
        """Flatten nested structure for sheets export."""
        return {
            "ID": self.id,
            "Номер документа": self.document_number,
            "Тип": self.cert_type.value,
            "Статус": self.status.value,
            "Дата выдачи": self.issue_date.isoformat() if self.issue_date else "",
            "Дата окончания": self.expiry_date.isoformat() if self.expiry_date else "",
            "Орган выдачи": self.issuing_authority,
            "Кол-во продуктов": self.product_count,
            "Продукты": ", ".join(self.product_ids),
            "Примечания": self.notes,
        }

    @classmethod
    def from_sheets_row(
        cls,
        row: Dict[str, Any],
        project: ProjectCode,
    ) -> "Certification":
        """Create Certification from Google Sheets row data."""
        cert_id = row.get("ID", row.get("id", ""))

        # Parse dates
        issue_date = None
        expiry_date = None

        if row.get("Дата выдачи"):
            try:
                issue_date = date.fromisoformat(str(row["Дата выдачи"]))
            except ValueError:
                pass

        if row.get("Дата окончания"):
            try:
                expiry_date = date.fromisoformat(str(row["Дата окончания"]))
            except ValueError:
                pass

        # Parse type
        cert_type = CertificationType.DECLARATION
        type_str = row.get("Тип", "").lower()
        if "сертификат" in type_str or "certificate" in type_str:
            cert_type = CertificationType.CERTIFICATE
        elif "регистрац" in type_str or "registration" in type_str:
            cert_type = CertificationType.REGISTRATION

        # Parse product IDs
        products_str = row.get("Продукты", row.get("product_ids", ""))
        product_ids = [p.strip() for p in str(products_str).split(",") if p.strip()]

        return cls(
            id=cert_id,
            project=project,
            document_number=row.get("Номер документа", row.get("document_number", "")),
            cert_type=cert_type,
            issue_date=issue_date,
            expiry_date=expiry_date,
            issuing_authority=row.get("Орган выдачи", row.get("issuing_authority", "")),
            product_ids=product_ids,
            notes=row.get("Примечания", row.get("notes", "")),
        )


class CertificationCollection(BaseModel):
    """Collection of certifications with metadata."""

    project: ProjectCode
    certifications: List[Certification] = Field(default_factory=list)
    last_sync: Optional[datetime] = None

    def get_by_document_number(self, doc_number: str) -> Optional[Certification]:
        """Find certification by document number."""
        for c in self.certifications:
            if c.document_number == doc_number:
                return c
        return None

    def get_by_id(self, cert_id: str) -> Optional[Certification]:
        """Find certification by ID."""
        for c in self.certifications:
            if c.id == cert_id:
                return c
        return None

    def get_for_product(self, product_id: str) -> List[Certification]:
        """Get all certifications covering a specific product."""
        return [c for c in self.certifications if product_id in c.product_ids]

    def get_expiring_soon(self, days: int = 30) -> List[Certification]:
        """Get certifications expiring within specified days."""
        result = []
        for c in self.certifications:
            if c.days_until_expiry is not None and 0 < c.days_until_expiry <= days:
                result.append(c)
        return result

    def get_expired(self) -> List[Certification]:
        """Get all expired certifications."""
        return [c for c in self.certifications if not c.is_valid]
