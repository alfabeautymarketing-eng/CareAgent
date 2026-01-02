"""
Pydantic models for Order entity - the canonical server-side representation.

Orders track the lifecycle of product orders from creation to delivery.
This is the "source of truth" for order data.
"""
from pydantic import BaseModel, Field, computed_field
from typing import Optional, List, Dict, Any
from datetime import datetime, date
from decimal import Decimal
from enum import Enum

from .price_models import ProjectCode


class OrderStage(str, Enum):
    """Order lifecycle stages."""
    DRAFT = "draft"                     # Черновик
    SUBMITTED = "submitted"             # Отправлен поставщику
    CONFIRMED = "confirmed"             # Подтверждён
    PRODUCTION = "production"           # В производстве
    SHIPPED = "shipped"                 # Отгружен
    IN_TRANSIT = "in_transit"           # В пути
    CUSTOMS = "customs"                 # На таможне
    CUSTOMS_CLEARED = "customs_cleared" # Растаможен
    WAREHOUSE = "warehouse"             # На складе
    DELIVERED = "delivered"             # Доставлен
    CANCELLED = "cancelled"             # Отменён


class OrderItemStatus(str, Enum):
    """Status of individual order item."""
    PENDING = "pending"
    CONFIRMED = "confirmed"
    PARTIAL = "partial"      # Partially fulfilled
    FULFILLED = "fulfilled"
    CANCELLED = "cancelled"


class OrderItem(BaseModel):
    """Single item in an order."""
    product_id: str
    article: str
    name: str
    quantity_ordered: int
    quantity_confirmed: int = 0
    quantity_received: int = 0
    unit_price: float = 0.0
    status: OrderItemStatus = OrderItemStatus.PENDING

    @computed_field
    @property
    def total_price(self) -> float:
        """Total price for this item."""
        return self.quantity_ordered * self.unit_price

    @computed_field
    @property
    def fulfillment_rate(self) -> float:
        """Percentage of order fulfilled (0-100)."""
        if self.quantity_ordered == 0:
            return 0.0
        return (self.quantity_received / self.quantity_ordered) * 100


class StageTransition(BaseModel):
    """Record of stage change."""
    from_stage: Optional[OrderStage] = None
    to_stage: OrderStage
    changed_at: datetime = Field(default_factory=datetime.utcnow)
    changed_by: Optional[str] = None
    notes: str = ""


class ShippingInfo(BaseModel):
    """Shipping and logistics information."""
    carrier: str = ""
    tracking_number: str = ""
    ship_date: Optional[date] = None
    estimated_arrival: Optional[date] = None
    actual_arrival: Optional[date] = None
    customs_declaration: str = ""
    customs_date: Optional[date] = None


class Order(BaseModel):
    """
    Canonical Order entity.

    Represents an order to a supplier with full lifecycle tracking.
    Structure:
    - id: Unique order ID (e.g., "ORD-MT-2024-001")
    - project: Project code
    - items: List of ordered products
    - stage: Current lifecycle stage
    - history: Stage transitions
    """

    # Primary identification
    id: str = Field(..., description="Unique order ID")
    project: ProjectCode
    order_number: str = ""  # Supplier's order number

    # Supplier info
    supplier_name: str = ""
    supplier_contact: str = ""

    # Order content
    items: List[OrderItem] = Field(default_factory=list)

    # Status
    stage: OrderStage = OrderStage.DRAFT
    stage_history: List[StageTransition] = Field(default_factory=list)

    # Dates
    created_date: date = Field(default_factory=date.today)
    submitted_date: Optional[date] = None
    expected_delivery: Optional[date] = None

    # Shipping
    shipping: ShippingInfo = Field(default_factory=ShippingInfo)

    # Financial
    currency: str = "USD"
    subtotal: float = 0.0
    shipping_cost: float = 0.0
    customs_duties: float = 0.0
    total_cost: float = 0.0

    # References
    invoice_id: Optional[str] = None
    related_orders: List[str] = Field(default_factory=list)

    # Notes
    notes: str = ""
    internal_notes: str = ""

    # Metadata
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime = Field(default_factory=datetime.utcnow)
    version: int = 1

    @computed_field
    @property
    def item_count(self) -> int:
        """Total number of line items."""
        return len(self.items)

    @computed_field
    @property
    def total_quantity(self) -> int:
        """Total quantity of all items."""
        return sum(item.quantity_ordered for item in self.items)

    @computed_field
    @property
    def is_active(self) -> bool:
        """Check if order is still active (not delivered/cancelled)."""
        return self.stage not in [OrderStage.DELIVERED, OrderStage.CANCELLED]

    @computed_field
    @property
    def days_since_created(self) -> int:
        """Days since order was created."""
        delta = date.today() - self.created_date
        return delta.days

    def transition_to(self, new_stage: OrderStage, notes: str = "", changed_by: str = None) -> None:
        """
        Transition order to a new stage.

        Args:
            new_stage: Target stage
            notes: Optional notes about the transition
            changed_by: User/system making the change
        """
        transition = StageTransition(
            from_stage=self.stage,
            to_stage=new_stage,
            notes=notes,
            changed_by=changed_by
        )
        self.stage_history.append(transition)
        self.stage = new_stage
        self.updated_at = datetime.utcnow()

        # Update relevant dates based on stage
        if new_stage == OrderStage.SUBMITTED:
            self.submitted_date = date.today()
        elif new_stage == OrderStage.SHIPPED:
            self.shipping.ship_date = date.today()
        elif new_stage == OrderStage.DELIVERED:
            self.shipping.actual_arrival = date.today()

    def add_item(
        self,
        product_id: str,
        article: str,
        name: str,
        quantity: int,
        unit_price: float = 0.0
    ) -> OrderItem:
        """Add an item to the order."""
        item = OrderItem(
            product_id=product_id,
            article=article,
            name=name,
            quantity_ordered=quantity,
            unit_price=unit_price
        )
        self.items.append(item)
        self._recalculate_totals()
        return item

    def remove_item(self, product_id: str) -> bool:
        """Remove an item from the order."""
        original_count = len(self.items)
        self.items = [i for i in self.items if i.product_id != product_id]
        if len(self.items) < original_count:
            self._recalculate_totals()
            return True
        return False

    def _recalculate_totals(self) -> None:
        """Recalculate order totals."""
        self.subtotal = sum(item.total_price for item in self.items)
        self.total_cost = self.subtotal + self.shipping_cost + self.customs_duties

    def to_sheets_row(self, columns: List[str]) -> List[Any]:
        """Convert order to a row for Google Sheets export."""
        flat = self._flatten()
        return [flat.get(col, "") for col in columns]

    def _flatten(self) -> Dict[str, Any]:
        """Flatten nested structure for sheets export."""
        return {
            "ID": self.id,
            "Номер заказа": self.order_number,
            "Проект": self.project.value,
            "Поставщик": self.supplier_name,
            "Стадия": self.stage.value,
            "Дата создания": self.created_date.isoformat(),
            "Дата отправки": self.submitted_date.isoformat() if self.submitted_date else "",
            "Ожидаемая доставка": self.expected_delivery.isoformat() if self.expected_delivery else "",
            "Кол-во позиций": self.item_count,
            "Общее кол-во": self.total_quantity,
            "Сумма": self.subtotal,
            "Итого": self.total_cost,
            "Валюта": self.currency,
            "Примечания": self.notes,
        }

    @classmethod
    def from_sheets_row(cls, row: Dict[str, Any], project: ProjectCode) -> "Order":
        """Create Order from Google Sheets row data."""
        order_id = row.get("ID", row.get("id", ""))

        # Parse dates
        created_date = date.today()
        if row.get("Дата создания"):
            try:
                created_date = date.fromisoformat(str(row["Дата создания"]))
            except ValueError:
                pass

        # Parse stage
        stage = OrderStage.DRAFT
        stage_str = row.get("Стадия", "").lower()
        for s in OrderStage:
            if s.value == stage_str:
                stage = s
                break

        return cls(
            id=order_id,
            project=project,
            order_number=row.get("Номер заказа", row.get("order_number", "")),
            supplier_name=row.get("Поставщик", row.get("supplier_name", "")),
            stage=stage,
            created_date=created_date,
            notes=row.get("Примечания", row.get("notes", "")),
        )


class OrderCollection(BaseModel):
    """Collection of orders with metadata."""

    project: ProjectCode
    orders: List[Order] = Field(default_factory=list)
    last_sync: Optional[datetime] = None

    def get_by_id(self, order_id: str) -> Optional[Order]:
        """Find order by ID."""
        for o in self.orders:
            if o.id == order_id:
                return o
        return None

    def get_by_stage(self, stage: OrderStage) -> List[Order]:
        """Get all orders at a specific stage."""
        return [o for o in self.orders if o.stage == stage]

    def get_active(self) -> List[Order]:
        """Get all active orders."""
        return [o for o in self.orders if o.is_active]

    def get_for_product(self, product_id: str) -> List[Order]:
        """Get all orders containing a specific product."""
        result = []
        for order in self.orders:
            if any(item.product_id == product_id for item in order.items):
                result.append(order)
        return result
