from typing import List, Optional, Any, Dict
from pydantic import BaseModel, Field

class MenuItemModel(BaseModel):
    """Элемент меню Google Таблицы."""
    label: Optional[str] = Field(None, description="Заголовок пункта")
    function_name: Optional[str] = Field(None, description="Имя функции GAS")
    separator: bool = Field(False, description="Разделитель")
    separator_after: bool = Field(False, description="Разделитель после")
    submenu: Optional[str] = Field(None, description="Подменю")
    items: Optional[List[dict]] = Field(None, description="Вложенные элементы")

    class Config:
        title = "Элемент меню"

class MenuGroupModel(BaseModel):
    """Группа элементов меню (подменю)."""
    title: str = Field(..., description="Название группы")
    items: List[MenuItemModel] = Field(..., description="Список элементов")

    class Config:
        title = "Группа меню"

class MenuConfigResponse(BaseModel):
    """Complete menu configuration for a project."""
    project: str = Field(..., description="Код проекта")
    menu_title: str = Field(..., description="Project menu title")
    order_sheet: str = Field(..., description="Order sheet name")
    sort_columns: Dict[str, str] = Field(..., description="Sortable columns mapping")
    primary_menu: MenuGroupModel = Field(..., description="Primary menu (🧾 Заказ)")
    order_stages_menu: MenuGroupModel = Field(..., description="Order stages menu (📊 Стадии)")
    menu_groups: List[MenuGroupModel] = Field(..., description="Static menu groups")

    class Config:
        title = "Menu Configuration"
