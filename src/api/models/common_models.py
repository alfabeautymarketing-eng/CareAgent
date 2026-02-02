from typing import Optional, Any
from pydantic import BaseModel, Field

class TaskStatusResponse(BaseModel):
    """Ответ со статусом фоновой задачи."""
    task_id: str = Field(..., description="ID задачи")
    status: str = Field(..., description="Статус: pending, running, completed, failed")
    current_phase: Optional[int] = Field(None, description="Текущая фаза")
    total_phases: Optional[int] = Field(None, description="Всего фаз")
    progress_percent: Optional[int] = Field(None, description="Процент выполнения")
    result: Optional[dict] = Field(None, description="Результат выполнения")
    error: Optional[str] = Field(None, description="Текст ошибки (если есть)")

    class Config:
        title = "Статус задачи"
