"""
Detailed function execution step logging service.
Tracks individual steps of function execution with context and status.
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional
from uuid import uuid4

from filelock import FileLock
from pydantic import BaseModel, Field

from src.utils.logger import logger


class FunctionStep(BaseModel):
    """Single step in function execution."""

    id: str = Field(default_factory=lambda: uuid4().hex)
    execution_id: str  # Links all steps of one function call
    timestamp: datetime = Field(default_factory=lambda: datetime.utcnow())

    # Function context
    module: str  # e.g., "src.services.sync"
    function_name: str  # e.g., "sync_event"

    # Step details
    step_number: int
    step_name: str  # e.g., "validate_event", "acquire_lock", "write_target"
    level: str  # "info", "warning", "error", "debug"
    message: str

    # Status
    status: str  # "started", "completed", "failed", "skipped"
    duration_ms: Optional[float] = None  # Time in milliseconds

    # Context
    row_key: Optional[str] = None
    spreadsheet_id: Optional[str] = None
    rule_id: Optional[str] = None

    # Additional data
    data: Dict[str, Any] = Field(default_factory=dict)
    error: Optional[str] = None


class FunctionLogService:
    """Service for tracking detailed function execution steps."""

    def __init__(
        self,
        data_dir: str = "data/function_logs",
        retention_days: int = 30,
    ):
        self.data_dir = Path(data_dir)
        self.retention_days = retention_days
        self.data_dir.mkdir(parents=True, exist_ok=True)

        # In-memory execution tracking for current operations
        self._executions: Dict[str, List[FunctionStep]] = {}

    def start_execution(self, execution_id: Optional[str] = None) -> str:
        """Start tracking a new function execution."""
        exec_id = execution_id or uuid4().hex
        self._executions[exec_id] = []
        return exec_id

    def log_step(
        self,
        *,
        execution_id: str,
        module: str,
        function_name: str,
        step_number: int,
        step_name: str,
        level: str = "info",
        message: str,
        status: str = "started",
        duration_ms: Optional[float] = None,
        row_key: Optional[str] = None,
        spreadsheet_id: Optional[str] = None,
        rule_id: Optional[str] = None,
        data: Optional[Dict[str, Any]] = None,
        error: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Log a single step in function execution."""
        step = FunctionStep(
            execution_id=execution_id,
            module=module,
            function_name=function_name,
            step_number=step_number,
            step_name=step_name,
            level=level,
            message=message,
            status=status,
            duration_ms=duration_ms,
            row_key=row_key,
            spreadsheet_id=spreadsheet_id,
            rule_id=rule_id,
            data=data or {},
            error=error,
        )

        # Keep in memory for this execution
        if execution_id in self._executions:
            self._executions[execution_id].append(step)

        # Log immediately
        self._write_step(step)

        return step.model_dump(mode="json")

    def end_execution(
        self,
        execution_id: str,
        total_duration_ms: Optional[float] = None,
        status: str = "completed",
    ) -> Dict[str, Any]:
        """Mark execution as complete and save to disk."""
        if execution_id not in self._executions:
            logger.warning("execution_not_found", execution_id=execution_id)
            return {}

        steps = self._executions[execution_id]

        # Get summary
        summary = {
            "execution_id": execution_id,
            "status": status,
            "total_duration_ms": total_duration_ms,
            "step_count": len(steps),
            "timestamp": datetime.utcnow().isoformat(),
            "steps": [s.model_dump(mode="json") for s in steps],
        }

        # Save to file
        try:
            self._write_execution(summary)
        except Exception as e:
            logger.error("execution_save_failed", error=str(e), execution_id=execution_id)

        # Clean up from memory
        del self._executions[execution_id]

        return summary

    def get_execution(self, execution_id: str) -> Optional[Dict[str, Any]]:
        """Get execution summary if it exists in memory."""
        if execution_id in self._executions:
            steps = self._executions[execution_id]
            return {
                "execution_id": execution_id,
                "step_count": len(steps),
                "steps": [s.model_dump(mode="json") for s in steps],
            }
        return None

    def list_executions(
        self,
        *,
        module: Optional[str] = None,
        function_name: Optional[str] = None,
        limit: int = 100,
        offset: int = 0,
    ) -> tuple[List[Dict[str, Any]], int]:
        """List past executions from disk."""
        log_dir = self.data_dir

        if not log_dir.exists():
            return [], 0

        executions = []

        try:
            for log_file in sorted(log_dir.glob("*.jsonl"), reverse=True):
                if len(executions) >= limit + offset:
                    break

                with open(log_file, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if not line:
                            continue

                        try:
                            exec_data = json.loads(line)
                        except json.JSONDecodeError:
                            continue

                        # Filter
                        first_step = exec_data.get("steps", [{}])[0]
                        if module and first_step.get("module") != module:
                            continue
                        if function_name and first_step.get("function_name") != function_name:
                            continue

                        executions.append(exec_data)

        except Exception as e:
            logger.warning("execution_list_failed", error=str(e))

        total = len(executions)
        return executions[offset:offset+limit], total

    def _write_step(self, step: FunctionStep) -> None:
        """Write a single step to the log file."""
        path = self.data_dir / "steps.jsonl"
        lock = self.data_dir / "steps.lock"

        payload = step.model_dump(mode="json")

        try:
            with FileLock(str(lock)):
                with open(path, "a", encoding="utf-8") as f:
                    f.write(json.dumps(payload, ensure_ascii=False) + "\n")
        except Exception as e:
            logger.error("step_write_failed", error=str(e))

    def _write_execution(self, summary: Dict[str, Any]) -> None:
        """Write complete execution to file."""
        path = self.data_dir / "executions.jsonl"
        lock = self.data_dir / "executions.lock"

        try:
            with FileLock(str(lock)):
                with open(path, "a", encoding="utf-8") as f:
                    f.write(json.dumps(summary, ensure_ascii=False) + "\n")
        except Exception as e:
            logger.error("execution_write_failed", error=str(e))


class FunctionLogContext:
    """Context manager for logging function execution."""

    def __init__(
        self,
        service: FunctionLogService,
        module: str,
        function_name: str,
        execution_id: Optional[str] = None,
    ):
        self.service = service
        self.module = module
        self.function_name = function_name
        self.execution_id = execution_id or uuid4().hex

        self._start_time = None
        self._step_count = 0

    def __enter__(self):
        self._start_time = datetime.utcnow()
        self.service.start_execution(self.execution_id)

        self.service.log_step(
            execution_id=self.execution_id,
            module=self.module,
            function_name=self.function_name,
            step_number=0,
            step_name="start",
            message=f"Starting {self.function_name}",
            status="started",
        )

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        from datetime import datetime as dt

        duration_ms = (dt.utcnow() - self._start_time).total_seconds() * 1000
        status = "failed" if exc_type else "completed"

        self.service.end_execution(
            self.execution_id,
            total_duration_ms=duration_ms,
            status=status,
        )

        return False

    def log(
        self,
        step_name: str,
        message: str,
        level: str = "info",
        status: str = "completed",
        duration_ms: Optional[float] = None,
        data: Optional[Dict[str, Any]] = None,
        error: Optional[str] = None,
        **kwargs
    ):
        """Log a step within this execution."""
        self._step_count += 1

        self.service.log_step(
            execution_id=self.execution_id,
            module=self.module,
            function_name=self.function_name,
            step_number=self._step_count,
            step_name=step_name,
            level=level,
            message=message,
            status=status,
            duration_ms=duration_ms,
            data=data or kwargs,
            error=error,
        )
