from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from typing import Any


def _now() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


@dataclass
class StageRecord:
    name: str
    status: str
    startedAt: str
    finishedAt: str
    output: dict[str, Any] = field(default_factory=dict)
    error: dict[str, str] | None = None


class PipelineRun:
    def __init__(self) -> None:
        self.started_at = _now()
        self.stages: list[StageRecord] = []
        self.errors: list[dict[str, str]] = []
        self.current_stage: str | None = None
        self.status = "running"

    def complete(self, name: str, output: dict[str, Any] | None = None) -> None:
        timestamp = _now()
        self.current_stage = name
        self.stages.append(
            StageRecord(
                name=name,
                status="completed",
                startedAt=timestamp,
                finishedAt=timestamp,
                output=output or {},
            )
        )

    def fail(self, name: str, code: str, message: str, suggestion: str) -> None:
        timestamp = _now()
        error = {"stage": name, "code": code, "message": message, "suggestion": suggestion}
        self.current_stage = name
        self.status = "needs-fix"
        self.errors.append(error)
        self.stages.append(
            StageRecord(
                name=name,
                status="failed",
                startedAt=timestamp,
                finishedAt=timestamp,
                error={key: value for key, value in error.items() if key != "stage"},
            )
        )

    def finish(self) -> None:
        if not self.errors:
            self.status = "completed"

    def to_dict(self) -> dict[str, Any]:
        return {
            "schemaVersion": "1.0",
            "status": self.status,
            "startedAt": self.started_at,
            "currentStage": self.current_stage,
            "stages": [asdict(stage) for stage in self.stages],
            "errors": self.errors,
        }
