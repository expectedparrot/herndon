from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class HerndonError(Exception):
    code: str
    message: str
    details: list[dict[str, Any]] = field(default_factory=list)
    exit_code: int = 1

    def to_dict(self) -> dict[str, Any]:
        return {
            "ok": False,
            "error": {
                "code": self.code,
                "message": self.message,
                "details": self.details,
            },
        }
