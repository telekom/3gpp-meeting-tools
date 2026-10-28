"""Data model representing a reusable PlantUML syntax element or pattern."""

from dataclasses import dataclass
from typing import Tuple


@dataclass(frozen=True)
class PlantUMLSnippet:
    """Represents a reusable PlantUML syntax element or pattern."""
    id: str
    label: str
    category: str
    diagram_scope: Tuple[str, ...]
    description: str
    code: str
    is_block_wrapper: bool = False
    is_user_defined: bool = False

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "label": self.label,
            "category": self.category,
            "diagram_scope": list(self.diagram_scope),
            "description": self.description,
            "code": self.code,
            "is_block_wrapper": self.is_block_wrapper,
            "is_user_defined": self.is_user_defined,
        }

    @classmethod
    def from_dict(cls, data: dict, is_user_defined: bool = False) -> "PlantUMLSnippet":
        scope_raw = data.get("diagram_scope", ["All"])
        if isinstance(scope_raw, list):
            scope = tuple(str(s) for s in scope_raw)
        elif isinstance(scope_raw, str):
            scope = (scope_raw,)
        else:
            scope = tuple(scope_raw)

        return cls(
            id=str(data.get("id", "")),
            label=str(data.get("label", "")),
            category=str(data.get("category", "General")),
            diagram_scope=scope,
            description=str(data.get("description", "")),
            code=str(data.get("code", "")),
            is_block_wrapper=bool(data.get("is_block_wrapper", False)),
            is_user_defined=bool(data.get("is_user_defined", is_user_defined)),
        )
