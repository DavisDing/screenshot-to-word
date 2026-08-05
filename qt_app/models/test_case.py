from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class TestItem:
    """One executable Excel row/test step."""

    row_index: int
    case_name: str
    checkpoint: str
    step_name: str = ""
    step_description: str = ""
    expected_result: str = ""

    @property
    def is_step(self) -> bool:
        return bool(self.step_name or self.step_description or self.expected_result)

    @property
    def display_title(self) -> str:
        if self.is_step and self.step_name:
            return f"{self.case_name} / {self.step_name}"
        return self.case_name

    def step_note(self) -> str:
        parts = []
        if self.step_name:
            parts.append(f"步骤名称：{self.step_name}")
        if self.step_description:
            parts.append(f"步骤描述：{self.step_description}")
        if self.expected_result:
            parts.append(f"预期结果：{self.expected_result}")
        return "\n".join(parts)
