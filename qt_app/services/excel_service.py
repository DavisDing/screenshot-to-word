from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from qt_app.models.test_case import TestItem
from utils.excel_schema import FIELD_SPECS, STEP_HINT_FIELDS, suggest_column_mapping

EXECUTED_STATUSES = {"已执行", "pass", "passed", "通过", "done", "completed", "完成", "成功"}
ROLE_CASE = "测试名称"
ROLE_CHECKPOINT = "验证点"
ROLE_STEP_NAME = "步骤名称"
ROLE_STEP_DESC = "步骤描述"
ROLE_EXPECTED = "预期结果"
ROLE_RESULT = "测试结果"
REQUIRED_ROLES = (ROLE_CASE, ROLE_CHECKPOINT, ROLE_RESULT)


@dataclass(frozen=True)
class ColumnInfo:
    index: int
    header: str
    samples: Tuple[str, ...]


@dataclass(frozen=True)
class ExcelPreview:
    file_path: Path
    sheet_name: str
    header_row: int
    columns: Tuple[ColumnInfo, ...]
    suggested_mapping: Dict[str, Optional[str]]
    reasons: Dict[str, str]

    @property
    def missing_required_roles(self) -> List[str]:
        return [role for role in REQUIRED_ROLES if not self.suggested_mapping.get(role)]


class ExcelService:
    """Excel reader/writer with tolerant column recognition and manual mapping support."""

    def __init__(self, logger):
        self.logger = logger
        self.file_path: Optional[Path] = None
        self.workbook = None
        self.sheet: Optional[Worksheet] = None
        self.header_row = 1
        self.columns: List[ColumnInfo] = []
        self.column_map: Dict[str, int] = {}
        self._merged_parent_cache: Dict[Tuple[int, int], Tuple[int, int]] = {}
        self.version = "基础版"

    def preview(self, file_path: Path) -> ExcelPreview:
        wb = load_workbook(file_path)
        ws = wb.active
        header_row = self._detect_header_row(ws)
        columns = self._read_columns(ws, header_row)
        suggestions, reasons = suggest_column_mapping([c.header for c in columns])
        wb.close()
        return ExcelPreview(
            file_path=Path(file_path),
            sheet_name=ws.title,
            header_row=header_row,
            columns=tuple(columns),
            suggested_mapping=suggestions,
            reasons=reasons,
        )

    def load(self, file_path: Path, mapping_by_role: Dict[str, Optional[str]]) -> List[TestItem]:
        self.file_path = Path(file_path)
        self.workbook = load_workbook(self.file_path)
        self.sheet = self.workbook.active
        self._merged_parent_cache = {}
        self.header_row = self._detect_header_row(self.sheet)
        self.columns = self._read_columns(self.sheet, self.header_row)
        header_to_index = {c.header: c.index for c in self.columns}
        self.column_map = {}
        for role, header in mapping_by_role.items():
            if header and header in header_to_index:
                self.column_map[role] = header_to_index[header]

        missing = [role for role in REQUIRED_ROLES if role not in self.column_map]
        if missing:
            readable = "、".join(missing)
            raise ValueError(f"Excel 缺少必要列映射：{readable}")

        self.version = "步骤版" if any(role in self.column_map for role in STEP_HINT_FIELDS) else "基础版"
        items = list(self._iter_pending_items())
        self.logger.log(f"Excel 加载完成：{self.file_path}，模式：{self.version}，待执行：{len(items)} 条")
        return items

    def mark_executed(self, row_indexes: Iterable[int], status: str = "已执行") -> None:
        if not self.workbook or not self.sheet or not self.file_path:
            raise RuntimeError("Excel 尚未加载")
        result_col = self.column_map[ROLE_RESULT]
        for row_index in row_indexes:
            self._mark_cell_value(row_index, result_col, status)
        self.workbook.save(self.file_path)
        self.logger.log(f"Excel 已标记执行结果：{self.file_path}")

    def close(self) -> None:
        if self.workbook:
            self.workbook.close()
        self.workbook = None
        self.sheet = None
        self._merged_parent_cache = {}

    def _detect_header_row(self, ws: Worksheet) -> int:
        best_row = 1
        best_score = -1
        max_probe = min(ws.max_row, 15)
        aliases = []
        for spec in FIELD_SPECS.values():
            aliases.extend([spec.label, *spec.aliases])
        normalized_aliases = {str(a).strip().lower() for a in aliases if a}

        for row in range(1, max_probe + 1):
            values = [ws.cell(row=row, column=col).value for col in range(1, ws.max_column + 1)]
            non_empty = [str(v).strip() for v in values if v not in (None, "")]
            if len(non_empty) < 2:
                continue
            score = len(non_empty)
            lowered = {v.lower() for v in non_empty}
            score += len(lowered & normalized_aliases) * 5
            if score > best_score:
                best_score = score
                best_row = row
        return best_row

    def _read_columns(self, ws: Worksheet, header_row: int) -> List[ColumnInfo]:
        columns: List[ColumnInfo] = []
        used_headers = set()
        for col in range(1, ws.max_column + 1):
            raw_header = ws.cell(row=header_row, column=col).value
            header = str(raw_header).strip() if raw_header not in (None, "") else f"第{col}列"
            if header in used_headers:
                header = f"{header}({col})"
            used_headers.add(header)
            samples = []
            for row in range(header_row + 1, min(ws.max_row, header_row + 5) + 1):
                value = ws.cell(row=row, column=col).value
                if value not in (None, ""):
                    samples.append(str(value).strip())
            columns.append(ColumnInfo(index=col, header=header, samples=tuple(samples[:3])))
        return columns

    def _get_cell_text(self, row: int, role: str, *, resolve_merged: bool = False) -> str:
        if not self.sheet or role not in self.column_map:
            return ""
        column = self.column_map[role]
        value = self.sheet.cell(row=row, column=column).value
        if value in (None, "") and resolve_merged:
            parent_row, parent_col = self._merged_parent(row, column)
            if (parent_row, parent_col) != (row, column):
                value = self.sheet.cell(row=parent_row, column=parent_col).value
        return str(value).strip() if value not in (None, "") else ""

    def _merged_parent(self, row: int, column: int) -> Tuple[int, int]:
        if not self.sheet:
            return row, column
        key = (row, column)
        if key in self._merged_parent_cache:
            return self._merged_parent_cache[key]
        parent = key
        for merged_range in self.sheet.merged_cells.ranges:
            if (merged_range.min_row <= row <= merged_range.max_row) and (merged_range.min_col <= column <= merged_range.max_col):
                parent = (merged_range.min_row, merged_range.min_col)
                break
        self._merged_parent_cache[key] = parent
        return parent

    def _mark_cell_value(self, row: int, column: int, value: str) -> None:
        if not self.sheet:
            return
        parent_row, parent_col = self._merged_parent(row, column)
        self.sheet.cell(row=parent_row, column=parent_col).value = value

    def _iter_pending_items(self):
        if not self.sheet:
            return

        current_case = ""
        current_checkpoint = ""
        for row in range(self.header_row + 1, self.sheet.max_row + 1):
            raw_case = self._get_cell_text(row, ROLE_CASE, resolve_merged=True)
            raw_checkpoint = self._get_cell_text(row, ROLE_CHECKPOINT, resolve_merged=True)

            if raw_case and raw_case != current_case:
                current_case = raw_case
                # 新用例开始但验证点为空时，不沿用上一条用例的验证点。
                current_checkpoint = raw_checkpoint or ""
            elif raw_case:
                current_case = raw_case
            if raw_checkpoint:
                current_checkpoint = raw_checkpoint

            step_name = self._get_cell_text(row, ROLE_STEP_NAME)
            step_description = self._get_cell_text(row, ROLE_STEP_DESC)
            expected_result = self._get_cell_text(row, ROLE_EXPECTED)
            case_name = current_case
            checkpoint = current_checkpoint

            if not any((case_name, checkpoint, step_name, step_description, expected_result)):
                continue

            # 测试结果列按当前行判断；测试名称/验证点可从合并单元格或上一行继承。
            result = self._get_cell_text(row, ROLE_RESULT, resolve_merged=False).lower()
            if result in EXECUTED_STATUSES:
                continue
            yield TestItem(
                row_index=row,
                case_name=case_name or f"未命名用例_{row}",
                checkpoint=checkpoint or "未填写验证点",
                step_name=step_name,
                step_description=step_description,
                expected_result=expected_result,
            )
