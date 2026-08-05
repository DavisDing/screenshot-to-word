from __future__ import annotations

import re
from collections import OrderedDict
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple


@dataclass(frozen=True)
class FieldSpec:
    key: str
    label: str
    help_text: str
    required: bool
    aliases: Tuple[str, ...]


FIELD_SPECS: "OrderedDict[str, FieldSpec]" = OrderedDict([
    (
        "测试名称",
        FieldSpec(
            key="测试名称",
            label="测试名称",
            help_text="用例标题、用例名称、Case Name、Test Name",
            required=True,
            aliases=("用例名", "用例名称", "测试用例", "case name", "test name", "name"),
        ),
    ),
    (
        "验证点",
        FieldSpec(
            key="验证点",
            label="验证点",
            help_text="检查点、校验点、检查项、Checkpoint",
            required=True,
            aliases=("检查点", "校验点", "检查项", "checkpoint", "check point"),
        ),
    ),
    (
        "步骤名称",
        FieldSpec(
            key="步骤名称",
            label="步骤名称",
            help_text="步骤标题、操作步骤、Step Name",
            required=False,
            aliases=("步骤", "操作步骤", "step name", "step", "步骤标题"),
        ),
    ),
    (
        "步骤描述",
        FieldSpec(
            key="步骤描述",
            label="步骤描述",
            help_text="步骤说明、操作说明、描述、Step Description",
            required=False,
            aliases=("步骤说明", "操作说明", "描述", "step description", "desc"),
        ),
    ),
    (
        "预期结果",
        FieldSpec(
            key="预期结果",
            label="预期结果",
            help_text="期望结果、Expected Result、Expected",
            required=False,
            aliases=("期望结果", "expected result", "expected", "期望"),
        ),
    ),
    (
        "测试结果",
        FieldSpec(
            key="测试结果",
            label="测试结果",
            help_text="结果、状态、执行结果、Result、Status",
            required=True,
            aliases=("结果", "状态", "执行结果", "result", "status", "pass/fail"),
        ),
    ),
])

BASIC_REQUIRED_FIELDS = ("测试名称", "验证点", "测试结果")
STEP_REQUIRED_FIELDS = ("测试名称", "验证点", "测试结果")
STEP_HINT_FIELDS = ("步骤名称", "步骤描述", "预期结果")


_NORMALIZE_PATTERN = re.compile(r"[\s\-_·./\\|:：,，。；;()（）\[\]{}<>《》\"'`~!@#$%^&*=+?]+")


def normalize_column_name(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return _NORMALIZE_PATTERN.sub("", text)


def _normalized_aliases(spec: FieldSpec) -> List[str]:
    return [normalize_column_name(spec.label), *[normalize_column_name(alias) for alias in spec.aliases]]


def suggest_column_mapping(columns: Iterable[str]) -> Tuple[Dict[str, Optional[str]], Dict[str, str]]:
    column_list = list(columns)
    normalized_columns = {normalize_column_name(col): col for col in column_list}
    suggestions: Dict[str, Optional[str]] = {key: None for key in FIELD_SPECS}
    reasons: Dict[str, str] = {}

    for key, spec in FIELD_SPECS.items():
        matched: Optional[str] = None
        reason = "未识别，请手动选择"
        aliases = _normalized_aliases(spec)

        for alias in aliases:
            if alias and alias in normalized_columns:
                matched = normalized_columns[alias]
                reason = f"自动识别为：{matched}"
                break

        if matched is None:
            for col in column_list:
                normalized_col = normalize_column_name(col)
                if not normalized_col:
                    continue
                if any(alias and (alias in normalized_col or normalized_col in alias) for alias in aliases):
                    matched = col
                    reason = f"智能匹配为：{matched}"
                    break

        suggestions[key] = matched
        reasons[key] = reason

    return suggestions, reasons


def format_column_preview(columns: Iterable[str]) -> str:
    lines = []
    for index, column in enumerate(columns, start=1):
        lines.append(f"第 {index} 列：{column}")
    return "\n".join(lines)


def fields_requiring_manual_mapping(suggestion_map: Dict[str, Optional[str]], required_fields: Iterable[str]) -> List[str]:
    missing = []
    for field in required_fields:
        if not suggestion_map.get(field):
            missing.append(field)
    return missing


def get_required_fields(version: str) -> Tuple[str, ...]:
    if version == "步骤版":
        return STEP_REQUIRED_FIELDS
    return BASIC_REQUIRED_FIELDS
