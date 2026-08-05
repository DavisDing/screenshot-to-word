from __future__ import annotations

import os
from typing import Dict, Optional

import pandas as pd
from openpyxl import load_workbook
from tkinter import filedialog, messagebox

from ui.excel_mapping_dialog import ExcelMappingDialog
from utils.excel_schema import (
    FIELD_SPECS,
    STEP_HINT_FIELDS,
    fields_requiring_manual_mapping,
    get_required_fields,
    normalize_column_name,
    suggest_column_mapping,
)


EXECUTED_STATUSES = {"已执行", "pass", "passed", "通过", "done", "completed"}


class ExcelHandler:
    def __init__(self, logger, input_dir, root):
        self.logger = logger
        self.input_dir = input_dir
        self.file_path = None
        self.df: Optional[pd.DataFrame] = None
        self.root = root
        self.version = "基础版"
        self.result_col_index: Optional[int] = None
        self.column_map: Dict[str, Optional[str]] = {field: None for field in FIELD_SPECS}
        self.column_reasons: Dict[str, str] = {}
        self.actual_columns = []

    def _show_error(self, msg: str):
        self.root.attributes('-topmost', True)
        messagebox.showerror("错误", msg, parent=self.root)
        self.root.attributes('-topmost', False)

    def _show_info(self, msg: str):
        self.root.attributes('-topmost', True)
        messagebox.showinfo("提示", msg, parent=self.root)
        self.root.attributes('-topmost', False)

    def select_excel_file(self) -> bool:
        try:
            files = [
                f for f in os.listdir(self.input_dir)
                if f.lower().endswith('.xlsx') and not f.startswith('~$')
            ]
            if not files:
                self._show_error("excel_input 目录下未找到 Excel 文件")
                self.logger.log("excel_input目录无Excel文件")
                return False

            if len(files) == 1:
                self.file_path = os.path.join(self.input_dir, files[0])
                self.logger.log(f"自动选中Excel文件：{files[0]}")
            else:
                self.root.attributes('-topmost', True)
                self.file_path = filedialog.askopenfilename(
                    title="请选择要执行的 Excel 文件",
                    initialdir=self.input_dir,
                    filetypes=[("Excel 文件", "*.xlsx")],
                    parent=self.root
                )
                self.root.attributes('-topmost', False)
                if self.file_path:
                    self.logger.log(f"用户选择Excel文件：{os.path.basename(self.file_path)}")
                else:
                    self.logger.log("用户取消选择Excel文件")
                    return False
            return True
        except Exception as e:
            self._show_error(f"选择Excel文件异常：{e}")
            self.logger.log(f"选择Excel文件异常：{e}")
            return False

    def load_excel(self) -> bool:
        if not self.file_path:
            self._show_error("未选择 Excel 文件")
            return False

        self.result_col_index = None
        self.column_map = {field: None for field in FIELD_SPECS}
        self.column_reasons = {}

        try:
            self.df = pd.read_excel(self.file_path, header=0)
            self.df.columns = [str(col).strip() for col in self.df.columns]
            self.actual_columns = list(self.df.columns)
            self.logger.log(f"成功加载Excel文件：{os.path.basename(self.file_path)}")
        except Exception as e:
            self._show_error(f"读取Excel失败：{e}")
            self.logger.log(f"读取Excel失败: {e}")
            return False

        if self.df.shape[1] < 3:
            self._show_error("Excel文件列数不足，至少需要3列（测试名称、验证点、测试结果）")
            self.logger.log("Excel列数不足3列")
            return False

        self.column_map, self.column_reasons = suggest_column_mapping(self.actual_columns)
        self.version = self._infer_version()
        required_fields = get_required_fields(self.version)
        missing_required = fields_requiring_manual_mapping(self.column_map, required_fields)
        need_format_confirmation = self._needs_format_confirmation(missing_required)

        if missing_required or need_format_confirmation:
            self.logger.log(
                "Excel列名需要人工确认："
                + ("缺失字段=" + ", ".join(missing_required) if missing_required else "列数较多，确认是否为步骤版")
            )
            if not self._request_manual_mapping(required_fields):
                self.logger.log("用户取消Excel列映射，加载终止")
                return False
            self.version = self._infer_version()
            required_fields = get_required_fields(self.version)
            missing_required = fields_requiring_manual_mapping(self.column_map, required_fields)
            if missing_required:
                self._show_error("Excel列映射不完整，缺少必填字段：" + "、".join(missing_required))
                self.logger.log("Excel列映射不完整：" + ", ".join(missing_required))
                return False

        self.result_col_index = self._find_excel_column_index(self.column_map.get("测试结果"))
        if self.result_col_index is None:
            self._show_error("Excel中未能确定“测试结果”列，请在映射窗口中选择正确列。")
            self.logger.log("未能确定测试结果列索引")
            return False

        self.logger.log(f"检测到{self.version}Excel格式")
        self.logger.log("列映射结果：" + self._format_column_map())
        self.logger.log(f"检测到“测试结果”列在第 {self.result_col_index} 列")
        return True

    def _needs_format_confirmation(self, missing_required) -> bool:
        """Ask for manual confirmation when format inference is ambiguous or step columns are partial."""
        if missing_required:
            return False
        if self.df is None:
            return False
        has_step_hint = any(self.column_map.get(field) for field in STEP_HINT_FIELDS)
        has_all_step_hints = all(self.column_map.get(field) for field in STEP_HINT_FIELDS)
        if self.df.shape[1] >= 5 and self.version == "步骤版" and not has_all_step_hints:
            return True
        return self.df.shape[1] >= 5 and not has_step_hint

    def _request_manual_mapping(self, required_fields) -> bool:
        self.root.attributes('-topmost', True)
        try:
            dialog = ExcelMappingDialog(
                self.root,
                columns=self.actual_columns,
                suggestions=self.column_map,
                reasons=self.column_reasons,
                sample_rows=self.df.head(5) if self.df is not None else [],
                required_fields=required_fields,
                version_hint=self.version,
            )
            self.root.wait_window(dialog)
            if dialog.result_map is None:
                return False
            self.column_map = dialog.result_map
            self.column_reasons = {
                field: ("用户手动选择：" + value if value else "未映射")
                for field, value in self.column_map.items()
            }
            return True
        finally:
            self.root.attributes('-topmost', False)

    def _infer_version(self) -> str:
        if any(self.column_map.get(field) for field in STEP_HINT_FIELDS):
            return "步骤版"
        return "基础版"

    def _find_excel_column_index(self, actual_column: Optional[str]) -> Optional[int]:
        if not actual_column:
            return None
        target_norm = normalize_column_name(actual_column)
        try:
            wb = load_workbook(self.file_path, read_only=True)
            ws = wb.active
            for i, cell in enumerate(ws[1], start=1):
                header = "" if cell.value is None else str(cell.value).strip()
                if header == actual_column or normalize_column_name(header) == target_norm:
                    wb.close()
                    return i
            wb.close()
        except Exception as e:
            self.logger.log(f"解析Excel列头失败：{e}")
        return None

    def _format_column_map(self) -> str:
        items = []
        for field in FIELD_SPECS:
            actual = self.column_map.get(field) or "未映射"
            items.append(f"{field}->{actual}")
        return "；".join(items)

    def _get_value(self, row, field: str, default=""):
        actual_column = self.column_map.get(field)
        if not actual_column:
            return default
        try:
            value = row.get(actual_column, default)
        except AttributeError:
            value = default
        if pd.isna(value):
            return default
        return value

    def _get_text(self, row, field: str, default="") -> str:
        value = self._get_value(row, field, default)
        if pd.isna(value):
            return default
        return str(value).strip()

    @staticmethod
    def _is_executed_status(status: str) -> bool:
        return status.strip().lower() in EXECUTED_STATUSES or status.strip() in EXECUTED_STATUSES

    def get_pending_cases(self):
        if self.df is None:
            return
        self.logger.log("开始获取未执行用例")

        if self.version == "步骤版":
            from collections import defaultdict
            case_group = defaultdict(list)

            last_test_name = ""
            last_checkpoint = ""

            for idx, row in self.df.iterrows():
                test_name = self._get_text(row, "测试名称") or last_test_name
                checkpoint = self._get_text(row, "验证点") or last_checkpoint
                last_test_name, last_checkpoint = test_name, checkpoint

                key = (test_name, checkpoint)
                status = self._get_text(row, "测试结果").lower()
                case_group[key].append((idx, status))

            for (test_name, check_point), items in case_group.items():
                all_status = [s for _, s in items]
                if all(self._is_executed_status(s) for s in all_status):
                    self.logger.log(f"用例已完成: {test_name}-{check_point}")
                    continue
                first_index = items[0][0]
                self.logger.log(f"待执行用例: 行={first_index}, 用例名={test_name}, 验证点={check_point}")
                yield first_index, test_name, check_point
            return

        for idx, row in self.df.iterrows():
            status = self._get_text(row, "测试结果").lower()
            if not self._is_executed_status(status):
                filename = self._get_text(row, "测试名称")
                checkpoint = self._get_text(row, "验证点")
                self.logger.log(f"待执行用例: 行={idx}, 用例名={filename}, 验证点={checkpoint}")
                yield idx, filename, checkpoint

    def mark_case_executed(self, index: int):
        if self.df is None:
            return
        result_column = self.column_map.get("测试结果")
        if result_column:
            self.df.at[index, result_column] = "已执行"
        try:
            if self.result_col_index is None:
                self._show_error("内部错误：测试结果列索引未初始化")
                self.logger.log("错误：result_col_index 为 None，可能未正确加载 Excel")
                return
            excel_row = index + 2  # DataFrame 第一行为表头
            wb = load_workbook(self.file_path)
            ws = wb.active
            ws.cell(row=excel_row, column=self.result_col_index).value = "已执行"
            wb.save(self.file_path)
            wb.close()
            self.logger.log(f"更新用例行 {excel_row} 状态为 已执行")
        except PermissionError:
            self._show_error("写入Excel失败，文件被占用或无权限。请关闭Excel后重试。")
            self.logger.log("写入Excel失败，权限不足或文件被占用。")
        except Exception as e:
            self._show_error(f"写入Excel失败：{e}")
            self.logger.log(f"写入Excel失败：{e}")

    def get_step_cases(self):
        if self.df is None:
            return {}
        from collections import defaultdict
        case_dict = defaultdict(list)
        current_test = ""
        current_check = ""
        for idx, row in self.df.iterrows():
            row_test = self._get_text(row, "测试名称")
            row_check = self._get_text(row, "验证点")
            if row_test:
                current_test = row_test
            if row_check:
                current_check = row_check
            case_dict[(current_test, current_check)].append({
                "index": idx,
                "步骤名称": self._get_text(row, "步骤名称"),
                "步骤描述": self._get_text(row, "步骤描述"),
                "预期结果": self._get_text(row, "预期结果")
            })
        return case_dict
