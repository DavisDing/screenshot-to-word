from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Dict, Iterable, Optional

from utils.excel_schema import FIELD_SPECS, FieldSpec

NO_MAPPING_OPTION = "<不映射>"


class ExcelMappingDialog(tk.Toplevel):
    """Modal dialog for mapping arbitrary Excel headers to canonical fields."""

    def __init__(
        self,
        parent: tk.Misc,
        columns: Iterable[str],
        suggestions: Dict[str, Optional[str]],
        reasons: Dict[str, str],
        sample_rows,
        required_fields: Iterable[str],
        version_hint: str,
    ):
        super().__init__(parent)
        self.parent = parent
        self.columns = [str(col) for col in columns]
        self.suggestions = suggestions
        self.reasons = reasons
        self.sample_rows = sample_rows
        self.required_fields = set(required_fields)
        self.version_hint = version_hint
        self.result_map: Optional[Dict[str, Optional[str]]] = None
        self.vars: Dict[str, tk.StringVar] = {}

        self.title("Excel 列映射确认")
        self.geometry("980x680")
        self.minsize(820, 560)
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

        try:
            self.configure(background="#f5f7fb")
        except Exception:
            pass

        self._build_ui()
        self.after(80, self.focus_force)

    def _build_ui(self) -> None:
        container = ttk.Frame(self, padding=18, style="App.TFrame")
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=0)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(1, weight=1)

        title = ttk.Label(container, text="需要确认 Excel 列含义", style="Header.TLabel")
        title.grid(row=0, column=0, columnspan=2, sticky="w")
        subtitle = ttk.Label(
            container,
            text=(
                f"当前推断为 {self.version_hint}。系统未能完整自动识别必需列，请根据右侧表头和数据预览，"
                "把业务字段映射到实际 Excel 列。带 * 的字段必须选择。"
            ),
            style="SubHeader.TLabel",
            wraplength=920,
            justify="left",
        )
        subtitle.grid(row=0, column=0, columnspan=2, sticky="w", pady=(30, 12))

        mapping_frame = ttk.Labelframe(container, text="业务字段映射", padding=14)
        mapping_frame.grid(row=1, column=0, sticky="nsw", padx=(0, 14))
        mapping_frame.columnconfigure(1, weight=1)

        self._build_mapping_rows(mapping_frame)

        preview_frame = ttk.Labelframe(container, text="Excel 表头与数据预览", padding=10)
        preview_frame.grid(row=1, column=1, sticky="nsew")
        preview_frame.rowconfigure(1, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        self._build_preview(preview_frame)

        footer = ttk.Frame(container, style="App.TFrame")
        footer.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        footer.columnconfigure(0, weight=1)

        self.warning_var = tk.StringVar(value="")
        warning = ttk.Label(footer, textvariable=self.warning_var, foreground="#dc2626", style="SubHeader.TLabel")
        warning.grid(row=0, column=0, sticky="w")

        ttk.Button(footer, text="取消", command=self._on_cancel, style="Secondary.TButton").grid(row=0, column=1, padx=(8, 0))
        ttk.Button(footer, text="确认映射", command=self._on_confirm, style="Primary.TButton").grid(row=0, column=2, padx=(8, 0))

    def _build_mapping_rows(self, parent: ttk.Frame) -> None:
        options = [NO_MAPPING_OPTION, *self.columns]

        header_font = None
        for col, text in enumerate(("业务字段", "Excel 列", "识别说明")):
            lbl = ttk.Label(parent, text=text, style="Section.TLabel")
            lbl.grid(row=0, column=col, sticky="w", padx=(0, 10), pady=(0, 8))
            header_font = lbl.cget("font")

        for row_index, (field, spec) in enumerate(FIELD_SPECS.items(), start=1):
            required_mark = " *" if field in self.required_fields else ""
            field_label = f"{spec.label}{required_mark}"
            ttk.Label(parent, text=field_label, style="Body.TLabel").grid(row=row_index, column=0, sticky="nw", padx=(0, 10), pady=6)

            var = tk.StringVar(value=self.suggestions.get(field) or NO_MAPPING_OPTION)
            self.vars[field] = var
            combo = ttk.Combobox(parent, textvariable=var, values=options, state="readonly", width=28)
            combo.grid(row=row_index, column=1, sticky="ew", padx=(0, 10), pady=6)

            hint = self._build_hint_text(field, spec)
            ttk.Label(parent, text=hint, style="Hint.TLabel", wraplength=240, justify="left").grid(
                row=row_index, column=2, sticky="w", pady=6
            )

    def _build_hint_text(self, field: str, spec: FieldSpec) -> str:
        reason = self.reasons.get(field, "未识别，请手动选择")
        required = "必填" if field in self.required_fields else "可选"
        return f"{reason}\n{required}；常见列名：{spec.help_text}"

    def _build_preview(self, parent: ttk.Frame) -> None:
        column_hint = "；".join(f"第 {idx} 列：{name}" for idx, name in enumerate(self.columns, start=1))
        ttk.Label(parent, text=column_hint, style="Hint.TLabel", wraplength=560, justify="left").grid(row=0, column=0, sticky="ew", pady=(0, 8))

        tree_container = ttk.Frame(parent, style="Card.TFrame")
        tree_container.grid(row=1, column=0, sticky="nsew")
        tree_container.rowconfigure(0, weight=1)
        tree_container.columnconfigure(0, weight=1)

        tree = ttk.Treeview(tree_container, columns=self.columns, show="headings", height=10)
        tree.grid(row=0, column=0, sticky="nsew")

        y_scroll = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(tree_container, orient="horizontal", command=tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        for col in self.columns:
            tree.heading(col, text=str(col))
            tree.column(col, width=max(120, min(240, len(str(col)) * 16)), anchor="w", stretch=False)

        rows = self._sample_rows_as_values()
        if not rows:
            tree.insert("", "end", values=["" for _ in self.columns])
        for row in rows:
            tree.insert("", "end", values=row)

    def _sample_rows_as_values(self):
        rows = []
        try:
            records = self.sample_rows.to_dict(orient="records")
        except Exception:
            records = []

        for record in records:
            rows.append([self._safe_text(record.get(col, "")) for col in self.columns])
        return rows

    @staticmethod
    def _safe_text(value) -> str:
        if value is None:
            return ""
        text = str(value)
        if text.lower() == "nan":
            return ""
        return text

    def _on_confirm(self) -> None:
        result: Dict[str, Optional[str]] = {}
        missing_required = []

        for field in FIELD_SPECS:
            selected = self.vars[field].get()
            selected = None if selected == NO_MAPPING_OPTION else selected
            result[field] = selected
            if field in self.required_fields and not selected:
                missing_required.append(field)

        if missing_required:
            self.warning_var.set("请先选择必填字段：" + "、".join(missing_required))
            return

        duplicates = self._find_duplicate_required_mapping(result)
        if duplicates:
            if not messagebox.askyesno(
                "确认复用列",
                "以下 Excel 列被多个业务字段复用：\n"
                + "\n".join(duplicates)
                + "\n\n通常不建议复用。仍然继续吗？",
                parent=self,
            ):
                return

        self.result_map = result
        self.destroy()

    def _find_duplicate_required_mapping(self, result: Dict[str, Optional[str]]):
        reverse: Dict[str, list[str]] = {}
        for field, column in result.items():
            if not column:
                continue
            reverse.setdefault(column, []).append(field)
        return [f"{column}：{', '.join(fields)}" for column, fields in reverse.items() if len(fields) > 1]

    def _on_cancel(self) -> None:
        self.result_map = None
        self.destroy()
