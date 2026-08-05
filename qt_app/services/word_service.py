from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.shared import Inches

INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]+')


class WordService:
    def __init__(self, output_dir: Path, logger):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.logger = logger

    def add_image(self, case_name: str, checkpoint: str, image_path: Path, step_note: str = "") -> Path:
        doc_path = self.output_dir / f"{self._safe_filename(case_name)}.docx"
        if doc_path.exists():
            doc = Document(str(doc_path))
            self.logger.log(f"打开已有 Word 报告：{doc_path}")
        else:
            doc = Document()
            doc.add_heading(case_name, level=1)
            doc.add_paragraph(f"验证点：{checkpoint}")
            self.logger.log(f"创建 Word 报告：{doc_path}")

        if checkpoint:
            doc.add_paragraph(f"验证点：{checkpoint}")
        if step_note:
            for line in step_note.splitlines():
                if line.strip():
                    doc.add_paragraph(line.strip())
        doc.add_picture(str(image_path), width=Inches(5.8))
        doc.add_paragraph()
        doc.save(str(doc_path))
        self.logger.log(f"截图已写入 Word：{doc_path}")
        return doc_path

    def _safe_filename(self, name: str) -> str:
        cleaned = INVALID_FILENAME_CHARS.sub("_", name).strip(" ._")
        return cleaned[:80] or "未命名用例"
