from __future__ import annotations

import hashlib
import re
import unicodedata
from pathlib import Path

INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]+')


class WordService:
    def __init__(self, output_dir: Path, logger):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.logger = logger

    def add_image(self, case_name: str, checkpoint: str, image_path: Path, step_note: str = "") -> Path:
        # Lazy import: python-docx is only needed when the user actually finishes a screenshot.
        from docx import Document
        from docx.shared import Inches

        doc_path, doc = self._resolve_report(case_name)
        if doc is not None:
            self.logger.log(f"打开已有 Word 报告：{doc_path}")
        else:
            doc = Document()
            doc.add_heading(case_name, level=1)
            self.logger.log(f"创建 Word 报告：{doc_path}")
        doc.core_properties.identifier = self._case_identifier(case_name)

        if checkpoint and not self._has_checkpoint(doc, checkpoint):
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

    def _resolve_report(self, case_name: str):
        from docx import Document

        base = self._safe_filename(case_name)
        identity = self._case_identifier(case_name)
        digest = identity.rsplit(":", 1)[1][:12]

        def candidate(index: int) -> str:
            if index == 0:
                return f"{base}.docx"
            suffix = f"-{digest}" + (f"-{index}" if index > 1 else "")
            return f"{base[:80 - len(suffix)].rstrip(' ._')}{suffix}.docx"

        # Treat case/Unicode-equivalent names as occupied on every platform. This
        # also protects reports copied between case-sensitive and macOS/Windows disks.
        existing = {}
        for path in self.output_dir.iterdir():
            existing.setdefault(self._filename_key(path.name), []).append(path)

        indices = {0, 1}
        suffix_pattern = re.compile(rf"-{digest}(?:-(\d+))?\.docx$", re.IGNORECASE)
        for filename in existing:
            match = suffix_pattern.search(filename)
            if match:
                index = int(match.group(1) or "1")
                if index > 0 and filename == self._filename_key(candidate(index)):
                    indices.add(index)

        # Find our report before choosing a free name, including after another
        # colliding report was removed. The heading is the identity used by old reports.
        for index in sorted(indices):
            paths = existing.get(self._filename_key(candidate(index)), [])
            if len(paths) != 1 or not paths[0].is_file():
                continue
            try:
                doc = Document(str(paths[0]))
            except Exception:
                # An unreadable/foreign file must never be overwritten or appended to.
                continue
            stored_identity = doc.core_properties.identifier
            if stored_identity:
                matches_case = stored_identity == identity
            else:
                heading = next(
                    (p.text for p in doc.paragraphs if p.style and p.style.style_id == "Heading1"),
                    None,
                )
                matches_case = heading == case_name
            if matches_case:
                return paths[0], doc

        index = 0
        while self._filename_key(candidate(index)) in existing:
            index += 1
        return self.output_dir / candidate(index), None

    @staticmethod
    def _filename_key(name: str) -> str:
        return unicodedata.normalize("NFC", name).casefold()

    @staticmethod
    def _case_identifier(case_name: str) -> str:
        digest = hashlib.sha256(case_name.encode("utf-8")).hexdigest()
        return f"screenshot-to-word:case-sha256:{digest}"

    def _has_checkpoint(self, doc, checkpoint: str) -> bool:
        expected = f"验证点：{checkpoint}".strip()
        return any(paragraph.text.strip() == expected for paragraph in doc.paragraphs)

    def _safe_filename(self, name: str) -> str:
        cleaned = INVALID_FILENAME_CHARS.sub("_", name).strip(" ._")
        return cleaned[:80].rstrip(" ._") or "未命名用例"
