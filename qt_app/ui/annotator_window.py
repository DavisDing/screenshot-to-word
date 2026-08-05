from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

from PySide6.QtCore import QPoint, QPointF, QRect, Qt, Signal
from PySide6.QtGui import QAction, QColor, QFont, QImage, QPainter, QPen, QPixmap, QShortcut, QKeySequence
from PySide6.QtWidgets import (
    QButtonGroup,
    QDialog,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QToolButton,
    QVBoxLayout,
    QWidget,
)


@dataclass
class Annotation:
    kind: str
    data: object


class AnnotationCanvas(QWidget):
    changed = Signal()

    def __init__(self, image_path: Path, parent=None):
        super().__init__(parent)
        self.image_path = Path(image_path)
        self.pixmap = QPixmap(str(image_path))
        self.image = QImage(str(image_path))
        if self.pixmap.isNull() or self.image.isNull():
            raise ValueError(f"无法打开图片：{image_path}")
        self.setFixedSize(self.pixmap.size())
        self.setMouseTracking(True)
        self.tool = "rect"
        self.annotations: List[Annotation] = []
        self.crop_rect: Optional[QRect] = None
        self._drag_start: Optional[QPoint] = None
        self._drag_current: Optional[QPoint] = None
        self._current_path: List[QPoint] = []

    def set_tool(self, tool: str) -> None:
        self.tool = tool

    def undo(self) -> None:
        if self.tool == "crop" and self.crop_rect is not None:
            self.crop_rect = None
        elif self.annotations:
            self.annotations.pop()
        self.update()
        self.changed.emit()

    def paintEvent(self, event):  # noqa: N802 - Qt API
        painter = QPainter(self)
        painter.drawPixmap(0, 0, self.pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        for annotation in self.annotations:
            self._draw_annotation(painter, annotation)
        self._draw_active_shape(painter)
        if self.crop_rect and self.crop_rect.isValid():
            self._draw_crop_overlay(painter, self.crop_rect)

    def mousePressEvent(self, event):  # noqa: N802 - Qt API
        if event.button() != Qt.MouseButton.LeftButton:
            return
        point = event.position().toPoint()
        if self.tool == "text":
            text, ok = QInputDialog.getText(self, "添加文字", "请输入标注文字：")
            if ok and text.strip():
                self.annotations.append(Annotation("text", (point, text.strip())))
                self.update()
                self.changed.emit()
            return
        self._drag_start = point
        self._drag_current = point
        self._current_path = [point]

    def mouseMoveEvent(self, event):  # noqa: N802 - Qt API
        if not self._drag_start:
            return
        point = event.position().toPoint()
        self._drag_current = point
        if self.tool in {"pen", "highlighter"}:
            self._current_path.append(point)
        self.update()

    def mouseReleaseEvent(self, event):  # noqa: N802 - Qt API
        if event.button() != Qt.MouseButton.LeftButton or not self._drag_start:
            return
        end = event.position().toPoint()
        rect = QRect(self._drag_start, end).normalized()
        if self.tool == "crop":
            if rect.width() >= 20 and rect.height() >= 20:
                self.crop_rect = rect.intersected(self.rect())
        elif self.tool in {"rect", "ellipse"}:
            if rect.width() >= 6 and rect.height() >= 6:
                self.annotations.append(Annotation(self.tool, rect))
        elif self.tool in {"pen", "highlighter"}:
            if len(self._current_path) >= 2:
                self.annotations.append(Annotation(self.tool, list(self._current_path)))
        self._drag_start = None
        self._drag_current = None
        self._current_path = []
        self.update()
        self.changed.emit()

    def save_to(self, output_path: Optional[Path] = None) -> Path:
        output_path = Path(output_path or self.image_path)
        crop = self.crop_rect.intersected(self.image.rect()) if self.crop_rect else QRect(0, 0, self.image.width(), self.image.height())
        if not crop.isValid() or crop.width() <= 0 or crop.height() <= 0:
            crop = QRect(0, 0, self.image.width(), self.image.height())
        final_image = self.image.copy(crop)
        painter = QPainter(final_image)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.translate(-crop.x(), -crop.y())
        for annotation in self.annotations:
            if self._annotation_intersects(annotation, crop):
                self._draw_annotation(painter, annotation)
        painter.end()
        if not final_image.save(str(output_path), "PNG"):
            raise IOError(f"保存图片失败：{output_path}")
        return output_path

    def _draw_active_shape(self, painter: QPainter) -> None:
        if not self._drag_start or not self._drag_current:
            return
        if self.tool in {"rect", "crop"}:
            pen = QPen(QColor("#ef4444") if self.tool == "rect" else QColor("#2563eb"), 3, Qt.PenStyle.DashLine)
            painter.setPen(pen)
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawRect(QRect(self._drag_start, self._drag_current).normalized())
        elif self.tool == "ellipse":
            painter.setPen(QPen(QColor("#ef4444"), 3, Qt.PenStyle.DashLine))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawEllipse(QRect(self._drag_start, self._drag_current).normalized())
        elif self.tool in {"pen", "highlighter"}:
            self._draw_polyline(painter, self._current_path, highlighter=self.tool == "highlighter")

    def _draw_annotation(self, painter: QPainter, annotation: Annotation) -> None:
        if annotation.kind == "rect":
            painter.setPen(QPen(QColor("#ef4444"), 4))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawRect(annotation.data)
        elif annotation.kind == "ellipse":
            painter.setPen(QPen(QColor("#ef4444"), 4))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawEllipse(annotation.data)
        elif annotation.kind == "pen":
            self._draw_polyline(painter, annotation.data, highlighter=False)
        elif annotation.kind == "highlighter":
            self._draw_polyline(painter, annotation.data, highlighter=True)
        elif annotation.kind == "text":
            point, text = annotation.data
            painter.setPen(QPen(QColor("#2563eb"), 1))
            font = QFont()
            font.setPointSize(18)
            font.setBold(True)
            painter.setFont(font)
            painter.drawText(point, text)

    def _draw_polyline(self, painter: QPainter, points: List[QPoint], highlighter: bool) -> None:
        if len(points) < 2:
            return
        color = QColor(250, 204, 21, 105) if highlighter else QColor("#ef4444")
        width = 18 if highlighter else 4
        painter.setPen(QPen(color, width, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
        for a, b in zip(points, points[1:]):
            painter.drawLine(a, b)

    def _draw_crop_overlay(self, painter: QPainter, rect: QRect) -> None:
        overlay = QColor(15, 23, 42, 95)
        full = self.rect()
        painter.fillRect(0, 0, full.width(), rect.top(), overlay)
        painter.fillRect(0, rect.bottom(), full.width(), full.height() - rect.bottom(), overlay)
        painter.fillRect(0, rect.top(), rect.left(), rect.height(), overlay)
        painter.fillRect(rect.right(), rect.top(), full.width() - rect.right(), rect.height(), overlay)
        painter.setPen(QPen(QColor("#2563eb"), 3))
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRect(rect)

    def _annotation_intersects(self, annotation: Annotation, rect: QRect) -> bool:
        if annotation.kind in {"rect", "ellipse"}:
            return annotation.data.intersects(rect)
        if annotation.kind in {"pen", "highlighter"}:
            return any(rect.contains(point) for point in annotation.data)
        if annotation.kind == "text":
            return rect.contains(annotation.data[0])
        return True


class AnnotatorWindow(QDialog):
    def __init__(self, image_path: Path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("截图标注")
        self.resize(1200, 800)
        self.result_path: Optional[Path] = None
        self.canvas = AnnotationCanvas(image_path, self)
        self._build_ui()
        QShortcut(QKeySequence("Ctrl+Z"), self, activated=self.canvas.undo)
        QShortcut(QKeySequence("Ctrl+S"), self, activated=self._save)
        QShortcut(QKeySequence("Escape"), self, activated=self.reject)

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        toolbar = QHBoxLayout()
        toolbar.setSpacing(8)
        title = QLabel("工具：")
        toolbar.addWidget(title)

        group = QButtonGroup(self)
        group.setExclusive(True)
        tools = [
            ("rect", "方形标记"),
            ("ellipse", "圆形标记"),
            ("highlighter", "荧光笔"),
            ("pen", "画笔"),
            ("text", "文字"),
            ("crop", "裁剪"),
        ]
        for key, label in tools:
            button = QToolButton()
            button.setText(label)
            button.setCheckable(True)
            if key == "rect":
                button.setChecked(True)
            button.clicked.connect(lambda checked=False, tool=key: self.canvas.set_tool(tool))
            group.addButton(button)
            toolbar.addWidget(button)

        toolbar.addStretch(1)
        hint = QLabel("快捷键：Ctrl+Z 撤销，Ctrl+S 保存，Esc 取消。裁剪只会保存蓝色区域。")
        hint.setObjectName("MutedLabel")
        toolbar.addWidget(hint)
        undo_btn = QPushButton("撤销")
        undo_btn.clicked.connect(self.canvas.undo)
        save_btn = QPushButton("保存")
        save_btn.setObjectName("PrimaryButton")
        save_btn.clicked.connect(self._save)
        cancel_btn = QPushButton("取消")
        cancel_btn.setObjectName("DangerButton")
        cancel_btn.clicked.connect(self.reject)
        toolbar.addWidget(undo_btn)
        toolbar.addWidget(save_btn)
        toolbar.addWidget(cancel_btn)
        root.addLayout(toolbar)

        scroll = QScrollArea()
        scroll.setWidget(self.canvas)
        scroll.setWidgetResizable(False)
        root.addWidget(scroll, stretch=1)

    def _save(self) -> None:
        try:
            self.result_path = self.canvas.save_to()
        except Exception as exc:
            QMessageBox.critical(self, "保存失败", str(exc))
            return
        self.accept()
