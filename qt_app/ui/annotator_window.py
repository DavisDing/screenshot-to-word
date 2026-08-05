from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

from PySide6.QtCore import QPoint, QRect, QSize, Qt, Signal
from PySide6.QtGui import QColor, QFont, QImage, QKeySequence, QPainter, QPen, QPixmap, QShortcut
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
    zoom_changed = Signal(float)

    def __init__(self, image_path: Path, parent=None):
        super().__init__(parent)
        self.image_path = Path(image_path)
        self.pixmap = QPixmap(str(image_path))
        self.image = QImage(str(image_path))
        if self.pixmap.isNull() or self.image.isNull():
            raise ValueError(f"无法打开图片：{image_path}")
        self.scale_factor = 1.0
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

    def set_scale(self, scale: float) -> None:
        self.scale_factor = max(0.12, min(scale, 3.0))
        new_size = QSize(
            max(1, round(self.pixmap.width() * self.scale_factor)),
            max(1, round(self.pixmap.height() * self.scale_factor)),
        )
        self.setFixedSize(new_size)
        self.update()
        self.zoom_changed.emit(self.scale_factor)

    def fit_to(self, size: QSize) -> None:
        if self.pixmap.width() <= 0 or self.pixmap.height() <= 0:
            return
        scale = min(
            max(0.12, (size.width() - 24) / self.pixmap.width()),
            max(0.12, (size.height() - 24) / self.pixmap.height()),
            1.0,
        )
        self.set_scale(scale)

    def undo(self) -> None:
        if self.tool == "crop" and self.crop_rect is not None:
            self.crop_rect = None
        elif self.annotations:
            self.annotations.pop()
        self.update()
        self.changed.emit()

    def paintEvent(self, event):  # noqa: N802 - Qt API
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)
        painter.scale(self.scale_factor, self.scale_factor)
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
        point = self._to_image_point(event.position().toPoint())
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
        point = self._to_image_point(event.position().toPoint())
        self._drag_current = point
        if self.tool in {"pen", "highlighter"}:
            self._current_path.append(point)
        self.update()

    def mouseReleaseEvent(self, event):  # noqa: N802 - Qt API
        if event.button() != Qt.MouseButton.LeftButton or not self._drag_start:
            return
        end = self._to_image_point(event.position().toPoint())
        rect = QRect(self._drag_start, end).normalized()
        image_rect = QRect(0, 0, self.image.width(), self.image.height())
        if self.tool == "crop":
            if rect.width() >= 20 and rect.height() >= 20:
                self.crop_rect = rect.intersected(image_rect)
        elif self.tool in {"rect", "ellipse"}:
            if rect.width() >= 6 and rect.height() >= 6:
                self.annotations.append(Annotation(self.tool, rect.intersected(image_rect)))
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

    def _to_image_point(self, display_point: QPoint) -> QPoint:
        x = round(display_point.x() / self.scale_factor)
        y = round(display_point.y() / self.scale_factor)
        return QPoint(
            max(0, min(x, self.image.width() - 1)),
            max(0, min(y, self.image.height() - 1)),
        )

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
        full = QRect(0, 0, self.image.width(), self.image.height())
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


from qt_app.ui.icons import app_icon


class AnnotatorWindow(QDialog):
    def __init__(self, image_path: Path, parent=None):
        super().__init__(parent)
        icon = app_icon()
        if not icon.isNull():
            self.setWindowIcon(icon)
        self.setWindowTitle("截图标注")
        self.resize(1100, 760)
        self.result_path: Optional[Path] = None
        self.canvas = AnnotationCanvas(image_path, self)
        self.auto_fit = True
        self._build_ui()
        QShortcut(QKeySequence("Ctrl+Z"), self, activated=self.canvas.undo)
        QShortcut(QKeySequence("Ctrl+S"), self, activated=self._save)
        QShortcut(QKeySequence("Escape"), self, activated=self.reject)

    def showEvent(self, event):  # noqa: N802 - Qt API
        super().showEvent(event)
        if self.auto_fit:
            self.fit_to_window()

    def resizeEvent(self, event):  # noqa: N802 - Qt API
        super().resizeEvent(event)
        if self.auto_fit and hasattr(self, "scroll"):
            self.fit_to_window()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)
        toolbar = QHBoxLayout()
        toolbar.setSpacing(6)
        title = QLabel("工具：")
        toolbar.addWidget(title)

        group = QButtonGroup(self)
        group.setExclusive(True)
        tools = [
            ("rect", "方形"),
            ("ellipse", "圆形"),
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

        toolbar.addSpacing(8)
        fit_btn = QPushButton("适配")
        fit_btn.clicked.connect(self.fit_to_window)
        actual_btn = QPushButton("100%")
        actual_btn.clicked.connect(lambda: self.set_zoom(1.0, auto=False))
        zoom_out = QPushButton("-")
        zoom_out.clicked.connect(lambda: self.set_zoom(self.canvas.scale_factor * 0.85, auto=False))
        zoom_in = QPushButton("+")
        zoom_in.clicked.connect(lambda: self.set_zoom(self.canvas.scale_factor * 1.15, auto=False))
        self.zoom_label = QLabel("100%")
        self.zoom_label.setObjectName("MutedLabel")
        self.canvas.zoom_changed.connect(lambda scale: self.zoom_label.setText(f"{round(scale * 100)}%"))
        toolbar.addWidget(fit_btn)
        toolbar.addWidget(actual_btn)
        toolbar.addWidget(zoom_out)
        toolbar.addWidget(zoom_in)
        toolbar.addWidget(self.zoom_label)

        toolbar.addStretch(1)
        hint = QLabel("Ctrl+Z 撤销，Ctrl+S 保存，Esc 取消")
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

        self.scroll = QScrollArea()
        self.scroll.setWidget(self.canvas)
        self.scroll.setWidgetResizable(False)
        self.scroll.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root.addWidget(self.scroll, stretch=1)

    def fit_to_window(self) -> None:
        self.auto_fit = True
        self.canvas.fit_to(self.scroll.viewport().size())

    def set_zoom(self, scale: float, auto: bool = False) -> None:
        self.auto_fit = auto
        self.canvas.set_scale(scale)

    def _save(self) -> None:
        try:
            self.result_path = self.canvas.save_to()
        except Exception as exc:
            QMessageBox.critical(self, "保存失败", str(exc))
            return
        self.accept()
