from __future__ import annotations

APP_QSS = """
* {
    font-family: "Microsoft YaHei", "PingFang SC", "Inter", "Segoe UI", sans-serif;
    font-size: 13px;
}
QMainWindow, QDialog, QWidget#AppRoot {
    background: #f4f7fb;
    color: #172033;
}
QFrame#Sidebar {
    background: #111827;
    border: none;
}
QLabel#BrandTitle {
    color: #ffffff;
    font-size: 22px;
    font-weight: 800;
}
QLabel#BrandSubTitle {
    color: #9ca3af;
    font-size: 12px;
}
QLabel#PageTitle {
    color: #111827;
    font-size: 23px;
    font-weight: 800;
}
QLabel#SectionTitle {
    color: #111827;
    font-size: 16px;
    font-weight: 700;
}
QLabel#MutedLabel {
    color: #64748b;
}
QLabel#PathLabel {
    color: #334155;
    font-size: 12px;
}
QLabel#PanelTitle {
    color: #0f172a;
    font-size: 18px;
    font-weight: 800;
}
QLabel#PanelKey {
    color: #64748b;
    font-size: 12px;
    font-weight: 700;
    min-width: 34px;
}
QLabel#PanelValue {
    color: #111827;
    font-size: 13px;
    line-height: 1.35;
}
QLabel#PanelTip {
    color: #64748b;
    font-size: 12px;
}
QLabel#BigNumber {
    color: #2563eb;
    font-size: 30px;
    font-weight: 800;
}
QFrame#Card, QFrame#FloatingCard {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 16px;
}
QFrame#FloatingCard {
    background: rgba(255, 255, 255, 235);
    border-color: #dbe3ef;
}
QPushButton {
    border: none;
    border-radius: 10px;
    padding: 9px 14px;
    background: #e5e7eb;
    color: #111827;
    font-weight: 700;
}
QPushButton:hover { background: #d1d5db; }
QPushButton:disabled {
    background: #eef2f7;
    color: #94a3b8;
}
QPushButton#PrimaryButton {
    background: #2563eb;
    color: #ffffff;
}
QPushButton#PrimaryButton:hover { background: #1d4ed8; }
QPushButton#DangerButton {
    background: #fee2e2;
    color: #b91c1c;
}
QPushButton#DangerButton:hover { background: #fecaca; }
QPushButton#GhostButton {
    background: transparent;
    color: #cbd5e1;
    text-align: left;
    padding: 10px 12px;
}
QPushButton#GhostButton:hover { background: #1f2937; }
QTextEdit, QListWidget, QTableWidget, QLineEdit {
    background: #ffffff;
    border: 1px solid #dbe3ef;
    border-radius: 10px;
    padding: 8px;
    selection-background-color: #bfdbfe;
}
QTextEdit {
    background: #fbfdff;
}
QComboBox {
    background: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 9px;
    padding: 7px 34px 7px 10px;
    min-height: 22px;
    color: #111827;
}
QComboBox:hover {
    border-color: #60a5fa;
    background: #f8fbff;
}
QComboBox:focus {
    border: 1px solid #2563eb;
}
QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 30px;
    border-left: 1px solid #e2e8f0;
    border-top-right-radius: 9px;
    border-bottom-right-radius: 9px;
    background: #f8fafc;
}
QComboBox::down-arrow {
    image: none;
    width: 0;
    height: 0;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid #475569;
    margin-right: 9px;
}
QComboBox QAbstractItemView {
    background: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    padding: 4px;
    selection-background-color: #dbeafe;
    selection-color: #1d4ed8;
    outline: 0;
}
QHeaderView::section {
    background: #f8fafc;
    padding: 8px;
    border: none;
    border-bottom: 1px solid #e5e7eb;
    font-weight: 700;
}
QProgressBar {
    background: #e5e7eb;
    border-radius: 7px;
    height: 10px;
    text-align: center;
}
QProgressBar::chunk {
    background: #2563eb;
    border-radius: 7px;
}
QToolButton {
    border: 1px solid #dbe3ef;
    border-radius: 9px;
    padding: 7px 10px;
    background: #ffffff;
}
QToolButton:hover { background: #f1f5f9; }
QToolButton:checked {
    background: #dbeafe;
    border-color: #2563eb;
    color: #1d4ed8;
}
"""
