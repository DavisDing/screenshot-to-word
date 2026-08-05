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
    font-size: 24px;
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
QLabel#BigNumber {
    color: #2563eb;
    font-size: 30px;
    font-weight: 800;
}
QFrame#Card {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 16px;
}
QPushButton {
    border: none;
    border-radius: 10px;
    padding: 10px 16px;
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
}
QPushButton#GhostButton:hover { background: #1f2937; }
QTextEdit, QListWidget, QTableWidget, QComboBox, QLineEdit {
    background: #ffffff;
    border: 1px solid #dbe3ef;
    border-radius: 10px;
    padding: 8px;
    selection-background-color: #bfdbfe;
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
    border-radius: 8px;
    height: 12px;
    text-align: center;
}
QProgressBar::chunk {
    background: #2563eb;
    border-radius: 8px;
}
QToolButton {
    border: 1px solid #dbe3ef;
    border-radius: 9px;
    padding: 8px 12px;
    background: #ffffff;
}
QToolButton:checked {
    background: #dbeafe;
    border-color: #2563eb;
    color: #1d4ed8;
}
"""
