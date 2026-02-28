# -*- coding: utf-8 -*-
from __future__ import annotations

from PySide6.QtCore import Qt, QEvent, QObject
from PySide6.QtWidgets import (
    QApplication, QWidget, QFrame, QVBoxLayout, QLabel, QPushButton,
    QProgressBar, QLineEdit
)

class NiceProgressBar(QProgressBar):
    """A pretty progress bar (keeps QProgressBar API)."""
    def __init__(self):
        super().__init__()
        self.setObjectName("NiceProgress")
        self.setRange(0, 100)
        self.setValue(0)
        self.setFixedHeight(14)
        self.setTextVisible(True)
        self.setFormat("%p%")
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet("""
            QProgressBar#NiceProgress{
                border:1px solid #d0d5dd;
                background:#eef0f3;
                border-radius:7px;
                text-align:center;
                color:#2b2f36;
                font-weight:600;
            }
            QProgressBar#NiceProgress::chunk{
                border-radius:6px;
                background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                                           stop:0 #2f6fed, stop:1 #0b57d0);
            }
        """)

class ClickResetFilter(QObject):
    """
    Global event filter:
    - left click blank area (not on interactive widgets) => reset_fn()
    """
    def __init__(self, reset_fn, interactive_widgets: tuple[type, ...]):
        super().__init__()
        self.reset_fn = reset_fn
        self.interactive_widgets = interactive_widgets

    def eventFilter(self, obj, event):
        if event.type() == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
            w = QApplication.widgetAt(event.globalPosition().toPoint())
            if w is None:
                self.reset_fn()
                return False
            for t in self.interactive_widgets:
                if isinstance(w, t):
                    return False
            self.reset_fn()
        return False

def make_card_vertical(title: str, buttons: list[str]) -> QWidget:
    """Create a vertical card with a title and a list of soft buttons."""
    w = QFrame()
    w.setObjectName("Card")
    w.setFrameShape(QFrame.NoFrame)

    lay = QVBoxLayout(w)
    lay.setContentsMargins(12, 12, 12, 12)
    lay.setSpacing(10)

    t = QLabel(title)
    t.setObjectName("CardTitle")
    lay.addWidget(t)

    for text in buttons:
        b = QPushButton(text)
        b.setObjectName("SoftBtn")
        b.setMinimumHeight(40)
        b.setMinimumWidth(170)
        lay.addWidget(b)

    lay.addStretch(1)
    return w

class DropLineEdit(QLineEdit):
    """Drag-drop line edit for Excel/Word/PDF paths."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, e):
        allowed_ext = (".xlsx", ".xls", ".doc", ".docx", ".pdf")
        if e.mimeData().hasUrls():
            urls = e.mimeData().urls()
            if urls and urls[0].toLocalFile().lower().endswith(allowed_ext):
                e.acceptProposedAction()
                return
        super().dragEnterEvent(e)

    def dragMoveEvent(self, e):
        self.dragEnterEvent(e)

    def dropEvent(self, e):
        allowed_ext = (".xlsx", ".xls", ".doc", ".docx", ".pdf")
        urls = e.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            if path.lower().endswith(allowed_ext):
                self.setText(path)
                return
        super().dropEvent(e)
