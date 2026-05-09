from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPainter
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QVBoxLayout, QWidget


COLORS = {
    "red": QColor("#dc2626"),
    "orange": QColor("#f97316"),
    "yellow": QColor("#d97706"),
    "green": QColor("#16a34a"),
    "blue": QColor("#2563eb"),
    "teal": QColor("#0f766e"),
    "muted": QColor("#e5e7eb"),
    "text": QColor("#111827"),
}


class StatCard(QFrame):
    def __init__(self, title: str, accent: str = "blue"):
        super().__init__()
        self.accent = accent
        self.setObjectName("statCard")
        self.setMinimumHeight(92)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(4)
        self.title_label = QLabel(title)
        self.title_label.setObjectName("statTitle")
        self.value_label = QLabel("-")
        self.value_label.setObjectName("statValue")
        self.note_label = QLabel("")
        self.note_label.setObjectName("statNote")
        self.note_label.setWordWrap(True)
        layout.addWidget(self.title_label)
        layout.addWidget(self.value_label)
        layout.addWidget(self.note_label)

    def set_value(self, value: object, note: str = "") -> None:
        self.value_label.setText(str(value))
        self.note_label.setText(note)


class SeverityBar(QWidget):
    def __init__(self):
        super().__init__()
        self.counts = {"Red": 0, "Orange": 0, "Yellow": 0}
        self.setMinimumHeight(126)

    def set_counts(self, red: int, orange: int, yellow: int) -> None:
        self.counts = {"Red": red, "Orange": orange, "Yellow": yellow}
        self.update()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        rect = self.rect().adjusted(4, 8, -4, -8)
        total = max(sum(self.counts.values()), 1)
        labels = [
            ("Red", self.counts["Red"], COLORS["red"]),
            ("Orange", self.counts["Orange"], COLORS["orange"]),
            ("Yellow", self.counts["Yellow"], COLORS["yellow"]),
        ]
        bar_left = rect.left()
        bar_top = rect.top() + 36
        bar_width = rect.width()
        bar_height = 18
        painter.setPen(Qt.NoPen)
        painter.setBrush(COLORS["muted"])
        painter.drawRoundedRect(bar_left, bar_top, bar_width, bar_height, 9, 9)

        x = bar_left
        for _, count, color in labels:
            width = int(bar_width * count / total)
            if count and width < 8:
                width = 8
            if width:
                painter.setBrush(color)
                painter.drawRoundedRect(x, bar_top, width, bar_height, 9, 9)
                x += width

        painter.setPen(COLORS["text"])
        font = painter.font()
        font.setBold(True)
        font.setPointSize(10)
        painter.setFont(font)
        painter.drawText(rect.left(), rect.top() + 18, "风险分布")

        font.setBold(False)
        font.setPointSize(9)
        painter.setFont(font)
        y = bar_top + 52
        x = rect.left()
        for label, count, color in labels:
            painter.setBrush(color)
            painter.setPen(Qt.NoPen)
            painter.drawRoundedRect(x, y - 10, 12, 12, 4, 4)
            painter.setPen(COLORS["text"])
            painter.drawText(x + 18, y, f"{label}: {count}")
            x += 120


class FileTypeStrip(QWidget):
    def __init__(self):
        super().__init__()
        self.counts = {"表格": 0, "PDF": 0, "图片": 0, "其他": 0}
        self.setMinimumHeight(126)

    def set_counts(self, total: int, tables: int, pdfs: int, images: int) -> None:
        other = max(total - tables - pdfs - images, 0)
        self.counts = {"表格": tables, "PDF": pdfs, "图片": images, "其他": other}
        self.update()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        rect = self.rect().adjusted(4, 8, -4, -8)
        total = max(sum(self.counts.values()), 1)
        colors = [COLORS["blue"], COLORS["teal"], COLORS["orange"], QColor("#64748b")]

        painter.setPen(COLORS["text"])
        font = painter.font()
        font.setBold(True)
        font.setPointSize(10)
        painter.setFont(font)
        painter.drawText(rect.left(), rect.top() + 18, "文件组成")

        bar_left = rect.left()
        bar_top = rect.top() + 36
        bar_width = rect.width()
        bar_height = 18
        painter.setPen(Qt.NoPen)
        painter.setBrush(COLORS["muted"])
        painter.drawRoundedRect(bar_left, bar_top, bar_width, bar_height, 9, 9)
        x = bar_left
        for (_, count), color in zip(self.counts.items(), colors):
            width = int(bar_width * count / total)
            if count and width < 8:
                width = 8
            if width:
                painter.setBrush(color)
                painter.drawRoundedRect(x, bar_top, width, bar_height, 9, 9)
                x += width

        font.setBold(False)
        font.setPointSize(9)
        painter.setFont(font)
        y = bar_top + 52
        slot = max(rect.width() // 4, 72)
        for (label, count), color in zip(self.counts.items(), colors):
            x = rect.left() + list(self.counts.keys()).index(label) * slot
            painter.setBrush(color)
            painter.setPen(Qt.NoPen)
            painter.drawRoundedRect(x, y - 10, 12, 12, 4, 4)
            painter.setPen(COLORS["text"])
            painter.drawText(x + 18, y, f"{label}: {count}")


class DashboardPanel(QFrame):
    def __init__(self, title: str):
        super().__init__()
        self.setObjectName("dashboardPanel")
        self.header = QLabel(title)
        self.header.setObjectName("panelTitle")
        self.body = QVBoxLayout()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 16)
        layout.addWidget(self.header)
        layout.addLayout(self.body)


def two_column_row(left: QWidget, right: QWidget) -> QWidget:
    row = QWidget()
    layout = QHBoxLayout(row)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(12)
    layout.addWidget(left)
    layout.addWidget(right)
    return row
