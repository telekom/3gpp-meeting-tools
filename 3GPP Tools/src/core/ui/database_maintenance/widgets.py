"""Reusable UI sub-widgets for the Database Maintenance Hub."""

from __future__ import annotations

from PyQt5.QtWidgets import QFrame, QLabel, QVBoxLayout

from core.ui.theme import C

from .styles import (
    KPI_CARD_STYLE,
    KPI_SUB_STYLE,
    KPI_TITLE_STYLE,
    build_kpi_value_style,
)


class KpiCard(QFrame):
    """Executive KPI metric card with uppercase title, bold value, and subtle description."""

    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.setObjectName("kpiCard")
        self.setStyleSheet(KPI_CARD_STYLE)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(2)

        self.title_label = QLabel(title.upper())
        self.title_label.setStyleSheet(KPI_TITLE_STYLE)

        self.value_label = QLabel("—")
        self.value_label.setWordWrap(False)
        self.value_label.setStyleSheet(build_kpi_value_style(C.TEXT_PRIMARY))

        self.sub_label = QLabel("")
        self.sub_label.setStyleSheet(KPI_SUB_STYLE)

        layout.addWidget(self.title_label)
        layout.addWidget(self.value_label)
        layout.addWidget(self.sub_label)

    def set_data(self, value: str, sub: str, value_color: str = C.TEXT_PRIMARY):
        self.value_label.setText(value)
        self.value_label.setStyleSheet(build_kpi_value_style(value_color))
        self.sub_label.setText(sub)
