from __future__ import annotations

import pandas as pd
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
from PySide6.QtGui import QBrush, QColor


class DataFrameModel(QAbstractTableModel):
    def __init__(self, dataframe: pd.DataFrame | None = None):
        super().__init__()
        self._df = dataframe if dataframe is not None else pd.DataFrame()

    def set_dataframe(self, dataframe: pd.DataFrame) -> None:
        self.beginResetModel()
        self._df = dataframe.copy()
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._df)

    def columnCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._df.columns)

    def data(self, index: QModelIndex, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        if role == Qt.DisplayRole:
            value = self._df.iat[index.row(), index.column()]
            return "" if pd.isna(value) else str(value)
        if role == Qt.BackgroundRole and "risk_level" in self._df.columns:
            risk = str(self._df.iloc[index.row()].get("risk_level", ""))
            colors = {
                "Red": QColor("#fee2e2"),
                "Orange": QColor("#ffedd5"),
                "Yellow": QColor("#fef9c3"),
            }
            if risk in colors:
                return QBrush(colors[risk])
        if role == Qt.TextAlignmentRole:
            return Qt.AlignVCenter | Qt.AlignLeft
        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section]) if section < len(self._df.columns) else ""
        return str(section + 1)


class HighlightDataFrameModel(DataFrameModel):
    def __init__(self, dataframe: pd.DataFrame | None = None):
        super().__init__(dataframe)
        self._highlight_rows: set[int] = set()
        self._highlight_columns: set[str] = set()
        self._highlight_cells: set[tuple[int, str]] = set()

    def set_highlights(
        self,
        rows: set[int] | None = None,
        columns: set[str] | None = None,
        cells: set[tuple[int, str]] | None = None,
    ) -> None:
        self._highlight_rows = rows or set()
        self._highlight_columns = columns or set()
        self._highlight_cells = cells or set()
        if self.rowCount() and self.columnCount():
            top_left = self.index(0, 0)
            bottom_right = self.index(self.rowCount() - 1, self.columnCount() - 1)
            self.dataChanged.emit(top_left, bottom_right, [Qt.BackgroundRole])

    def data(self, index: QModelIndex, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        if role == Qt.BackgroundRole:
            row_number = index.row() + 2
            column_name = str(self._df.columns[index.column()]) if index.column() < len(self._df.columns) else ""
            if (row_number, column_name) in self._highlight_cells:
                return QBrush(QColor("#facc15"))
            if row_number in self._highlight_rows:
                return QBrush(QColor("#dbeafe"))
            if column_name in self._highlight_columns:
                return QBrush(QColor("#dcfce7"))
        return super().data(index, role)
