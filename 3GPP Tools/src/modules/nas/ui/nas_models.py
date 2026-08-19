from typing import Any, List
import pandas as pd
from PyQt5.QtCore import QAbstractTableModel, QModelIndex, Qt
from PyQt5.QtGui import QBrush, QColor

from modules.nas.core.nas_db import parse_version_tuple


class NASEvolutionMatrixModel(QAbstractTableModel):
    """Pivots Information Elements across multiple versions while strictly preserving specification row order."""

    def __init__(self, raw_df: pd.DataFrame = None):
        super().__init__()
        self._raw_df = raw_df if raw_df is not None else pd.DataFrame()
        self._pivot_df = pd.DataFrame()
        self._versions: List[str] = []
        self._setup_matrix()

    def _setup_matrix(self):
        if self._raw_df.empty:
            self._pivot_df = pd.DataFrame()
            self._versions = []
            return

        df = self._raw_df.copy()
        df["details"] = (
            df["presence"].fillna("")
            + " | "
            + df["format"].fillna("")
            + " | "
            + df["length"].fillna("")
        )

        # Chronological version ordering for horizontal diffing
        self._versions = sorted(df["version"].unique().tolist(), key=parse_version_tuple)

        # 1. Compute canonical specification row order for each distinct IE
        order_map = (
            df.groupby(["iei", "ie_name", "type_reference"])["order_index"]
            .min()
            .reset_index()
        )

        # 2. Pivot: Rows = IEI + IE Name + Type, Columns = Versions
        pivot = df.pivot_table(
            index=["iei", "ie_name", "type_reference"],
            columns="version",
            values="details",
            aggfunc="first",
        ).reset_index()

        # 3. Merge order_map and sort rows strictly by specification order
        merged = pd.merge(
            pivot, order_map, on=["iei", "ie_name", "type_reference"], how="left"
        )
        merged = merged.sort_values(by="order_index", ascending=True).reset_index(drop=True)

        # 4. Drop order_index so it remains hidden from the table view
        self._pivot_df = merged.drop(columns=["order_index"])

        # Fill missing values
        for v in self._versions:
            if v not in self._pivot_df.columns:
                self._pivot_df[v] = "-"
            else:
                self._pivot_df[v] = self._pivot_df[v].fillna("-")

    def rowCount(self, parent=QModelIndex()) -> int:
        return len(self._pivot_df)

    def columnCount(self, parent=QModelIndex()) -> int:
        return len(self._pivot_df.columns)

    def data(self, index: QModelIndex, role=Qt.DisplayRole) -> Any:
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()
        val = self._pivot_df.iloc[row, col]

        if role == Qt.DisplayRole:
            return str(val) if pd.notna(val) else "-"

        if role == Qt.TextAlignmentRole:
            if col >= 3:
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        # Visual diffing across version columns
        if role == Qt.BackgroundRole and col >= 3:
            current_ver_col = self._pivot_df.columns[col]
            current_val = str(val)

            if current_ver_col in self._versions:
                v_idx = self._versions.index(current_ver_col)
                if v_idx > 0:
                    prev_ver_col = self._versions[v_idx - 1]
                    prev_val = str(self._pivot_df.iloc[row][prev_ver_col])

                    if prev_val == "-" and current_val != "-":
                        return QBrush(QColor("#E8F5E9"))  # Added (Green)
                    elif prev_val != "-" and current_val == "-":
                        return QBrush(QColor("#FFEBEE"))  # Removed (Red)
                    elif (
                        prev_val != current_val
                        and prev_val != "-"
                        and current_val != "-"
                    ):
                        return QBrush(QColor("#FFF9C4"))  # Modified (Yellow)

        return None

    def headerData(
        self, section: int, orientation: Qt.Orientation, role=Qt.DisplayRole
    ) -> Any:
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            col_name = self._pivot_df.columns[section]
            header_map = {
                "iei": "IEI",
                "ie_name": "Information Element",
                "type_reference": "Type / Reference",
            }
            return header_map.get(col_name, f"v{col_name}")
        return None