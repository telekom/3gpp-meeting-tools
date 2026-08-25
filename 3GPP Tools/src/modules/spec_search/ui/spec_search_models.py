"""
Evolution Matrix Table Model for Substring Search Results.
Identifies and color-codes when specific text was first introduced into a specification clause.
"""

from typing import Any, Dict, List, Optional
import pandas as pd
from PyQt5.QtCore import QAbstractTableModel, QModelIndex, Qt
from PyQt5.QtGui import QBrush, QColor, QFont

from modules.spec_search.core.spec_search_db import parse_version_tuple


class SpecEvolutionMatrixModel(QAbstractTableModel):
    """
    Pivots search matches across versions while tracking the exact release when text was first introduced.
    """

    def __init__(self, raw_df: Optional[pd.DataFrame] = None, all_selected_versions: Optional[List[Dict[str, Any]]] = None):
        super().__init__()
        self._raw_df = raw_df if raw_df is not None else pd.DataFrame()
        self._all_selected_versions = all_selected_versions or []
        self._pivot_df = pd.DataFrame()
        self._version_cols: List[str] = []
        self._setup_matrix()

    def _setup_matrix(self):
        if self._raw_df.empty:
            self._pivot_df = pd.DataFrame()
            self._version_cols = []
            return

        df = self._raw_df.copy()

        # Build column label: 'TS 24.501 v18.4.0' if multi-spec, or 'v18.4.0'
        multi_spec = df["spec_number"].nunique() > 1
        if multi_spec:
            df["ver_label"] = "TS " + df["spec_number"] + " v" + df["version"]
        else:
            df["ver_label"] = "v" + df["version"]

        # Sort available version columns chronologically
        unique_vers = df[["version", "ver_label"]].drop_duplicates()
        sorted_ver_labels = unique_vers.sort_values(
            by="version", key=lambda s: s.map(parse_version_tuple)
        )["ver_label"].tolist()

        self._version_cols = sorted_ver_labels

        # Pivot to create Matrix
        pivot = df.pivot_table(
            index=["clause_number", "clause_title"],
            columns="ver_label",
            values="snippet_text",
            aggfunc="first",
        ).reset_index()

        # Compute initial appearance order
        order_map = df.groupby(["clause_number", "clause_title"], as_index=False)["order_index"].min()
        merged = pd.merge(pivot, order_map, on=["clause_number", "clause_title"], how="left")
        merged = merged.sort_values(by="order_index", ascending=True).reset_index(drop=True)
        self._pivot_df = merged.drop(columns=["order_index"])

        # Fill non-matches with empty marker
        for v in self._version_cols:
            if v not in self._pivot_df.columns:
                self._pivot_df[v] = "-"
            else:
                self._pivot_df[v] = self._pivot_df[v].fillna("-")

    def rowCount(self, parent=QModelIndex()) -> int:
        return len(self._pivot_df)

    def columnCount(self, parent=QModelIndex()) -> int:
        # Columns: Clause, Title, followed by version columns
        return 2 + len(self._version_cols)

    def data(self, index: QModelIndex, role=Qt.DisplayRole) -> Any:
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()

        if col == 0:
            val = self._pivot_df.iloc[row]["clause_number"]
            if role == Qt.DisplayRole:
                return str(val)
            if role == Qt.FontRole:
                f = QFont()
                f.setBold(True)
                return f

        elif col == 1:
            val = self._pivot_df.iloc[row]["clause_title"]
            if role == Qt.DisplayRole:
                return str(val)

        else:
            v_idx = col - 2
            ver_col = self._version_cols[v_idx]
            cell_val = str(self._pivot_df.iloc[row][ver_col])
            is_match = cell_val != "-"

            if role == Qt.DisplayRole:
                if not is_match:
                    return "-"

                # Check if this version is the FIRST introduction of the text
                if v_idx == 0:
                    return "🟢 Added"

                prev_ver_col = self._version_cols[v_idx - 1]
                prev_val = str(self._pivot_df.iloc[row][prev_ver_col])
                if prev_val == "-":
                    return "🟢 Added"
                return "✓ Present"

            if role == Qt.TextAlignmentRole:
                return Qt.AlignCenter

            if role == Qt.ToolTipRole:
                if is_match:
                    clean_snip = cell_val.replace("<mark>", "【").replace("</mark>", "】")
                    return f"Match in {ver_col}:\n{clean_snip}"
                return f"Text not present in {ver_col}"

            # Visual diff highlighting
            if role == Qt.BackgroundRole:
                if is_match:
                    if v_idx == 0 or str(self._pivot_df.iloc[row][self._version_cols[v_idx - 1]]) == "-":
                        return QBrush(QColor("#DCFCE7"))  # Soft Green (First Added)
                    return QBrush(QColor("#F0FDF4"))      # Very Pale Green (Retained)
                else:
                    if v_idx > 0 and str(self._pivot_df.iloc[row][self._version_cols[v_idx - 1]]) != "-":
                        return QBrush(QColor("#FEE2E2"))  # Soft Red (Removed in this version)
                    return None

            if role == Qt.ForegroundRole:
                if is_match:
                    if v_idx == 0 or str(self._pivot_df.iloc[row][self._version_cols[v_idx - 1]]) == "-":
                        return QBrush(QColor("#15803D"))  # Bold Dark Green
                    return QBrush(QColor("#166534"))
                return QBrush(QColor("#94A3B8"))

        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role=Qt.DisplayRole) -> Any:
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            if section == 0:
                return "Clause"
            if section == 1:
                return "Clause Title"
            return self._version_cols[section - 2]
        return None