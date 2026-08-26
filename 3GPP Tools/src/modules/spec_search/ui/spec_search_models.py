"""
Evolution Matrix Table Model for Substring Search Results.
Identifies and color-codes when specific text was first introduced into a specification clause,
with support for Date cutoff filtering.
"""

from typing import Any, Dict, List, Optional
import pandas as pd
from PyQt5.QtCore import QAbstractTableModel, QModelIndex, Qt
from PyQt5.QtGui import QBrush, QColor, QFont

from modules.spec_search.core.spec_search_db import parse_version_tuple


class SpecEvolutionMatrixModel(QAbstractTableModel):
    """
    Pivots search matches across versions while tracking the exact release when text was first introduced.
    Supports filtering and highlighting matches relative to a patent priority cutoff date.
    """

    def __init__(
        self,
        raw_df: Optional[pd.DataFrame] = None,
        cutoff_date: Optional[str] = None,
        only_added_after_cutoff: bool = False,
    ):
        super().__init__()
        self._raw_df = raw_df if raw_df is not None else pd.DataFrame()
        self._cutoff_date = cutoff_date.strip() if cutoff_date and cutoff_date.strip() else None
        self._only_added_after_cutoff = only_added_after_cutoff
        self._pivot_df = pd.DataFrame()
        self._version_cols: List[str] = []
        self._ver_date_map: Dict[str, str] = {}
        self._setup_matrix()

    def _setup_matrix(self):
        if self._raw_df.empty:
            self._pivot_df = pd.DataFrame()
            self._version_cols = []
            self._ver_date_map = {}
            return

        df = self._raw_df.copy()

        multi_spec = df["spec_number"].nunique() > 1
        if multi_spec:
            df["ver_label"] = "TS " + df["spec_number"] + " v" + df["version"]
        else:
            df["ver_label"] = "v" + df["version"]

        # Cache release dates
        for _, r in df[["ver_label", "release_date"]].drop_duplicates().iterrows():
            self._ver_date_map[r["ver_label"]] = str(r["release_date"]) if pd.notna(r["release_date"]) and str(r["release_date"]).strip() else ""

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
        merged_pivot = merged.drop(columns=["order_index"])

        # Fill non-matches with empty marker
        for v in self._version_cols:
            if v not in merged_pivot.columns:
                merged_pivot[v] = "-"
            else:
                merged_pivot[v] = merged_pivot[v].fillna("-")

        # Filter by Patent Priority Date Cutoff if requested
        if self._cutoff_date and self._only_added_after_cutoff:
            rows_to_keep = []
            for idx, row in merged_pivot.iterrows():
                first_added_ver = None
                for v_idx, v_col in enumerate(self._version_cols):
                    if str(row[v_col]) != "-":
                        if v_idx == 0 or str(row[self._version_cols[v_idx - 1]]) == "-":
                            first_added_ver = v_col
                            break

                if first_added_ver:
                    v_date = self._ver_date_map.get(first_added_ver, "")
                    # If first introduced on/after cutoff date, keep row
                    if v_date and v_date >= self._cutoff_date:
                        rows_to_keep.append(idx)
                    elif not v_date:
                        rows_to_keep.append(idx)

            self._pivot_df = merged_pivot.loc[rows_to_keep].reset_index(drop=True)
        else:
            self._pivot_df = merged_pivot

    def rowCount(self, parent=QModelIndex()) -> int:
        return len(self._pivot_df)

    def columnCount(self, parent=QModelIndex()) -> int:
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
            is_first_added = is_match and (v_idx == 0 or str(self._pivot_df.iloc[row][self._version_cols[v_idx - 1]]) == "-")
            ver_date = self._ver_date_map.get(ver_col, "")

            is_post_cutoff = False
            if is_first_added and self._cutoff_date and ver_date and ver_date >= self._cutoff_date:
                is_post_cutoff = True

            if role == Qt.DisplayRole:
                if not is_match:
                    return "-"
                if is_first_added:
                    return "⚡ Post-Cutoff Added" if is_post_cutoff else "🟢 Added"
                return "✓ Present"

            if role == Qt.TextAlignmentRole:
                return Qt.AlignCenter

            if role == Qt.ToolTipRole:
                date_str = f" [Date: {ver_date}]" if ver_date else ""
                if is_match:
                    clean_snip = cell_val.replace("<mark>", "【").replace("</mark>", "】")
                    status_prefix = ""
                    if is_first_added:
                        status_prefix = f"⚠️ FIRST INTRODUCED IN THIS VERSION{date_str}\n"
                    return f"Match in {ver_col}{date_str}:\n{status_prefix}{clean_snip}"
                return f"Text not present in {ver_col}{date_str}"

            # Visual diff highlighting
            if role == Qt.BackgroundRole:
                if is_match:
                    if is_post_cutoff:
                        return QBrush(QColor("#FEF08A"))  # Soft Yellow for Data Post-Cutoff additions
                    if is_first_added:
                        return QBrush(QColor("#DCFCE7"))  # Soft Green (First Added)
                    return QBrush(QColor("#F0FDF4"))      # Pale Green (Retained)
                else:
                    if v_idx > 0 and str(self._pivot_df.iloc[row][self._version_cols[v_idx - 1]]) != "-":
                        return QBrush(QColor("#FEE2E2"))  # Soft Red (Removed in this version)
                    return None

            if role == Qt.ForegroundRole:
                if is_match:
                    if is_post_cutoff:
                        return QBrush(QColor("#B45309"))  # Dark Amber
                    if is_first_added:
                        return QBrush(QColor("#15803D"))  # Bold Dark Green
                    return QBrush(QColor("#166534"))
                return QBrush(QColor("#94A3B8"))

        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role=Qt.DisplayRole) -> Any:
        if orientation == Qt.Horizontal:
            if role == Qt.DisplayRole:
                if section == 0:
                    return "Clause"
                if section == 1:
                    return "Clause Title"
                ver_col = self._version_cols[section - 2]
                ver_date = self._ver_date_map.get(ver_col, "")
                return f"{ver_col}\n({ver_date})" if ver_date else ver_col

            if role == Qt.ToolTipRole and section >= 2:
                ver_col = self._version_cols[section - 2]
                ver_date = self._ver_date_map.get(ver_col, "Unknown date")
                return f"Specification Release: {ver_col}\nOfficial Date: {ver_date}"
        return None