from typing import Any, List, Optional
import pandas as pd
from PyQt5.QtCore import QAbstractTableModel, QModelIndex, Qt
from PyQt5.QtGui import QBrush, QColor

from modules.nas.core.nas_db import parse_version_tuple


class NASEvolutionMatrixModel(QAbstractTableModel):
    """
    Pivots Information Elements & ASN.1 Fields across multiple versions while
    preserving specification row order, rendering indentation, and applying filtering.
    """

    def __init__(
        self,
        raw_df: pd.DataFrame = None,
        ie_filter: Optional[str] = None,
        search_descriptions: bool = False,
    ):
        super().__init__()
        self._raw_df = raw_df if raw_df is not None else pd.DataFrame()
        self._ie_filter = ie_filter.strip().lower() if ie_filter else None
        self._search_descriptions = search_descriptions
        self._pivot_df = pd.DataFrame()
        self._versions: List[str] = []
        self._setup_matrix()

    def _setup_matrix(self):
        if self._raw_df.empty:
            self._pivot_df = pd.DataFrame()
            self._versions = []
            return

        df = self._raw_df

        # Vectorized string formatting for details column
        has_depth = "depth" in df.columns and (df["depth"] > 0).any()
        if has_depth:
            details_series = df["presence"].fillna("") + " | " + df["format"].fillna("")
        else:
            details_series = (
                    df["presence"].fillna("")
                    + " | "
                    + df["format"].fillna("")
                    + " | "
                    + df["length"].fillna("")
            )

        df = df.assign(details=details_series)

        if "field_path" not in df.columns:
            df["field_path"] = df["ie_name"]
        if "depth" not in df.columns:
            df["depth"] = 0

        multiple_specs = df["spec_number"].nunique() > 1 if "spec_number" in df.columns else False
        if multiple_specs:
            df["ver_col"] = df["spec_number"] + " v" + df["version"]
        else:
            df["ver_col"] = df["version"]

        unique_vers = (
            df[["spec_number", "version", "ver_col"]].drop_duplicates()
            if "spec_number" in df.columns
            else df[["version", "ver_col"]].drop_duplicates()
        )
        sorted_ver_cols = unique_vers.sort_values(
            by="version", key=lambda s: s.map(parse_version_tuple)
        )["ver_col"].tolist()
        self._versions = sorted_ver_cols

        # 1. Compute canonical specification order
        order_map = (
            df.groupby(["iei", "ie_name", "field_path", "type_reference", "depth"], as_index=False)["order_index"]
            .min()
        )

        # 2. Pivot version columns
        pivot = df.pivot_table(
            index=["iei", "ie_name", "field_path", "type_reference", "depth"],
            columns="ver_col",
            values="details",
            aggfunc="first",
        ).reset_index()

        # 3. Sort by canonical specification order
        merged = pd.merge(
            pivot, order_map, on=["iei", "ie_name", "field_path", "type_reference", "depth"], how="left"
        )
        merged = merged.sort_values(by="order_index", ascending=True).reset_index(drop=True)
        self._pivot_df = merged.drop(columns=["order_index"])

        for v in self._versions:
            if v not in self._pivot_df.columns:
                self._pivot_df[v] = "-"
            else:
                self._pivot_df[v] = self._pivot_df[v].fillna("-")

        # 4. Filter search query
        if self._ie_filter and not self._pivot_df.empty:
            q = self._ie_filter
            mask = (
                    self._pivot_df["ie_name"].astype(str).str.lower().str.contains(q, na=False)
                    | self._pivot_df["field_path"].astype(str).str.lower().str.contains(q, na=False)
                    | self._pivot_df["type_reference"].astype(str).str.lower().str.contains(q, na=False)
                    | self._pivot_df["iei"].astype(str).str.lower().str.contains(q, na=False)
            )

            if self._search_descriptions and "ie_description" in df.columns:
                desc_match_df = df[df["ie_description"].astype(str).str.lower().str.contains(q, na=False)]
                if not desc_match_df.empty:
                    matching_keys = set(
                        zip(
                            desc_match_df["iei"].astype(str),
                            desc_match_df["field_path"].astype(str),
                            desc_match_df["type_reference"].astype(str),
                        )
                    )
                    desc_mask = self._pivot_df.apply(
                        lambda row: (str(row["iei"]), str(row["field_path"]),
                                     str(row["type_reference"])) in matching_keys,
                        axis=1,
                    )
                    mask = mask | desc_mask

            self._pivot_df = self._pivot_df[mask].reset_index(drop=True)

    def rowCount(self, parent=QModelIndex()) -> int:
        return len(self._pivot_df)

    def columnCount(self, parent=QModelIndex()) -> int:
        return len([c for c in self._pivot_df.columns if c not in ("field_path", "depth")])

    def _get_visible_column_name(self, col: int) -> str:
        visible_cols = [c for c in self._pivot_df.columns if c not in ("field_path", "depth")]
        return visible_cols[col]

    def data(self, index: QModelIndex, role=Qt.DisplayRole) -> Any:
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()
        col_name = self._get_visible_column_name(col)
        val = self._pivot_df.iloc[row][col_name]

        if role == Qt.DisplayRole:
            if col_name == "ie_name":
                depth = self._pivot_df.iloc[row].get("depth", 0)
                if depth > 0:
                    indent = "    " * (depth - 1) + "└─ "
                    return f"{indent}{val}"
                return str(val)
            return str(val) if pd.notna(val) else "-"

        if role == Qt.ToolTipRole:
            field_path = self._pivot_df.iloc[row].get("field_path", "")
            if field_path:
                return f"Path: {field_path}"
            return None

        if role == Qt.TextAlignmentRole:
            if col >= 3:
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        if role == Qt.BackgroundRole and col >= 3:
            current_ver_col = col_name
            current_val = str(val)

            if current_ver_col in self._versions:
                v_idx = self._versions.index(current_ver_col)
                if v_idx > 0:
                    prev_ver_col = self._versions[v_idx - 1]
                    prev_val = str(self._pivot_df.iloc[row][prev_ver_col])

                    if prev_val == "-" and current_val != "-":
                        return QBrush(QColor("#E8F5E9"))  # Added
                    elif prev_val != "-" and current_val == "-":
                        return QBrush(QColor("#FFEBEE"))  # Removed
                    elif (
                        prev_val != current_val
                        and prev_val != "-"
                        and current_val != "-"
                    ):
                        return QBrush(QColor("#FFF9C4"))  # Modified

        return None

    def headerData(
        self, section: int, orientation: Qt.Orientation, role=Qt.DisplayRole
    ) -> Any:
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            col_name = self._get_visible_column_name(section)
            header_map = {
                "iei": "IEI / ID",
                "ie_name": "Field / Information Element",
                "type_reference": "Type / Reference",
            }
            if col_name in header_map:
                return header_map[col_name]
            return col_name if col_name.startswith("TS ") or col_name.startswith("v") else f"v{col_name}"
        return None