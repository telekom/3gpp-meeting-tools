from typing import Any, Dict, List, Optional
import pandas as pd
from PyQt5.QtCore import QAbstractTableModel, QModelIndex, Qt
from PyQt5.QtGui import QBrush, QColor

from modules.nas.core.nas_db import parse_version_tuple

# Pre-allocated brush constants to eliminate GC overhead during rendering
BRUSH_ADDED = QBrush(QColor("#E8F5E9"))      # Soft Green
BRUSH_REMOVED = QBrush(QColor("#FFEBEE"))    # Soft Red
BRUSH_MODIFIED = QBrush(QColor("#FFF9C4"))   # Soft Yellow


class NASEvolutionMatrixModel(QAbstractTableModel):
    def __init__(
        self,
        raw_df: pd.DataFrame = None,
        ie_filter: Optional[str] = None,
        search_descriptions: bool = False,
        interface_filter: Optional[str] = None,
    ):
        super().__init__()
        self._raw_df = raw_df if raw_df is not None else pd.DataFrame()
        self._ie_filter = ie_filter.strip().lower() if ie_filter else None
        self._search_descriptions = search_descriptions
        self._interface_filter = interface_filter.strip().upper() if interface_filter and interface_filter.upper() != "ALL" else None
        self._pivot_df = pd.DataFrame()
        self._versions: List[str] = []
        self._appl_list: List[str] = []

        # Fast memory caches for O(1) cell access
        self._visible_columns: List[str] = []
        self._data_matrix: List[List[str]] = []
        self._depth_list: List[int] = []
        self._tooltip_list: List[str] = []
        self._bg_brush_matrix: List[List[Optional[QBrush]]] = []

        self._setup_matrix()

    def _setup_matrix(self):
        if self._raw_df.empty:
            self._pivot_df = pd.DataFrame()
            self._versions = []
            self._visible_columns = []
            self._data_matrix = []
            self._depth_list = []
            self._tooltip_list = []
            self._bg_brush_matrix = []
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
        agg_dict = {"order_index": "min"}
        if "applicability" in df.columns:
            agg_dict["applicability"] = "first"

        order_map = (
            df.groupby(["iei", "ie_name", "field_path", "type_reference", "depth"], as_index=False)
            .agg(agg_dict)
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

        # 4. Filter by Interface Applicability (Option B)
        if self._interface_filter and "applicability" in self._pivot_df.columns:
            target_if = self._interface_filter
            appl_series = self._pivot_df["applicability"].fillna("").astype(str).str.upper()
            mask_if = appl_series.str.contains(target_if) | (appl_series == "") | (appl_series == "ALL")
            self._pivot_df = self._pivot_df[mask_if].reset_index(drop=True)

        for v in self._versions:
            if v not in self._pivot_df.columns:
                self._pivot_df[v] = "-"
            else:
                self._pivot_df[v] = self._pivot_df[v].fillna("-")

        # 5. Filter search query
        if self._ie_filter and not self._pivot_df.empty:
            q = self._ie_filter
            mask = (
                self._pivot_df["ie_name"].astype(str).str.lower().str.contains(q, na=False)
                | self._pivot_df["field_path"].astype(str).str.lower().str.contains(q, na=False)
                | self._pivot_df["type_reference"].astype(str).str.lower().str.contains(q, na=False)
                | self._pivot_df["iei"].astype(str).str.lower().str.contains(q, na=False)
            )

            # Vectorized description match
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
                    row_keys = list(
                        zip(
                            self._pivot_df["iei"].astype(str),
                            self._pivot_df["field_path"].astype(str),
                            self._pivot_df["type_reference"].astype(str),
                        )
                    )
                    desc_mask = pd.Series([k in matching_keys for k in row_keys], index=self._pivot_df.index)
                    mask = mask | desc_mask

            self._pivot_df = self._pivot_df[mask].reset_index(drop=True)

        # 5. Pre-build fast access caches
        self._build_fast_caches()

    def _build_fast_caches(self):
        self._visible_columns = [c for c in self._pivot_df.columns if c not in ("field_path", "depth", "applicability")]
        num_rows = len(self._pivot_df)
        num_cols = len(self._visible_columns)

        if num_rows == 0:
            self._data_matrix = []
            self._depth_list = []
            self._tooltip_list = []
            self._bg_brush_matrix = []
            return

        self._depth_list = self._pivot_df.get("depth", pd.Series([0] * num_rows)).fillna(0).astype(int).tolist()
        self._tooltip_list = self._pivot_df.get("field_path", pd.Series([""] * num_rows)).fillna("").astype(str).tolist()

        # Build raw string matrix
        raw_columns_data = [self._pivot_df[col].fillna("-").astype(str).tolist() for col in self._visible_columns]
        self._data_matrix = [
            [raw_columns_data[c][r] for c in range(num_cols)]
            for r in range(num_rows)
        ]

        # Precompute visual diff brushes
        self._bg_brush_matrix = [[None] * num_cols for _ in range(num_rows)]
        ver_col_indices = [
            (c_idx, self._visible_columns[c_idx])
            for c_idx in range(3, num_cols)
            if self._visible_columns[c_idx] in self._versions
        ]

        for c_idx, ver_col in ver_col_indices:
            v_idx = self._versions.index(ver_col)
            if v_idx > 0:
                prev_ver_col = self._versions[v_idx - 1]
                if prev_ver_col in self._visible_columns:
                    prev_c_idx = self._visible_columns.index(prev_ver_col)
                    for r_idx in range(num_rows):
                        current_val = self._data_matrix[r_idx][c_idx]
                        prev_val = self._data_matrix[r_idx][prev_c_idx]

                        if prev_val == "-" and current_val != "-":
                            self._bg_brush_matrix[r_idx][c_idx] = BRUSH_ADDED
                        elif prev_val != "-" and current_val == "-":
                            self._bg_brush_matrix[r_idx][c_idx] = BRUSH_REMOVED
                        elif prev_val != current_val and prev_val != "-" and current_val != "-":
                            self._bg_brush_matrix[r_idx][c_idx] = BRUSH_MODIFIED

        if "applicability" in self._pivot_df.columns:
            self._appl_list = self._pivot_df["applicability"].fillna("").astype(str).tolist()
        else:
            self._appl_list = [""] * num_rows

    def rowCount(self, parent=QModelIndex()) -> int:
        return len(self._data_matrix)

    def columnCount(self, parent=QModelIndex()) -> int:
        return len(self._visible_columns)

    def _get_visible_column_name(self, col: int) -> str:
        return self._visible_columns[col] if 0 <= col < len(self._visible_columns) else ""

    def data(self, index: QModelIndex, role=Qt.DisplayRole) -> Any:
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()

        if role == Qt.DisplayRole:
            val = self._data_matrix[row][col]
            if col == 1:
                depth = self._depth_list[row]
                if depth > 0:
                    return f"{'    ' * (depth - 1)}└─ {val}"
            return val

        if role == Qt.BackgroundRole:
            if col >= 3:
                return self._bg_brush_matrix[row][col]
            return None

        if role == Qt.ToolTipRole:
            field_path = self._tooltip_list[row]
            appl = self._appl_list[row] if row < len(self._appl_list) else ""
            tip = f"Path: {field_path}" if field_path else ""
            if appl:
                tip += f"\nApplicability: {appl}"
            return tip if tip else None

        if role == Qt.TextAlignmentRole:
            if col >= 3:
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter

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