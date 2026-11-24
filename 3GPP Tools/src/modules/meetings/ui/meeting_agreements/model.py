"""Table model and preview formatting for Chairman's Notes meeting agreements."""

from __future__ import annotations

import html
import re
from typing import List, Optional

from PyQt5.QtCore import QAbstractTableModel, QModelIndex, Qt
from PyQt5.QtGui import QColor

from core.ui.theme import C


def _clean_agreement_preview(text: str, max_chars: int = 140) -> str:
    """Produces a clean, single-line preview from multi-line agreement text."""
    if not text:
        return ""
    # Strip leading tags and header labels
    t = re.sub(
        r"^(?:🟢|🟡|🔵)?\s*\[\s*(?:AGREEMENT|WORKING ASSUMPTION|CONCLUSION|DECISION|FFS)\s*\]\s*:?",
        "",
        text,
        flags=re.IGNORECASE,
    )
    t = re.sub(r"^(?:Agreements?|Working\s+Assumptions?)\s*[:\n]?", "", t, flags=re.IGNORECASE)
    lines = [line.strip().lstrip("•-*1234567890. ") for line in t.splitlines() if line.strip()]
    joined = " ".join(lines)
    if len(joined) > max_chars:
        return joined[:max_chars].rstrip() + "..."
    return joined


class AgreementsTableModel(QAbstractTableModel):
    """Table model representing meeting agreements for the viewer dialog."""

    HEADERS = ["Kind", "TDoc", "Agenda Item", "Work Item", "Agreement Summary"]

    def __init__(self, agreements: List[dict], parent=None):
        super().__init__(parent)
        self._all_agreements: List[dict] = list(agreements or [])
        self._filtered_agreements: List[dict] = list(self._all_agreements)

    def rowCount(self, parent=QModelIndex()):
        return len(self._filtered_agreements)

    def columnCount(self, parent=QModelIndex()):
        return len(self.HEADERS)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.HEADERS[section]
        return None

    def get_row(self, row: int) -> Optional[dict]:
        if 0 <= row < len(self._filtered_agreements):
            return self._filtered_agreements[row]
        return None

    def data(self, index: QModelIndex, role=Qt.DisplayRole):
        if not index.isValid() or index.row() >= len(self._filtered_agreements):
            return None

        ag = self._filtered_agreements[index.row()]
        col = index.column()
        kind = str(ag.get("kind", "Agreement") or "Agreement").strip()
        is_wa = "working assumption" in kind.lower()
        is_ffs = "ffs" in kind.lower() or "open issue" in kind.lower()
        is_pcr = "pcr" in kind.lower() or "p-cr" in kind.lower()
        is_cr = kind.lower().endswith("cr") and not is_pcr
        is_ls = "ls" in kind.lower()

        if role == Qt.DisplayRole:
            if col == 0:
                if is_pcr:
                    return "📝 Agreed pCR"
                elif is_cr:
                    return "📄 Approved CR"
                elif is_ls:
                    return "📤 Approved LS"
                elif is_wa:
                    return "🟡 Working Assumption"
                elif is_ffs:
                    return "🔵 FFS / Open Issue"
                return "🟢 Agreement"
            elif col == 1:
                return ag.get("tdoc") or ag.get("source_tdoc") or "AI Level"
            elif col == 2:
                ai = str(ag.get("ai") or "").strip()
                return ai if ai else "-"
            elif col == 3:
                wi = str(ag.get("wi") or "").strip()
                return wi if wi and wi.lower() != "general" else "-"
            elif col == 4:
                txt = ag.get("text", "")
                if not txt.strip() or txt.strip().lower() in ("agreed", "approved", "noted"):
                    txt = ag.get("title", "")
                return _clean_agreement_preview(txt)

        elif role == Qt.BackgroundRole:
            if is_wa:
                return QColor(C.WARNING_SOFT_LIGHT)  # Amber 50
            elif is_ffs:
                return QColor(C.INFO_SOFT)  # Blue 50
            return QColor(C.SUCCESS_SOFT)  # Emerald 50

        elif role == Qt.ForegroundRole:
            if col == 0:
                if is_wa:
                    return QColor(C.WARNING)
                elif is_ffs:
                    return QColor(C.INFO)
                return QColor(C.SUCCESS)
            elif col == 1:
                return QColor(C.TEXT_PRIMARY)

        elif role == Qt.TextAlignmentRole:
            if col in (0, 1, 2, 3):
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        elif role == Qt.ToolTipRole:
            txt = ag.get("text", "")
            return f"<div style='max-width: 450px; white-space: pre-wrap;'>{html.escape(txt)}</div>"

        return None

    def apply_filter(self, search_text: str = "", kind_filter: str = "ALL", ai_filter: str = "ALL"):
        """Filters agreements by search query, agreement kind, and agenda item."""
        self.beginResetModel()
        query = (search_text or "").lower().strip()
        kind_req = (kind_filter or "ALL").upper()
        ai_req = (ai_filter or "ALL").strip()

        filtered = []
        for ag in self._all_agreements:
            # 1. Kind filter
            k = str(ag.get("kind", "")).lower()
            if kind_req == "AGREEMENT" and (
                "working assumption" in k or "ffs" in k or "pcr" in k or "cr" in k or "ls" in k
            ):
                continue
            elif kind_req == "PCR" and ("pcr" not in k and "p-cr" not in k):
                continue
            elif kind_req == "CR" and ("cr" not in k or "pcr" in k or "p-cr" in k):
                continue
            elif kind_req == "LS" and "ls" not in k:
                continue
            elif kind_req == "WA" and "working assumption" not in k:
                continue
            elif kind_req == "FFS" and ("ffs" not in k and "open issue" not in k):
                continue

            # 2. AI filter
            if ai_req != "ALL":
                ai_val = str(ag.get("ai", "")).strip()
                if ai_val != ai_req:
                    continue

            # 3. Text search
            if query:
                searchable = " ".join([
                    str(ag.get("tdoc", "")),
                    str(ag.get("ai", "")),
                    str(ag.get("wi", "")),
                    str(ag.get("title", "")),
                    str(ag.get("company", "")),
                    str(ag.get("text", "")),
                ]).lower()
                if query not in searchable:
                    continue

            filtered.append(ag)

        self._filtered_agreements = filtered
        self.endResetModel()
