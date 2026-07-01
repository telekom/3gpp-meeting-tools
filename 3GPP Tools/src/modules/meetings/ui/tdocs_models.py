import re
from pathlib import Path

from PyQt5.QtCore import QAbstractTableModel, Qt, QModelIndex, QSortFilterProxyModel


def natural_sort_key(s):
    """Splits string into chunks of digits and non-digits for natural sorting."""
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

class TDocsTableModel(QAbstractTableModel):
    def __init__(self, meeting_dir: Path, data=None):
        super().__init__()
        self.meeting_dir = meeting_dir
        self._data = data or []
        self._headers = [
            "", "TDoc", "Title", "Source", "Type", "For",
            "Abstract", "Secretary Remarks", "Agenda Item", "TDoc Status", "Related TDocs"
        ]
        self.valid_tdocs = {str(r.get("TDoc", "")) for r in self._data if r.get("TDoc")}
        self.loading_tdocs = set()
        self.revisions = {}  # Stores {base_tdoc: ["r01", "r02"]}

    def update_data(self, new_data):
        self.beginResetModel()
        self._data = new_data
        self.valid_tdocs = {str(r.get("TDoc", "")) for r in self._data if r.get("TDoc")}
        self.loading_tdocs.clear()
        # Note: We specifically DO NOT clear self.revisions here so they persist during updates
        self.endResetModel()

    def set_loading(self, tdoc: str, is_loading: bool):
        if is_loading:
            self.loading_tdocs.add(tdoc)
        else:
            self.loading_tdocs.discard(tdoc)
        for r in range(self.rowCount()):
            if self._data[r].get("TDoc") == tdoc:
                idx = self.index(r, 0)
                self.dataChanged.emit(idx, idx)
                break

    def _linkify(self, prefix: str, text: str, html: bool) -> str:
        if not text: return ""
        if not html: return f"{prefix}: {text}"

        def repl(match):
            tdoc = match.group(0)
            tdoc_upper = tdoc.upper()

            # 1. Check if the exact string is in the meeting
            is_local = (tdoc in self.valid_tdocs) or (tdoc_upper in self.valid_tdocs)

            # 2. If it's not, check if it's a revision string (e.g., S2-2606287R11)
            # whose BASE TDoc (S2-2606287) belongs to this meeting!
            if not is_local:
                import re
                base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_upper)
                if base_match:
                    base_tdoc = base_match.group(1)
                    if base_tdoc in self.valid_tdocs:
                        is_local = True

            if is_local:
                # ---> Local TDoc (or local revision): Bold Blue
                return f'<a href="{tdoc}" style="color: #005A9E; font-weight: bold; text-decoration: underline;">{tdoc}</a>'
            else:
                # ---> External TDoc: Dark Orange
                return f'<a href="{tdoc}" style="color: #D83B01; text-decoration: underline;">{tdoc}</a>'

        # Safely capture TDocs and optional revision suffixes
        linked_text = re.sub(r'[a-zA-Z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', repl, text, flags=re.IGNORECASE)
        return f"<span style='color: #444;'><b>{prefix}:</b></span> {linked_text}"

    def _format_related_tdocs(self, row_data: dict, html=False) -> str:
        parts = []
        if r_rev_of := row_data.get("Is revision of"): parts.append(self._linkify("⬅️ Rev of", r_rev_of, html))
        if r_rev_to := row_data.get("Revised to"): parts.append(self._linkify("➡️ Rev to", r_rev_to, html))
        if r_orig := row_data.get("Original LS"): parts.append(self._linkify("✉️ Orig LS", r_orig, html))
        if r_reply := row_data.get("Reply in"): parts.append(self._linkify("↩️ Reply", r_reply, html))
        return ("<br>" if html else "\n").join(parts)

    def data(self, index, role):
        if not index.isValid(): return None
        row = self._data[index.row()]
        col_name = self._headers[index.column()]

        if col_name == "":
            if role == Qt.UserRole: return row.get("TDoc", "")
            if role == Qt.UserRole + 1:
                tdoc = row.get("TDoc", "")
                if tdoc in self.loading_tdocs: return "LOADING"
                zip_path = self.meeting_dir / tdoc / f"{tdoc}.zip"
                return "EXISTS" if zip_path.exists() else "MISSING"
            if role == Qt.UserRole + 2:
                # Indicates if revisions exist in memory for this TDoc
                return len(self.revisions.get(row.get("TDoc", ""), [])) > 0
            return None

        if role == Qt.UserRole + 2:
            val = row.get(col_name, "")
            return str(val).strip() if val is not None else ""

        if role == Qt.DisplayRole:
            if col_name == "Related TDocs": return self._format_related_tdocs(row, html=True)
            val = row.get(col_name, "")
            val_str = str(val).strip() if val is not None else ""

            # ---> NEW: Replace Abstract text with an icon
            if col_name == "Abstract":
                return "📝" if val_str else ""

            # Keep the truncation logic for Secretary Remarks
            if col_name == "Secretary Remarks" and len(val_str) > 90:
                return val_str.replace('\n', ' ').replace('\r', '')[:87] + "..."
            return val_str

        elif role == Qt.UserRole:
            if col_name == "Related TDocs": return self._format_related_tdocs(row, html=False)
            val = row.get(col_name, "")
            return str(val).strip() if val is not None else ""

        elif role == Qt.ToolTipRole:
            val = row.get(col_name, "")
            val_str = str(val).strip() if val is not None else ""

            if col_name in ["Abstract", "Secretary Remarks"] and val_str:
                return f"<div style='width: 400px; white-space: pre-wrap;'>{val_str}</div>"
            elif col_name in ["Title", "Source"] and len(val_str) > 30:
                return val_str
            return None

        elif role == Qt.TextAlignmentRole:
            # ---> UPDATE: Added "Abstract" to the list of centered columns
            if col_name in ["TDoc", "Type", "For", "Abstract", "Agenda Item", "TDoc Status", "Related TDocs"]:
                return Qt.AlignCenter

            # Everything else stays left-aligned and vertically centered
            return Qt.AlignLeft | Qt.AlignVCenter

    def rowCount(self, index=QModelIndex()):
        return len(self._data)

    def columnCount(self, index=QModelIndex()):
        return len(self._headers)

    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal:
            col_name = self._headers[section]

            if role == Qt.DisplayRole:
                # Replace the wide word with a compact icon
                return "📝" if col_name == "Abstract" else col_name

            elif role == Qt.ToolTipRole:
                # Add a tooltip to the header so users know what the icon means
                return "Abstract" if col_name == "Abstract" else col_name

        return None

    def merge_agenda_data(self, agenda_data: dict, ui_logger=None):
        """Merges parsed HTML data into the existing Excel data set."""
        import logging
        import re
        self.beginResetModel()

        existing_tdocs = {row.get('TDoc', ''): idx for idx, row in enumerate(self._data)}
        tdoc_dict = {row.get('TDoc', ''): row for row in self._data}  # Build lookup early
        new_rows = []

        for tdoc_id, info in agenda_data.items():
            comments = info.get('Comments', '')
            email = info.get('e-mail_Discussion', '')

            remarks = ""
            if comments and email:
                remarks = f"{comments}\n\n[e-mail_Discussion]: {email}"
            elif comments:
                remarks = comments
            elif email:
                remarks = email

            if tdoc_id in existing_tdocs:
                if remarks:
                    idx = existing_tdocs[tdoc_id]
                    self._data[idx]['Secretary Remarks'] = remarks
            else:
                if ui_logger: ui_logger.emit(f"✨ Injecting new on-the-fly TDoc: {tdoc_id}", logging.INFO)

                agenda_item = 'N/A'
                doc_type = 'Revision'  # Temporary default
                predecessor = None
                base_tdoc = None

                # ---> CHECK 1: Explicit comment "Revision of S2-XXXXXXr11"
                comment_match = re.search(r'(?:revision of|rev of)\s*(S2-\d{6,8}(?:r\d{1,2}[a-zA-Z]?)?)', remarks,
                                          re.IGNORECASE)

                # ---> CHECK 2: Standard suffix (e.g. S2-260123r01)
                id_match = re.search(r'^(.*?)-?(?:r|rev)(\d{1,2}[a-zA-Z]?)$', tdoc_id, re.IGNORECASE)

                if comment_match:
                    predecessor = comment_match.group(1).upper()
                    base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', predecessor, re.IGNORECASE)
                    base_tdoc = base_match.group(1).upper() if base_match else predecessor

                elif id_match:
                    base_tdoc = id_match.group(1).upper()
                    predecessor = base_tdoc
                    try:
                        rev_num = int(re.match(r'\d+', id_match.group(2)).group())
                        if rev_num > 1:
                            prev_rev = f"{base_tdoc}r{rev_num - 1:02d}"
                            if prev_rev in tdoc_dict:
                                predecessor = prev_rev
                    except Exception:
                        pass

                # If we found a base document, inherit its exact Agenda Item and Type!
                if base_tdoc and base_tdoc in tdoc_dict:
                    agenda_item = tdoc_dict[base_tdoc].get('Agenda Item', 'N/A')
                    doc_type = tdoc_dict[base_tdoc].get('Type', 'TDoc')

                new_row = {
                    'TDoc': tdoc_id,
                    'Title': info.get('Title', 'Unknown (Parsed from Agenda)'),
                    'Source': info.get('Source', ''),
                    'Secretary Remarks': remarks,
                    'Agenda Item': agenda_item,
                    'Type': doc_type,
                    'TDoc Status': 'Unknown',
                    'Is revision of': predecessor if predecessor else '',
                    'Revised to': '',
                    'Original LS': '',
                    'Reply in': ''
                }
                new_rows.append(new_row)

        if new_rows:
            self._data.extend(new_rows)
            # Re-sort chronologically
            self._data.sort(key=lambda x: str(x.get('TDoc', '')))

            # ---> UPDATE PREDECESSORS: Build relationships for the newly injected rows
            tdoc_dict = {row.get('TDoc', ''): row for row in self._data}

            for new_row in new_rows:
                tdoc_id = new_row['TDoc']
                predecessor = new_row['Is revision of']

                if predecessor:
                    # ---> THE PHANTOM FIX: If the predecessor (e.g. S2-2606287r11) isn't a standalone row,
                    # strip the 'r11' and link it to the BASE TDoc (S2-2606287) instead!
                    if predecessor not in tdoc_dict:
                        base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', predecessor, re.IGNORECASE)
                        if base_match:
                            base_tdoc = base_match.group(1).upper()
                            if base_tdoc in tdoc_dict:
                                predecessor = base_tdoc

                                # Update the predecessor bi-directionally
                    if predecessor in tdoc_dict:
                        curr_revs = str(tdoc_dict[predecessor].get('Revised to', '')).strip()
                        if tdoc_id not in curr_revs:
                            if curr_revs and curr_revs != 'None':
                                tdoc_dict[predecessor]['Revised to'] = f"{curr_revs}, {tdoc_id}"
                            else:
                                tdoc_dict[predecessor]['Revised to'] = tdoc_id

        self.valid_tdocs = {str(r.get("TDoc", "")) for r in self._data if r.get("TDoc")}
        self.endResetModel()


class TDocsFilterProxyModel(QSortFilterProxyModel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.global_filter = ""
        self.type_filters = set()
        self.status_filters = set()
        self.ai_filters = set()
        self.filter_no_comments = False

    def setNoCommentsFilter(self, enabled: bool):
        self.filter_no_comments = enabled
        self.invalidateFilter()

    def setGlobalFilter(self, text):
        self.global_filter = str(text).lower().strip()
        self.invalidateFilter()

    def setTypeFilters(self, types):
        self.type_filters = set(types)
        self.invalidateFilter()

    def setStatusFilters(self, statuses):
        self.status_filters = set(statuses)
        self.invalidateFilter()

    def setAIFilters(self, ais):
        self.ai_filters = set(ais)
        self.invalidateFilter()

    def lessThan(self, left, right):
        left_data = self.sourceModel().data(left, Qt.UserRole + 2)
        right_data = self.sourceModel().data(right, Qt.UserRole + 2)

        # Only apply natural sort to Agenda Items
        if self.sourceModel()._headers[left.column()] == "Agenda Item":
            return natural_sort_key(left_data) < natural_sort_key(right_data)

        return super().lessThan(left, right)

    def filterAcceptsRow(self, source_row, source_parent):
        model = self.sourceModel()

        # ---> 1. Apply the 'No Comments' logic first for speed
        if self.filter_no_comments:
            # Column 7 is "Secretary Remarks" in your headers list
            remarks = model.data(model.index(source_row, 7, source_parent), Qt.UserRole)
            if remarks and str(remarks).strip():
                return False  # Exclude this row because it HAS comments

        # ---> 2. Standard categorical filters
        if model.data(model.index(source_row, 4, source_parent), Qt.UserRole) not in self.type_filters: return False
        if model.data(model.index(source_row, 8, source_parent), Qt.UserRole) not in self.ai_filters: return False
        if model.data(model.index(source_row, 9, source_parent), Qt.UserRole) not in self.status_filters: return False

        # ---> 3. Global Text Search
        if self.global_filter:
            match_found = False
            for col in [1, 2, 3, 6, 10]:
                data = model.data(model.index(source_row, col, source_parent), Qt.UserRole)
                if data and self.global_filter in str(data).lower():
                    match_found = True
                    break
            if not match_found: return False

        return True