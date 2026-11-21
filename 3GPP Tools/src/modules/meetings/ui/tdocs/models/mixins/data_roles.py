"""Qt data() role rendering mixin for TDocsTableModel."""

import html

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor

from core.ui.theme import C
from core.utils.company_sanitizer import CompanySanitizer


class TDocsDataRolesMixin:
    """Implements QAbstractTableModel.data() roles for emails, badges, notes, and tooltips."""

    def data(self, index, role):
        if not index.isValid():
            return None
        row = self._data[index.row()]
        col_name = self._headers[index.column()]

        # Subtle background tint for quick-opened highlight, agreements & working assumptions
        if role == Qt.BackgroundRole:
            if self._highlighted_tdoc and str(row.get("TDoc", "")).strip().upper() == self._highlighted_tdoc:
                return QColor(C.INFO_SOFT_STRONG)
            if row.get('_has_agreement'):
                kind = str(row.get('_agreement_kind', 'Agreement')).lower()
                if "working assumption" in kind:
                    return QColor(C.WARNING_SOFT_LIGHT)
                elif "ffs" in kind:
                    return QColor(C.INFO_SOFT)
                return QColor(C.SUCCESS_SOFT)
            return None

        # Render Emails badge column (Precomputed O(1) Lookups)
        if col_name == "Emails":
            tot = row.get('_email_tot', 0)
            unread = row.get('_email_unread', 0)
            flagged = row.get('_email_flagged', 0)

            if role == Qt.DisplayRole:
                if tot == 0:
                    return ""
                disp = f"✉️ {tot}"
                if unread > 0:
                    disp += f" (🔵 {unread})"
                if flagged > 0:
                    sep = "\n" if unread > 0 else " "
                    disp += f"{sep}🚩 {flagged}"
                return disp
            elif role == Qt.UserRole:
                if tot == 0:
                    return ""
                parts = []
                if unread > 0:
                    parts.append(f"unread: {unread}")
                if flagged > 0:
                    parts.append(f"flagged: {flagged}")
                return f"{tot} ({', '.join(parts)})" if parts else str(tot)
            elif role == Qt.UserRole + 2:
                return row.get('_email_sort_key', (0, 0, 0))
            elif role == Qt.ToolTipRole:
                if tot == 0:
                    return "No indexed emails mentioning this TDoc or its revisions."
                tdoc = str(row.get("TDoc", "")).upper()
                family = self.get_family_tdocs(tdoc)
                flag_txt = f", {flagged} flagged" if flagged > 0 else ""
                return f"<b>{tot} email(s) total</b> ({unread} unread{flag_txt}) mentioning family: {', '.join(family)}"
            elif role == Qt.TextAlignmentRole:
                return Qt.AlignCenter
            elif role == Qt.ForegroundRole and unread > 0:
                return QColor(C.OFFICE_BLUE)
            elif role == Qt.ForegroundRole and flagged > 0:
                return QColor(C.WARNING)
            return None

        if col_name == "":
            if role == Qt.UserRole:
                return row.get("TDoc", "")
            if role == Qt.UserRole + 1:
                if row.get("_is_tracked_thread"):
                    return "THREAD"
                tdoc = row.get("TDoc", "")
                if tdoc in self.loading_tdocs:
                    return "LOADING"
                if not self.meeting_dir:
                    return "MISSING"
                zip_path = self.meeting_dir / tdoc / f"{tdoc}.zip"
                return "EXISTS" if zip_path.exists() else "MISSING"
            if role == Qt.UserRole + 2:
                return len(self.revisions.get(row.get("TDoc", ""), [])) > 0
            return None

        # Expose sanitized companies directly to the ProxyModel
        if role == Qt.UserRole + 3 and col_name == "Source":
            return row.get('_Sanitized_Companies', ["Other"])

        # Expose Sub-Item directly to the AgendaItemDelegate and ProxyModel
        if role == Qt.UserRole + 4 and col_name == "Agenda Item":
            return str(row.get("Sub-Item", "") or "").strip()

        if role == Qt.ForegroundRole:
            if col_name == "TDoc Status" and row.get('_has_agreement'):
                kind = str(row.get('_agreement_kind', 'Agreement')).lower()
                if "working assumption" in kind:
                    return QColor(C.WARNING)
                elif "ffs" in kind:
                    return QColor(C.INFO)
                return QColor(C.SUCCESS)
            return None

        if role == Qt.UserRole + 2:
            val = row.get(col_name, "")
            return str(val).strip() if val is not None else ""

        if role == Qt.DisplayRole:
            if col_name == "Related TDocs":
                return self._format_related_tdocs(row, html=True)
            val = row.get(col_name, "")
            val_str = str(val).strip() if val is not None else ""

            if col_name == "Abstract":
                return "📝" if val_str else ""
            if col_name == "My Status" and val_str == "⚪ Neutral":
                return ""
            if col_name == "My Notes":
                tag_list = row.get("_tag_list") or []
                if tag_list and val_str:
                    tag_preview = ", ".join(tag_list[:2]) + ("…" if len(tag_list) > 2 else "")
                    return f"🏷️ {tag_preview} · 📓"
                if tag_list:
                    tag_preview = ", ".join(tag_list[:2]) + ("…" if len(tag_list) > 2 else "")
                    return f"🏷️ {tag_preview}"
                if val_str:
                    return "📓 Note"
                return ""

            if col_name == "Secretary Remarks":
                if row.get('_has_agreement'):
                    return self._format_agreement_remarks_html(val_str, row)
                linked = self._linkify("", val_str, html=True)
                return linked.replace('\n', '<br>')

            return val_str

        elif role == Qt.UserRole:
            if col_name == "Related TDocs":
                return self._format_related_tdocs(row, html=False)
            val = row.get(col_name, "")
            return str(val).strip() if val is not None else ""

        elif role == Qt.ToolTipRole:
            val = row.get(col_name, "")
            val_str = str(val).strip() if val is not None else ""

            if col_name == "Agenda Item":
                ai_num = val_str
                agenda_item = self.agenda_map.get(ai_num)
                ai_desc = ""
                if agenda_item and getattr(agenda_item, "description", None):
                    ai_desc = agenda_item.description
                else:
                    excel_desc = str(row.get("Agenda item description", "")).strip()
                    if excel_desc:
                        ai_desc = excel_desc

                tooltip_lines = []
                if ai_num:
                    tooltip_lines.append(f"<b>Agenda Item {ai_num}</b>: {ai_desc}" if ai_desc else f"<b>Agenda Item {ai_num}</b>")
                sub_item = str(row.get("Sub-Item", "") or "").strip()
                if sub_item:
                    tooltip_lines.append(f"<b>Sub-Item:</b> {sub_item}")
                return "<br>".join(tooltip_lines) if tooltip_lines else None

            if (col_name in ["Secretary Remarks", "TDoc Status"]) and row.get('_has_agreement'):
                return self._format_agreement_tooltip_html(row)

            if col_name == "My Notes":
                tag_list = row.get("_tag_list") or []
                parts = []
                if tag_list:
                    parts.append(f"<b>🏷️ Tags:</b> {html.escape(', '.join(tag_list))}")
                if val_str:
                    parts.append(f"<b>📓 Notes:</b><br>{html.escape(val_str)}")
                return f"<div style='width: 400px; white-space: pre-wrap;'>{'<br><br>'.join(parts)}</div>" if parts else None

            if col_name in ["Abstract", "Secretary Remarks"] and val_str:
                return f"<div style='width: 400px; white-space: pre-wrap;'>{val_str}</div>"
            elif col_name == "Source" and val_str:
                if CompanySanitizer.has_tentative_sources(val_str):
                    return CompanySanitizer.build_source_tooltip_html(val_str)
                if len(val_str) > 30:
                    return val_str
            elif col_name == "Title" and len(val_str) > 30:
                return val_str
            return None

        elif role == Qt.TextAlignmentRole:
            if col_name in [
                "TDoc", "Type", "For", "Abstract", "My Status", "My Notes",
                "Agenda Item", "TDoc Status", "Related TDocs"
            ]:
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter
