# --- File: src/modules/meetings/ui/tdocs/window/mixins/filter_tracked_wis.py ---
"""
Tracked/starred Work Item filter buttons, their context menus and token loading.

Mixin of TDocsWindow: relies on attributes set in TDocsWindow.__init__ / _setup_*.
"""
import logging
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QMenu
from modules.meetings.core.settings import MeetingsSettings
from core.utils.paths import get_project_root
from modules.work_items.core.wi_settings import WorkItemsSettings
from modules.work_items.core.wi_database import WorkItemsDatabase
from core.ui.theme import C


class TrackedWIFilterMixin:
    """Tracked/starred Work Item filter buttons, their context menus and token loading."""

    def _toggle_starred_wis_filter(self):
        self._load_tracked_wi_tokens()
        self.proxy.setTrackedWIFilters(
            self.btn_followed_wis.isChecked(),
            self.btn_starred_wis.isChecked()
        )
        self._update_tracked_wi_button_styles()
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _show_tracked_wi_context_menu(self, pos):
        sender = self.sender()
        menu = QMenu(self)
        act = menu.addAction("Include unassigned Incoming LS (LS in)")
        act.setCheckable(True)
        act.setChecked(getattr(self.proxy, "include_unassigned_ls_in", True))
        act.toggled.connect(self._on_toggle_include_unassigned_ls_in)
        if sender and hasattr(sender, "mapToGlobal"):
            menu.exec_(sender.mapToGlobal(pos))
        else:
            from PyQt5.QtGui import QCursor
            menu.exec_(QCursor.pos())

    def _on_toggle_include_unassigned_ls_in(self, checked: bool):
        if not hasattr(self, "settings"):
            self.settings = MeetingsSettings()
        self.settings.set_include_unassigned_ls_in(checked)
        self.proxy.setIncludeUnassignedLsIn(checked)
        if hasattr(self, "_update_facet_counts"):
            self._update_facet_counts()
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _show_company_filter_context_menu(self, pos):
        sender = self.sender()
        menu = QMenu(self)
        act = menu.addAction("Include unconfirmed (?, [...]) carry-over sources")
        act.setCheckable(True)
        act.setChecked(getattr(self.proxy, "include_tentative_sources", False))
        act.setToolTip(
            "When checked, filtering by Company also matches revisions where the company is marked with ? or [...] as unconfirmed."
        )
        act.toggled.connect(self._on_toggle_include_tentative_sources)
        if sender and hasattr(sender, "mapToGlobal"):
            menu.exec_(sender.mapToGlobal(pos))
        else:
            menu.exec_(QCursor.pos())

    def _on_toggle_include_tentative_sources(self, checked: bool):
        if not hasattr(self, "settings"):
            self.settings = MeetingsSettings()
        self.settings.set_include_tentative_sources(checked)
        self.proxy.setIncludeTentativeSources(checked)
        self._refresh_comboboxes()
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _update_tracked_wi_button_styles(self):
        """Update active/inactive visual styles of Followed and Starred WI quick-filter chips."""
        if not hasattr(self, "btn_followed_wis") or not hasattr(self, "btn_starred_wis"):
            return

        if self.btn_followed_wis.isChecked():
            self.btn_followed_wis.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.INFO_SOFT_STRONG};
                    color: {C.INFO_DARK};
                    border: 1.5px solid {C.INFO_ACCENT};
                    border-radius: 4px;
                    padding: 2px;
                    font-size: 13px;
                }}
                QPushButton:hover {{
                    background-color: {C.INFO_BORDER};
                    border-color: {C.ACCENT_ALT};
                }}
            """)
        else:
            self.btn_followed_wis.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.BG_SURFACE};
                    color: {C.TEXT_SECONDARY};
                    border: 1px solid {C.BORDER_STRONG};
                    border-radius: 4px;
                    padding: 2px;
                    font-size: 13px;
                }}
                QPushButton:hover {{
                    background-color: {C.BG_CANVAS};
                    border-color: {C.TEXT_FAINT};
                }}
            """)

        if self.btn_starred_wis.isChecked():
            self.btn_starred_wis.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.WARNING_SOFT_ALT};
                    color: {C.WARNING_DARK};
                    border: 1.5px solid {C.WARNING_ACCENT};
                    border-radius: 4px;
                    padding: 2px;
                    font-size: 13px;
                }}
                QPushButton:hover {{
                    background-color: {C.WARNING_BORDER};
                    border-color: {C.WARNING_AMBER};
                }}
            """)
        else:
            self.btn_starred_wis.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.BG_SURFACE};
                    color: {C.TEXT_SECONDARY};
                    border: 1px solid {C.BORDER_STRONG};
                    border-radius: 4px;
                    padding: 2px;
                    font-size: 13px;
                }}
                QPushButton:hover {{
                    background-color: {C.BG_CANVAS};
                    border-color: {C.TEXT_FAINT};
                }}
            """)

        if hasattr(self, "btn_agreements_filter"):
            if self.btn_agreements_filter.isChecked():
                self.btn_agreements_filter.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {C.SUCCESS_SOFT_ALT};
                        color: {C.SUCCESS_DARK};
                        border: 1.5px solid #22C55E;
                        border-radius: 4px;
                        padding: 2px;
                        font-size: 13px;
                        font-weight: bold;
                    }}
                    QPushButton:hover {{
                        background-color: {C.SUCCESS_BORDER};
                        border-color: {C.GREEN_ACCENT};
                    }}
                """)
            else:
                self.btn_agreements_filter.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {C.BG_SURFACE};
                        color: {C.TEXT_SECONDARY};
                        border: 1px solid {C.BORDER_STRONG};
                        border-radius: 4px;
                        padding: 2px;
                        font-size: 13px;
                    }}
                    QPushButton:hover {{
                        background-color: {C.BG_CANVAS};
                        border-color: {C.TEXT_FAINT};
                    }}
                """)

        if hasattr(self, "btn_unread_emails"):
            if self.btn_unread_emails.isChecked():
                self.btn_unread_emails.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {C.INFO_SOFT_STRONG};
                        color: {C.INFO_DARK};
                        border: 1.5px solid {C.INFO_ACCENT};
                        border-radius: 4px;
                        padding: 2px;
                        font-size: 13px;
                        font-weight: bold;
                    }}
                    QPushButton:hover {{
                        background-color: {C.INFO_BORDER};
                        border-color: {C.ACCENT_ALT};
                    }}
                """)
            else:
                self.btn_unread_emails.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {C.BG_SURFACE};
                        color: {C.TEXT_SECONDARY};
                        border: 1px solid {C.BORDER_STRONG};
                        border-radius: 4px;
                        padding: 2px;
                        font-size: 13px;
                    }}
                    QPushButton:hover {{
                        background-color: {C.BG_CANVAS};
                        border-color: {C.TEXT_FAINT};
                    }}
                """)

        if hasattr(self, "btn_flagged_emails"):
            if self.btn_flagged_emails.isChecked():
                self.btn_flagged_emails.setStyleSheet(f"""
                    QPushButton {{
                        background-color: #FEE2E2;
                        color: #B91C1C;
                        border: 1.5px solid #DC2626;
                        border-radius: 4px;
                        padding: 2px;
                        font-size: 13px;
                        font-weight: bold;
                    }}
                    QPushButton:hover {{
                        background-color: #FECACA;
                        border-color: #B91C1C;
                    }}
                """)
            else:
                self.btn_flagged_emails.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {C.BG_SURFACE};
                        color: {C.TEXT_SECONDARY};
                        border: 1px solid {C.BORDER_STRONG};
                        border-radius: 4px;
                        padding: 2px;
                        font-size: 13px;
                    }}
                    QPushButton:hover {{
                        background-color: {C.BG_CANVAS};
                        border-color: {C.TEXT_FAINT};
                    }}
                """)

    def _load_tracked_wi_tokens(self):
        """Loads followed/starred WI codes from settings, resolves tokens from 3gpp_data.db, and updates proxy."""
        if not hasattr(self, "proxy"):
            return
        try:
            settings = WorkItemsSettings()
            starred_codes, following_codes = settings.get_tracking_sets()
            db_path = get_project_root() / "db" / "3gpp_data.db"
            if db_path.exists():
                wi_db = WorkItemsDatabase(db_path)
                followed_tokens = wi_db.get_tracked_wi_tokens(following_codes)
                starred_tokens = wi_db.get_tracked_wi_tokens(starred_codes)
            else:
                followed_tokens = set(following_codes)
                starred_tokens = set(starred_codes)
            self.proxy.setTrackedWITokens(followed_tokens, starred_tokens)
        except Exception as e:
            logging.error(f"Failed to load tracked WI tokens: {e}", exc_info=True)
            self.proxy.setTrackedWITokens(set(), set())
