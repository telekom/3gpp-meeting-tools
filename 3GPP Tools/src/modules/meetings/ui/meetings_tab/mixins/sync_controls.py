"""Sync menu, custom sync launcher, and inline live progress banner for MeetingsTab."""

from __future__ import annotations

from typing import Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMenu,
    QProgressBar,
    QPushButton,
    QSizePolicy,
)

from core.ui.font_scale import scale_px
from core.ui.page_ribbon import apply_lucide_icon
from core.ui.theme import C, R, S, T
from core.ui.toast import notify_info, notify_success
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_DANGER, BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.meetings.core.crawler.sources import MEETING_SOURCES
from modules.meetings.ui.meetings_sync_dialog import MeetingsSyncDialog


class MeetingsTabSyncMixin:
    """Provides the intent-driven Sync menu and inline sync progress banner."""

    def _setup_sync_controls(self) -> None:
        """Build the Sync dropdown button/menu and inline sync progress banner."""
        self.sync_menu_btn = QPushButton("Sync ▾")
        self.sync_menu_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.sync_menu_btn.setMinimumHeight(scale_px(30))
        self.sync_menu_btn.setCursor(Qt.PointingHandCursor)
        apply_lucide_icon(self.sync_menu_btn, "refresh-cw", size=14)
        self.sync_menu_btn.setToolTip(
            "Synchronize 3GPP meeting folders, official DynaReport dates/Portal IDs, and TDoc ranges."
        )

        self.sync_menu = QMenu(self)
        self.sync_menu.setToolTipsVisible(True)
        self.sync_menu.aboutToShow.connect(self._refresh_sync_menu)
        self._refresh_sync_menu()

        self.sync_menu_btn.setMenu(self.sync_menu)
        self.ribbon_card.add_utility_widget(self.sync_menu_btn)
        self.update_btn = self.sync_menu_btn  # Compatibility alias
        self.btn_add_meeting = self.sync_menu_btn  # Compatibility alias

        # Inline live sync progress banner (hidden when idle)
        self.sync_status_banner = QFrame(self.ribbon_card)
        self.sync_status_banner.setObjectName("meetingsSyncBanner")
        self.sync_status_banner.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        self.sync_status_banner.setStyleSheet(
            f"""
            QFrame#meetingsSyncBanner {{
                background-color: {C.INFO_SOFT};
                border: 1px solid {C.INFO_BORDER};
                border-radius: {R.MD}px;
            }}
            """
        )
        self.sync_status_banner.setVisible(False)
        banner_layout = QHBoxLayout(self.sync_status_banner)
        banner_layout.setContentsMargins(S.SM + 2, S.XS + 1, S.SM + 2, S.XS + 1)
        banner_layout.setSpacing(S.SM)

        self.sync_status_label = QLabel("⏳ Synchronizing 3GPP meetings...", self.sync_status_banner)
        self.sync_status_label.setWordWrap(True)
        self.sync_status_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.sync_status_label.setStyleSheet(
            f"color: {C.INFO_DARK}; font-size: {scale_px(T.SIZE_CAPTION)}px; "
            f"font-weight: 600; background: transparent; border: none;"
        )
        banner_layout.addWidget(self.sync_status_label, 1, Qt.AlignVCenter)

        self.sync_progress_bar = QProgressBar(self.sync_status_banner)
        self.sync_progress_bar.setTextVisible(False)
        self.sync_progress_bar.setRange(0, 0)
        self.sync_progress_bar.setFixedHeight(scale_px(10))
        self.sync_progress_bar.setMinimumWidth(scale_px(130))
        self.sync_progress_bar.setMaximumWidth(scale_px(210))
        self.sync_progress_bar.setStyleSheet(
            f"""
            QProgressBar {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.INFO_BORDER};
                border-radius: {R.SM}px;
            }}
            QProgressBar::chunk {{
                background-color: {C.ACCENT_PRIMARY};
                border-radius: {R.SM}px;
            }}
            """
        )
        banner_layout.addWidget(self.sync_progress_bar, 0, Qt.AlignVCenter)

        self.btn_stop_sync = QPushButton("Stop", self.sync_status_banner)
        self.btn_stop_sync.setStyleSheet(BUTTON_STYLE_TOOLBAR_DANGER)
        self.btn_stop_sync.setMinimumHeight(scale_px(26))
        self.btn_stop_sync.setCursor(Qt.PointingHandCursor)
        apply_lucide_icon(self.btn_stop_sync, "square", color=C.DANGER_DARK, size=11)
        self.btn_stop_sync.setToolTip("Stop the active meetings synchronization cleanly.")
        self.btn_stop_sync.clicked.connect(self.abort_sync_requested.emit)
        banner_layout.addWidget(self.btn_stop_sync, 0, Qt.AlignVCenter)

        self.ribbon_card.add_banner_widget(self.sync_status_banner)

    def _get_selected_known_wgs(self) -> list:
        if not hasattr(self, "wg_filter"):
            return []
        return [w for w in self.wg_filter.getCheckedItems() if w in MEETING_SOURCES]

    def _refresh_sync_menu(self) -> None:
        self.sync_menu.clear()

        act_quick = self.sync_menu.addAction(
            "⚡ Quick Sync — All WGs (~15–30s)",
            lambda: self.request_quick_sync(),
        )
        act_quick.setToolTip(
            "Scans all WG FTP folders, updates dates & Portal IDs via DynaReport, and refreshes recent/missing Docs/ ranges."
        )

        selected_wgs = self._get_selected_known_wgs()
        if 0 < len(selected_wgs) < len(MEETING_SOURCES):
            if len(selected_wgs) <= 3:
                wg_summary = ", ".join(selected_wgs)
            else:
                wg_summary = f"{', '.join(selected_wgs[:3])} +{len(selected_wgs) - 3}"
            act_wg = self.sync_menu.addAction(
                f"🎯 Sync {wg_summary} Only (~3–8s)",
                lambda _=False, wgs=list(selected_wgs): self.request_quick_sync(target_wgs=wgs),
            )
            act_wg.setToolTip(
                f"Quickly sync FTP folders, DynaReport dates/Portal IDs, and recent Docs/ only for: {', '.join(selected_wgs)}."
            )

        act_custom = self.sync_menu.addAction(
            "🔍 Custom / Deep Sync...",
            self._open_custom_sync_dialog,
        )
        act_custom.setToolTip(
            "Choose specific Working Groups, sync steps, or a full historical Docs/ re-scan."
        )

        act_add = self.sync_menu.addAction(
            "➕ Add / Fetch Specific Meeting...",
            self._open_add_meeting_dialog,
        )
        act_add.setToolTip(
            "Manually query, preview, and import a specific meeting via Portal ID or FTP URL."
        )

        self.sync_menu.addSeparator()
        act_cfg = self.sync_menu.addAction(
            "⚙️ Cache && Preferences...",
            self.config_dialog.exec_,
        )
        act_cfg.setToolTip("Configure local cache directory, Work Item enrichment, and database maintenance.")

    def request_quick_sync(self, target_wgs: Optional[list] = None) -> None:
        """Trigger a fast-first sync (FTP folders + DynaReport + smart Docs/)."""
        wgs = [w for w in (target_wgs or []) if w in MEETING_SOURCES]
        label = f"3GPP_Meetings_{','.join(wgs[:3])}" if wgs else "3GPP_Meetings"
        self.update_meetings_requested.emit(
            {
                "sync_wg": True,
                "sync_dyna": True,
                "docs_mode": "smart",
                "target_wgs": wgs,
                "label": label,
            }
        )

    def _open_custom_sync_dialog(self) -> None:
        dlg = MeetingsSyncDialog(initial_wgs=self._get_selected_known_wgs(), parent=self)
        if dlg.exec_() == QDialog.Accepted:
            self.update_meetings_requested.emit(dlg.get_sync_options())

    def on_sync_started(self, display_name: str = "") -> None:
        """Show the inline sync progress banner when a meetings sync starts."""
        self.sync_menu_btn.setText("Syncing…")
        self.sync_status_label.setText(f"⏳ Synchronizing meetings ({display_name or '3GPP'})...")
        self.sync_progress_bar.setRange(0, 0)
        self.btn_stop_sync.setEnabled(True)
        self.sync_status_banner.setVisible(True)
        self.sync_status_banner.updateGeometry()
        self.ribbon_card.updateGeometry()

    def on_sync_progress(self, message: str, current: int, total: int) -> None:
        """Update the inline sync progress banner with live step and progress counts."""
        if not self.sync_status_banner.isVisible():
            self.sync_status_banner.setVisible(True)
        self.sync_status_label.setText(f"⏳ {message}")
        if total > 0:
            self.sync_progress_bar.setRange(0, total)
            self.sync_progress_bar.setValue(max(0, min(current, total)))
        else:
            self.sync_progress_bar.setRange(0, 0)

    def on_sync_finished(self, summary: Optional[dict] = None) -> None:
        """Hide the inline sync banner, refresh UI state, and notify the user."""
        summary = summary or {}
        self.sync_menu_btn.setText("Sync ▾")
        self.sync_status_banner.setVisible(False)
        self.ribbon_card.updateGeometry()

        self._populate_filters()
        self.refresh_table()
        self.refresh_quick_tdoc_lookup()

        elapsed = summary.get("elapsed_s", 0)
        if summary.get("cancelled"):
            notify_info(self, "Meetings Sync Stopped", f"Stopped after {elapsed}s.")
            return
        if summary.get("error"):
            return

        parts = []
        new_cnt = int(summary.get("new_meetings", 0) or 0)
        upd_cnt = int(summary.get("updated_meetings", 0) or 0)
        docs_cnt = int(summary.get("docs_scanned", 0) or 0)
        if new_cnt:
            parts.append(f"+{new_cnt} new")
        if upd_cnt:
            parts.append(f"{upd_cnt} updated")
        if docs_cnt:
            parts.append(f"{docs_cnt} Docs/ scanned")
        detail = ", ".join(parts) if parts else "Database up to date"
        notify_success(self, "Meetings Synced", f"{detail} ({elapsed}s).")
