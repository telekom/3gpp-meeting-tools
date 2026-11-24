"""UI layout and design-token QSS setup mixin for AppUpdateDialog."""

from __future__ import annotations

import webbrowser

from PyQt5.QtCore import QSize, Qt
from PyQt5.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QPushButton,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.config.config import HELP_URL
from core.ui.icons import lucide_icon
from core.ui.theme import C, R, S, T


class AppUpdateUiSetupMixin:
    """Widget hierarchy and styling setup for AppUpdateDialog."""

    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(S.LG, S.LG, S.LG, S.LG)
        layout.setSpacing(S.MD)

        self.summary_card = QFrame()
        self.summary_card.setObjectName("updateSummaryCard")
        card_layout = QVBoxLayout(self.summary_card)
        card_layout.setContentsMargins(S.MD, S.MD, S.MD, S.MD)
        card_layout.setSpacing(6)

        top_row = QHBoxLayout()
        self.status_title_lbl = QLabel("Checking for updates...")
        self.status_title_lbl.setObjectName("updateStatusTitle")
        top_row.addWidget(self.status_title_lbl, 1)

        self.open_repo_btn = QPushButton("Open GitHub Repo")
        self.open_repo_btn.setCursor(Qt.PointingHandCursor)
        self.open_repo_btn.setIcon(lucide_icon("external-link", color=C.TEXT_SECONDARY, size=14))
        self.open_repo_btn.setIconSize(QSize(14, 14))
        self.open_repo_btn.clicked.connect(self._open_github_commits)
        top_row.addWidget(self.open_repo_btn)
        card_layout.addLayout(top_row)

        meta_row = QHBoxLayout()
        meta_row.setSpacing(S.LG)

        col_local = QVBoxLayout()
        col_local.setSpacing(2)
        lbl_l = QLabel("Installed Version")
        lbl_l.setObjectName("updateMetaCaption")
        self.local_ver_lbl = QLabel("Inspecting...")
        self.local_ver_lbl.setObjectName("updateMetaValue")
        col_local.addWidget(lbl_l)
        col_local.addWidget(self.local_ver_lbl)
        meta_row.addLayout(col_local, 1)

        col_remote = QVBoxLayout()
        col_remote.setSpacing(2)
        lbl_r = QLabel("Latest on GitHub (master)")
        lbl_r.setObjectName("updateMetaCaption")
        self.remote_ver_lbl = QLabel("—")
        self.remote_ver_lbl.setObjectName("updateMetaValue")
        col_remote.addWidget(lbl_r)
        col_remote.addWidget(self.remote_ver_lbl)
        meta_row.addLayout(col_remote, 1)

        card_layout.addLayout(meta_row)

        self.notice_lbl = QLabel("")
        self.notice_lbl.setObjectName("updateNoticeLabel")
        self.notice_lbl.setWordWrap(True)
        self.notice_lbl.setVisible(False)
        card_layout.addWidget(self.notice_lbl)

        layout.addWidget(self.summary_card)

        feed_header = QLabel("What's Changed on master")
        feed_header.setObjectName("updateSectionHeader")
        layout.addWidget(feed_header)

        self.changes_browser = QTextBrowser()
        self.changes_browser.setObjectName("updateChangesBrowser")
        self.changes_browser.setOpenExternalLinks(True)
        layout.addWidget(self.changes_browser, 1)

        pres_lbl = QLabel(
            "Local databases (3gpp_data.db), proxy/AI settings, user presets, PlantUML JAR, "
            "and output/ files are always preserved."
        )
        pres_lbl.setObjectName("updatePreservationHint")
        pres_lbl.setWordWrap(True)
        layout.addWidget(pres_lbl)

        self.progress_widget = QWidget()
        prog_layout = QVBoxLayout(self.progress_widget)
        prog_layout.setContentsMargins(0, 0, 0, 0)
        prog_layout.setSpacing(4)
        self.progress_label = QLabel("")
        self.progress_label.setObjectName("updateMetaCaption")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setTextVisible(True)
        prog_layout.addWidget(self.progress_label)
        prog_layout.addWidget(self.progress_bar)
        self.progress_widget.setVisible(False)
        layout.addWidget(self.progress_widget)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(S.SM)

        self.check_btn = QPushButton("Check Now")
        self.check_btn.setCursor(Qt.PointingHandCursor)
        self.check_btn.setIcon(lucide_icon("refresh-cw", color=C.TEXT_SECONDARY, size=14))
        self.check_btn.setIconSize(QSize(14, 14))
        self.check_btn.clicked.connect(self.start_check)
        btn_row.addWidget(self.check_btn)

        self.skip_btn = QPushButton("Skip This Version")
        self.skip_btn.setCursor(Qt.PointingHandCursor)
        self.skip_btn.clicked.connect(self._skip_current_update)
        self.skip_btn.setVisible(False)
        btn_row.addWidget(self.skip_btn)

        btn_row.addStretch(1)

        self.cancel_dl_btn = QPushButton("Cancel Download")
        self.cancel_dl_btn.setCursor(Qt.PointingHandCursor)
        self.cancel_dl_btn.clicked.connect(self._cancel_download)
        self.cancel_dl_btn.setVisible(False)
        btn_row.addWidget(self.cancel_dl_btn)

        self.force_zip_btn = QPushButton("Overwrite via ZIP Update...")
        self.force_zip_btn.setCursor(Qt.PointingHandCursor)
        self.force_zip_btn.clicked.connect(self._confirm_and_force_zip_update)
        self.force_zip_btn.setVisible(False)
        btn_row.addWidget(self.force_zip_btn)

        self.apply_btn = QPushButton("Update && Restart")
        self.apply_btn.setObjectName("updatePrimaryBtn")
        self.apply_btn.setCursor(Qt.PointingHandCursor)
        self.apply_btn.setIcon(lucide_icon("download", color=C.TEXT_ON_ACCENT, size=14))
        self.apply_btn.setIconSize(QSize(14, 14))
        self.apply_btn.clicked.connect(self._start_apply_update)
        self.apply_btn.setEnabled(False)
        btn_row.addWidget(self.apply_btn)

        self.close_btn = QPushButton("Close")
        self.close_btn.setCursor(Qt.PointingHandCursor)
        self.close_btn.clicked.connect(self.close)
        btn_row.addWidget(self.close_btn)

        layout.addLayout(btn_row)

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            f"""
            QDialog {{
                background-color: {C.BG_CANVAS};
            }}
            QFrame#updateSummaryCard {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.LG}px;
            }}
            QLabel#updateStatusTitle {{
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_SECTION_TITLE}px;
                font-weight: 700;
                color: {C.TEXT_PRIMARY};
            }}
            QLabel#updateSectionHeader {{
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_BODY}px;
                font-weight: 700;
                color: {C.TEXT_PRIMARY};
            }}
            QLabel#updateMetaCaption {{
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
                color: {C.TEXT_MUTED};
            }}
            QLabel#updateMetaValue {{
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_BODY}px;
                font-weight: 600;
                color: {C.TEXT_PRIMARY};
            }}
            QLabel#updateNoticeLabel {{
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
                color: {C.WARNING};
                background-color: {C.WARNING_SOFT};
                border: 1px solid {C.WARNING_BORDER};
                border-radius: {R.SM}px;
                padding: 6px 8px;
            }}
            QLabel#updatePreservationHint {{
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
                color: {C.TEXT_MUTED};
            }}
            QTextBrowser#updateChangesBrowser {{
                background-color: {C.BG_CANVAS};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.MD}px;
                padding: 4px;
            }}
            QPushButton#updatePrimaryBtn {{
                background-color: {C.ACCENT_PRIMARY};
                color: {C.TEXT_ON_ACCENT};
                border: 1px solid {C.ACCENT_PRIMARY};
                border-radius: {R.MD}px;
                padding: 6px 14px;
                font-weight: 600;
            }}
            QPushButton#updatePrimaryBtn:hover {{
                background-color: {C.ACCENT_HOVER};
            }}
            QPushButton#updatePrimaryBtn:disabled {{
                background-color: {C.BORDER_STRONG};
                border-color: {C.BORDER_STRONG};
                color: {C.TEXT_MUTED};
            }}
            """
        )

    def _open_github_commits(self) -> None:
        url = "https://github.com/telekom/3gpp-meeting-tools/commits/master/3GPP%20Tools"
        try:
            webbrowser.open(url)
        except Exception:
            webbrowser.open(HELP_URL)
