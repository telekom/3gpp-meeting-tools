"""UI layout construction mixin for TopicDossierDialog."""

from PyQt5.QtCore import QDate, Qt
from PyQt5.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDateEdit,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMenu,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.theme import C
from modules.meetings.ui.dialogs.topic_dossier.styles import (
    STYLE_CARD,
    STYLE_PRIMARY_BTN,
    STYLE_SECONDARY_BTN,
)


class TopicDossierUiSetupMixin:
    """Builds the 2-column Meta-Topics navigation list, scope picker, and action bar."""

    def _init_ui(self):
        root_layout = QHBoxLayout(self)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(12)

        # ==========================================
        # LEFT PANE: Topics Navigation & Management
        # ==========================================
        left_frame = QFrame()
        left_frame.setFixedWidth(260)
        left_layout = QVBoxLayout(left_frame)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        lbl_topics = QLabel("<b>🗂️ Meta-Topics</b>")
        lbl_topics.setStyleSheet(f"font-size: 13px; color: {C.TEXT_STRONG};")
        left_layout.addWidget(lbl_topics)

        self.list_topics = QListWidget()
        self.list_topics.setTextElideMode(Qt.ElideRight)
        self.list_topics.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.list_topics.setStyleSheet(f"""
            QListWidget {{
                outline: none;
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                padding: 4px;
            }}
            QListWidget::item {{
                padding: 8px 10px;
                border: 1px solid transparent;
                border-radius: 5px;
                margin-bottom: 2px;
            }}
            QListWidget::item:hover:!selected {{
                background-color: {C.BG_SURFACE};
            }}
            QListWidget::item:selected {{
                background-color: {C.INFO_SOFT};
                color: {C.INFO};
                font-weight: bold;
                border: 1px solid {C.INFO_BORDER};
            }}
        """)
        self.list_topics.currentRowChanged.connect(self._on_topic_selected)
        left_layout.addWidget(self.list_topics)

        # Toolbar under topics list
        t_btn_layout = QGridLayout()
        t_btn_layout.setSpacing(4)
        self.btn_new = QPushButton("➕ New")
        self.btn_new.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_new.clicked.connect(self._on_new_topic)
        t_btn_layout.addWidget(self.btn_new, 0, 0)

        self.btn_edit = QPushButton("✏️ Edit")
        self.btn_edit.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_edit.clicked.connect(self._on_edit_topic)
        t_btn_layout.addWidget(self.btn_edit, 0, 1)

        self.btn_clone = QPushButton("📋 Clone")
        self.btn_clone.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_clone.clicked.connect(self._on_clone_topic)
        t_btn_layout.addWidget(self.btn_clone, 1, 0)

        self.btn_del = QPushButton("🗑️ Delete")
        self.btn_del.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_del.clicked.connect(self._on_delete_topic)
        t_btn_layout.addWidget(self.btn_del, 1, 1)
        left_layout.addLayout(t_btn_layout)

        self.btn_restore = QPushButton("🔄 Restore Built-in Topics...")
        self.btn_restore.setToolTip("Restore standard 3GPP topics (Ambient IoT, AI/ML Phase 2, 6G Architecture, Network Energy Savings, Data Framework) without overwriting custom topics.")
        self.btn_restore.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_restore.clicked.connect(self._on_restore_defaults)
        left_layout.addWidget(self.btn_restore)

        root_layout.addWidget(left_frame)

        # ==========================================
        # RIGHT PANE: Topic Runner & Options
        # ==========================================
        right_frame = QFrame()
        right_frame.setObjectName("CardFrame")
        right_frame.setStyleSheet(STYLE_CARD)
        right_layout = QVBoxLayout(right_frame)
        right_layout.setContentsMargins(16, 16, 16, 16)
        right_layout.setSpacing(12)

        # 1. Topic Header
        self.lbl_topic_title = QLabel("Select a Topic")
        self.lbl_topic_title.setWordWrap(True)
        self.lbl_topic_title.setStyleSheet(f"font-size: 16px; font-weight: bold; color: {C.TEXT_PRIMARY};")
        right_layout.addWidget(self.lbl_topic_title)

        self.lbl_topic_desc = QLabel("Description...")
        self.lbl_topic_desc.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px;")
        self.lbl_topic_desc.setWordWrap(True)
        right_layout.addWidget(self.lbl_topic_desc)

        # 2. Meeting Scope & Timeframe Picker
        cycle_box = QGroupBox("Meeting Scope && Timeframe")
        cycle_layout = QGridLayout(cycle_box)

        cycle_layout.addWidget(QLabel("<b>Meeting Timeframe:</b>"), 0, 0)
        self.combo_cycle_window = QComboBox()
        self.combo_cycle_window.addItems([
            "Past 4 Meeting Rounds (~1 Year)",
            "Past 6 Meeting Rounds (~1.5 Years)",
            "Past 8 Meeting Rounds (~2 Years)",
            "Release 20 Timeframe (2025–2027)",
            "Release 19 Timeframe (2023–2025)",
            "Release 18 Timeframe (2021–2023)",
            "📅 Custom Date Range...",
            "All Available Meetings",
        ])
        self.combo_cycle_window.currentIndexChanged.connect(self._on_timeframe_preset_changed)
        cycle_layout.addWidget(self.combo_cycle_window, 0, 1)

        # Date range picker container (shown when "📅 Custom Date Range..." is selected)
        self.date_range_widget = QWidget()
        dr_layout = QHBoxLayout(self.date_range_widget)
        dr_layout.setContentsMargins(0, 2, 0, 2)
        dr_layout.setSpacing(8)

        dr_layout.addWidget(QLabel("From:"))
        self.edit_date_from = QDateEdit()
        self.edit_date_from.setCalendarPopup(True)
        self.edit_date_from.setDisplayFormat("yyyy-MM-dd")
        self.edit_date_from.setDate(QDate.currentDate().addYears(-1))
        self.edit_date_from.dateChanged.connect(self._update_cohorts_preview)
        dr_layout.addWidget(self.edit_date_from)

        dr_layout.addWidget(QLabel("To:"))
        self.edit_date_to = QDateEdit()
        self.edit_date_to.setCalendarPopup(True)
        self.edit_date_to.setDisplayFormat("yyyy-MM-dd")
        self.edit_date_to.setDate(QDate.currentDate().addYears(1))
        self.edit_date_to.dateChanged.connect(self._update_cohorts_preview)
        dr_layout.addWidget(self.edit_date_to)
        dr_layout.addStretch()

        self.date_range_widget.setVisible(False)
        cycle_layout.addWidget(self.date_range_widget, 1, 1)

        self.lbl_cohorts_summary = QLabel("Covering: SA2, RAN2, CT1 across 4 meeting rounds")
        self.lbl_cohorts_summary.setWordWrap(True)
        self.lbl_cohorts_summary.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
        cycle_layout.addWidget(self.lbl_cohorts_summary, 2, 0, 1, 2)
        right_layout.addWidget(cycle_box)

        # 4. Target Directory & Incremental State
        dest_box = QGroupBox("Target Folder && Incremental Updating")
        dest_layout = QGridLayout(dest_box)

        dest_layout.addWidget(QLabel("<b>Destination:</b>"), 0, 0)
        self.edit_dest_folder = QLineEdit()
        self.edit_dest_folder.textChanged.connect(self._check_incremental_status)
        dest_layout.addWidget(self.edit_dest_folder, 0, 1)

        btn_browse = QPushButton("Browse...")
        btn_browse.setStyleSheet(STYLE_SECONDARY_BTN)
        btn_browse.clicked.connect(self._on_browse_dest)
        dest_layout.addWidget(btn_browse, 0, 2)

        self.lbl_inc_status = QLabel("🟢 New Folder")
        self.lbl_inc_status.setWordWrap(True)
        self.lbl_inc_status.setStyleSheet("font-size: 11px; font-weight: bold;")
        dest_layout.addWidget(self.lbl_inc_status, 1, 0, 1, 3)
        right_layout.addWidget(dest_box)

        # 4. Conversion Options (2x2 grid prevents excessive minimum dialog width)
        opt_box = QGroupBox("Dossier Content && Options")
        opt_layout = QGridLayout(opt_box)
        opt_layout.setHorizontalSpacing(12)
        opt_layout.setVerticalSpacing(6)
        self.chk_discussions = QCheckBox("Include Discussion Papers && WFs")
        self.chk_discussions.setChecked(True)
        self.chk_discussions.setToolTip(
            "Include non-agreed Discussion papers and Way Forwards alongside binding decisions and LSs "
            "(recommended for E2E company position tracking; uncheck for a lightweight Decisions/CRs/LSs baseline)."
        )
        opt_layout.addWidget(self.chk_discussions, 0, 0)

        self.chk_dedup = QCheckBox("Deduplicate Revisions")
        self.chk_dedup.setChecked(True)
        opt_layout.addWidget(self.chk_dedup, 0, 1)

        self.chk_consolidated = QCheckBox("Generate Consolidated Single .md")
        self.chk_consolidated.setChecked(True)
        opt_layout.addWidget(self.chk_consolidated, 1, 0)

        self.chk_download = QCheckBox("Download missing documents from 3GPP FTP")
        self.chk_download.setChecked(True)
        self.chk_download.setToolTip("Always downloads missing TDoc documents and meeting archives from 3GPP FTP to ensure complete dossier export.")
        opt_layout.addWidget(self.chk_download, 1, 1)
        right_layout.addWidget(opt_box)

        # 6. Progress & Actions
        self.prog_bar = QProgressBar()
        self.prog_bar.setVisible(False)
        self.prog_bar.setFixedHeight(12)
        right_layout.addWidget(self.prog_bar)

        self.lbl_status = QLabel("Ready")
        self.lbl_status.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px;")
        right_layout.addWidget(self.lbl_status)

        action_layout = QHBoxLayout()
        action_layout.setSpacing(8)

        actions_flow_widget = QWidget()
        flow_sp = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        flow_sp.setHeightForWidth(True)
        actions_flow_widget.setSizePolicy(flow_sp)
        actions_flow = FlowLayout(actions_flow_widget, margin=0, h_spacing=6, v_spacing=6)

        self.btn_view_inputs = QPushButton("👁️ View Considered Inputs...")
        self.btn_view_inputs.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_view_inputs.setToolTip("Open interactive visual inspector of all candidate decisions, pCRs, CRs, agreements, and liaison statements")
        self.btn_view_inputs.clicked.connect(self._on_view_considered_inputs)
        actions_flow.addWidget(self.btn_view_inputs)

        self.btn_check_status = QPushButton("🔍 Check Status && Refresh Index")
        self.btn_check_status.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_check_status.setEnabled(False)
        self.btn_check_status.setToolTip(
            "Fast offline scan of the destination folder to link _company_positions_analysis.md and "
            "company_positions_e2e_synthesis.md, flag stale deltas, and refresh index.md without re-exporting documents."
        )
        self.btn_check_status.clicked.connect(lambda: self._on_check_status(silent=False))
        actions_flow.addWidget(self.btn_check_status)

        self.btn_open_folder = QPushButton("📂 Open Output Folder")
        self.btn_open_folder.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_open_folder.setEnabled(False)
        self.btn_open_folder.clicked.connect(self._on_open_output)
        actions_flow.addWidget(self.btn_open_folder)

        self.btn_launch_cli = QPushButton("💬 Launch CLI Agent")
        self.btn_launch_cli.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_launch_cli.setEnabled(False)
        self.btn_launch_cli.setToolTip("Open Codex CLI, Claude CLI, or Gemini CLI directly inside the exported Cross-WG Topic Dossier.")
        self.cli_menu = QMenu(self)
        self.cli_menu.aboutToShow.connect(self._populate_cli_menu)
        self.btn_launch_cli.setMenu(self.cli_menu)
        actions_flow.addWidget(self.btn_launch_cli)

        action_layout.addWidget(actions_flow_widget, 1)

        self.btn_export = QPushButton("🚀 Export / Update Topic Dossier")
        self.btn_export.setStyleSheet(STYLE_PRIMARY_BTN)
        self.btn_export.clicked.connect(self._on_start_export)
        action_layout.addWidget(self.btn_export, 0, Qt.AlignBottom | Qt.AlignRight)

        right_layout.addLayout(action_layout)
        root_layout.addWidget(right_frame)
