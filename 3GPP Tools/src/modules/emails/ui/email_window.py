# --- File: modules/emails/ui/email_window.py ---
import datetime
import json
import os
import re
import webbrowser
from pathlib import Path
from PyQt5.QtCore import Qt, pyqtSignal, QDate
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QLineEdit, QTableView, QHeaderView, QSplitter, QTextBrowser,
                             QMessageBox, QInputDialog, QDialog, QDateEdit, QMenu, QApplication,
                             QAbstractItemView, QSizePolicy)

from modules.emails.core.email_db import EmailDatabase
from modules.emails.core.email_threads import EmailSyncThread, EmailMoveThread, EmailTargetRescanThread
from modules.emails.core.stats.statistics_threads import EmailStatsExporterThread
from modules.emails.ui.email_models import EmailTableModel, EmailProxyModel, TDocSummaryModel, TDocProxyModel
from modules.meetings.ui.tdocs_components import CheckableComboBox

class EmailManagerWindow(QWidget):
    tdoc_open_requested = pyqtSignal(str)

    def __init__(self, meeting_dir: Path, ai_lookup: dict, meeting_start: str = "", meeting_end: str = ""):
        super().__init__()
        self.meeting_dir = meeting_dir
        self.ai_lookup = ai_lookup

        self.db = EmailDatabase(self.meeting_dir / "Agenda" / "emails.db")
        self.config_path = self.meeting_dir / "Agenda" / "email_config.json"

        config_data = self._load_config()
        self.source_folder = config_data.get("source_folder", "")
        self.target_folder = config_data.get("target_folder", "")
        self.stats_config = config_data.get("stats_config", {})

        saved_start = config_data.get("start_date", "")
        saved_end = config_data.get("end_date", "")

        if saved_start and saved_end:
            self.start_date = saved_start
            self.end_date = saved_end
        else:
            try:
                if meeting_start and meeting_end:
                    s_dt = datetime.datetime.strptime(meeting_start, "%Y-%m-%d") - datetime.timedelta(days=3)
                    e_dt = datetime.datetime.strptime(meeting_end, "%Y-%m-%d") + datetime.timedelta(days=3)
                    self.start_date = s_dt.strftime("%Y-%m-%d")
                    self.end_date = e_dt.strftime("%Y-%m-%d")
                else:
                    self.start_date = ""
                    self.end_date = ""
            except Exception:
                self.start_date, self.end_date = "", ""

        self.setWindowTitle("📧 eMeeting Email Manager")
        self.resize(1300, 800)
        self.setStyleSheet("QWidget { background-color: #FAFAFA; }")

        self._setup_ui()
        self._refresh_table()

    def _load_config(self) -> dict:
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r') as f:
                    return json.load(f)
            except:
                pass
        return {
            "source_folder": "",
            "target_folder": "",
            "stats_config": {
                "email_top_companies": 25,
                "email_top_delegates": 25,
                "email_heatmap_top_comps": 25,
                "email_heatmap_top_ais": 25
            }
        }

    def _save_config(self, source_folder: str, target_folder: str, sd: str, ed: str, stats_cfg: dict = None):
        self.source_folder = source_folder
        self.target_folder = target_folder
        self.start_date = sd
        self.end_date = ed

        # Keep stats config in memory
        if stats_cfg is not None:
            self.stats_config = stats_cfg
        elif not hasattr(self, 'stats_config'):
            self.stats_config = self._load_config().get("stats_config", {})

        with open(self.config_path, 'w') as f:
            json.dump({
                "source_folder": source_folder,
                "target_folder": target_folder,
                "start_date": sd,
                "end_date": ed,
                "stats_config": self.stats_config
            }, f, indent=4)

    def _configure_folders(self):
        from modules.emails.ui.config_dialog import EmailConfigDialog

        current_stats = getattr(self, 'stats_config', self._load_config().get("stats_config", {}))

        dialog = EmailConfigDialog(self.source_folder, self.target_folder, current_stats, self)
        if dialog.exec_() == QDialog.Accepted:
            config_data = dialog.get_config_data()

            src = config_data["source_folder"]
            tgt = config_data["target_folder"]
            new_stats = config_data["stats_config"]

            self.source_folder = src
            self.target_folder = tgt
            self.stats_config = new_stats

            self._save_config(src, tgt, self.start_date, self.end_date, new_stats)

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)

        # ---> POLISH: Extremely tight outer margins to maximize grid space
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(4)

        def get_btn_style(primary=False):
            if primary:
                return """
                QPushButton { font-weight: bold; background-color: #0078D7; color: white; padding: 4px 8px; border-radius: 4px; border: 1px solid #005A9E; }
                QPushButton:hover { background-color: #005A9E; }
                """
            return """
                QPushButton { font-weight: bold; background-color: #FFFFFF; color: #333333; padding: 4px 8px; border-radius: 4px; border: 1px solid #CCCCCC; }
                QPushButton:hover { background-color: #F0F4F8; border-color: #005A9E; color: #005A9E; }
                """

        # =====================================================================
        # TOP CONTAINER: Ultra-Compressed Toolbars & Filters
        # =====================================================================
        top_controls_widget = QWidget()

        # ---> NEW: Force the top container to NEVER expand vertically past its exact minimum size
        top_controls_widget.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)

        top_controls_layout = QVBoxLayout(top_controls_widget)
        top_controls_layout.setContentsMargins(0, 0, 0, 2)  # Barely any bottom margin
        top_controls_layout.setSpacing(2)  # Squish the two rows together

        # --- TOOLBAR ---
        toolbar = QHBoxLayout()
        toolbar.setContentsMargins(0, 0, 0, 0)
        toolbar.setSpacing(6)

        self.btn_sync = QPushButton("🔄 Sync Source")
        self.btn_sync.setStyleSheet(get_btn_style(primary=True))
        self.btn_sync.clicked.connect(self._run_sync)

        self.btn_move = QPushButton("➡️ Move Selected")
        self.btn_move.setStyleSheet(get_btn_style())
        self.btn_move.clicked.connect(self._run_move)

        self.btn_move_all = QPushButton("⏭️ Move All")
        self.btn_move_all.setStyleSheet(get_btn_style())
        self.btn_move_all.clicked.connect(self._run_move_all)

        self.btn_rescan = QPushButton("🔁 Scan Target")
        self.btn_rescan.setStyleSheet(get_btn_style())
        self.btn_rescan.clicked.connect(self._run_target_rescan)

        self.btn_stats = QPushButton("📊 Statistics")
        self.btn_stats.setStyleSheet(get_btn_style())
        self.btn_stats.clicked.connect(self._generate_statistics)

        self.btn_config = QPushButton("⚙️")
        self.btn_config.setStyleSheet(get_btn_style())
        self.btn_config.clicked.connect(self._configure_folders)

        self.lbl_status = QLabel("Ready.")
        self.lbl_status.setStyleSheet("color: #666; font-style: italic; margin-left: 10px;")

        for btn in [self.btn_sync, self.btn_move, self.btn_move_all, self.btn_rescan, self.btn_stats, self.btn_config]:
            toolbar.addWidget(btn)
        toolbar.addWidget(self.lbl_status)
        toolbar.addStretch()
        top_controls_layout.addLayout(toolbar)

        # --- GLOBAL FILTERS ---
        filter_layout = QHBoxLayout()
        filter_layout.setContentsMargins(0, 0, 0, 0)
        filter_layout.setSpacing(6)

        self.dt_start = QDateEdit()
        self.dt_start.setCalendarPopup(True)
        if getattr(self, 'start_date', None):
            self.dt_start.setDate(QDate.fromString(self.start_date, Qt.ISODate))
        else:
            self.dt_start.setDate(QDate.currentDate().addDays(-7))

        self.dt_end = QDateEdit()
        self.dt_end.setCalendarPopup(True)
        if getattr(self, 'end_date', None):
            self.dt_end.setDate(QDate.fromString(self.end_date, Qt.ISODate))
        else:
            self.dt_end.setDate(QDate.currentDate().addDays(7))

        self.dt_start.dateChanged.connect(lambda d: setattr(self, 'start_date', d.toString(Qt.ISODate)))
        self.dt_end.dateChanged.connect(lambda d: setattr(self, 'end_date', d.toString(Qt.ISODate)))
        self.dt_start.dateChanged.connect(
            lambda: self._save_config(self.source_folder, self.target_folder, self.start_date, self.end_date))
        self.dt_end.dateChanged.connect(
            lambda: self._save_config(self.source_folder, self.target_folder, self.start_date, self.end_date))

        self.cb_ai = CheckableComboBox("Filter by AI")
        self.cb_ai.setMinimumWidth(150)
        self.cb_ai.selectionChanged.connect(self._apply_filters)

        self.btn_filter_star = QPushButton("⭐ Starred")
        self.btn_filter_star.setCheckable(True)
        # ---> POLISH: Compressed padding to make filter buttons tighter
        self.btn_filter_star.setStyleSheet(
            "QPushButton { padding: 3px 6px; border: 1px solid #CCC; background: white; border-radius: 3px; } QPushButton:checked { background-color: #FFF4CE; font-weight: bold; border-color: #E2C08D; }")
        self.btn_filter_star.clicked.connect(self._apply_filters)

        self.btn_filter_follow = QPushButton("👀 Followed AIs")
        self.btn_filter_follow.setCheckable(True)
        # ---> POLISH: Compressed padding to make filter buttons tighter
        self.btn_filter_follow.setStyleSheet(
            "QPushButton { padding: 3px 6px; border: 1px solid #CCC; background: white; border-radius: 3px; } QPushButton:checked { background-color: #E6F4E6; font-weight: bold; color: #0C6B0C; border-color: #0C6B0C; }")
        self.btn_filter_follow.clicked.connect(self._apply_filters)

        self.lbl_count = QLabel("Showing 0 Threads | 0 Emails")
        self.lbl_count.setStyleSheet("font-size: 13px; color: #555; font-weight: bold; padding-left: 10px;")

        filter_layout.addWidget(QLabel("📅 Filter:"))
        filter_layout.addWidget(self.dt_start)
        filter_layout.addWidget(QLabel("-"))
        filter_layout.addWidget(self.dt_end)
        filter_layout.addSpacing(10)
        filter_layout.addWidget(self.btn_filter_star)
        filter_layout.addWidget(self.btn_filter_follow)
        filter_layout.addWidget(self.cb_ai)
        filter_layout.addStretch()
        filter_layout.addWidget(self.lbl_count)
        top_controls_layout.addLayout(filter_layout)

        main_layout.addWidget(top_controls_widget)

        # =====================================================================
        # MASTER-DETAIL SPLITTER
        # =====================================================================
        self.main_splitter = QSplitter(Qt.Horizontal)

        # --- LEFT PANEL: TDOC THREADS ---
        self.left_panel = QWidget()
        self.left_panel.setMinimumWidth(350)

        left_layout = QVBoxLayout(self.left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)

        left_header_layout = QHBoxLayout()
        self.lbl_left_title = QLabel("📄 Active Threads")
        self.lbl_left_title.setStyleSheet("font-weight: bold; color: #555;")

        self.tdoc_search_input = QLineEdit()
        self.tdoc_search_input.setPlaceholderText("🔍 Search threads...")
        self.tdoc_search_input.textChanged.connect(self._apply_filters)

        left_header_layout.addWidget(self.lbl_left_title)
        left_header_layout.addStretch()
        left_header_layout.addWidget(self.tdoc_search_input)
        left_layout.addLayout(left_header_layout)

        self.tdoc_view = QTableView()
        self.tdoc_view.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tdoc_view.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tdoc_view.verticalHeader().setVisible(False)
        self.tdoc_view.setAlternatingRowColors(True)
        self.tdoc_view.setStyleSheet("QTableView { background: white; gridline-color: #EEE; }")

        left_layout.addWidget(self.tdoc_view)

        # --- RIGHT PANEL: EMAILS ---
        self.right_panel = QWidget()
        right_layout = QVBoxLayout(self.right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)

        right_header_layout = QHBoxLayout()
        self.lbl_right_title = QLabel("📧 Thread History (Select a TDoc from the left)")
        self.lbl_right_title.setStyleSheet("font-weight: bold; color: #0078D7;")

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 Search this thread...")
        self.search_input.setMaximumWidth(200)
        self.search_input.textChanged.connect(self._apply_filters)

        self.cb_company = CheckableComboBox("Company")
        self.cb_company.setMinimumWidth(110)
        self.cb_company.selectionChanged.connect(self._apply_filters)

        self.cb_sender = CheckableComboBox("Sender")
        self.cb_sender.setMinimumWidth(110)
        self.cb_sender.selectionChanged.connect(self._apply_filters)

        right_header_layout.addWidget(self.lbl_right_title)
        right_header_layout.addStretch()
        right_header_layout.addWidget(self.cb_company)
        right_header_layout.addWidget(self.cb_sender)
        right_header_layout.addWidget(self.search_input)
        right_layout.addLayout(right_header_layout)

        right_splitter = QSplitter(Qt.Vertical)

        self.email_view = QTableView()
        self.email_view.setMinimumHeight(200)
        self.email_view.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.email_view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.email_view.verticalHeader().setVisible(False)
        self.email_view.setStyleSheet("QTableView { background: white; gridline-color: #EEE; }")
        self.email_view.setCursor(Qt.PointingHandCursor)
        self.email_view.clicked.connect(self._on_table_clicked)
        self.email_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.email_view.customContextMenuRequested.connect(self._show_context_menu)
        right_splitter.addWidget(self.email_view)

        # Reading Pane
        pane_widget = QWidget()
        pane_widget.setMinimumHeight(150)

        pane_layout = QVBoxLayout(pane_widget)
        pane_layout.setContentsMargins(0, 10, 0, 0)

        self.btn_open_msg = QPushButton("📄 Open .msg File")
        self.btn_open_msg.setEnabled(False)
        self.btn_open_msg.clicked.connect(self._open_current_msg)

        self.btn_toggle_star = QPushButton("⭐ Star TDoc")
        self.btn_toggle_star.setEnabled(False)
        self.btn_toggle_star.clicked.connect(self._toggle_current_star)

        self.btn_toggle_follow = QPushButton("👀 Follow AI")
        self.btn_toggle_follow.setEnabled(False)
        self.btn_toggle_follow.clicked.connect(self._toggle_current_ai_follow)

        btn_layout = QHBoxLayout()
        btn_layout.addWidget(self.btn_open_msg)
        btn_layout.addWidget(self.btn_toggle_star)
        btn_layout.addWidget(self.btn_toggle_follow)
        btn_layout.addStretch()

        self.reading_pane = QTextBrowser()
        self.reading_pane.setStyleSheet("background: white; border: 1px solid #CCC; border-radius: 4px; padding: 10px;")
        pane_layout.addLayout(btn_layout)
        pane_layout.addWidget(self.reading_pane)

        right_splitter.addWidget(pane_widget)
        right_splitter.setStretchFactor(0, 3)
        right_splitter.setStretchFactor(1, 1)
        right_splitter.setSizes([700, 250])

        right_layout.addWidget(right_splitter)

        self.main_splitter.addWidget(self.left_panel)
        self.main_splitter.addWidget(self.right_panel)
        self.main_splitter.setStretchFactor(0, 1)
        self.main_splitter.setStretchFactor(1, 3)
        self.main_splitter.setSizes([350, 1000])
        main_layout.addWidget(self.main_splitter)

        # =====================================================================
        # MODEL SETUP
        # =====================================================================
        self.tdoc_model = TDocSummaryModel()
        self.tdoc_proxy = TDocProxyModel()
        self.tdoc_proxy.setSourceModel(self.tdoc_model)
        self.tdoc_view.setModel(self.tdoc_proxy)

        self.email_model = EmailTableModel()
        self.email_proxy = EmailProxyModel()
        self.email_proxy.setSourceModel(self.email_model)
        self.email_view.setModel(self.email_proxy)

        self.tdoc_view.selectionModel().selectionChanged.connect(self._on_tdoc_selected)
        self.email_view.selectionModel().selectionChanged.connect(self._on_email_selected)

        self.tdoc_proxy.layoutChanged.connect(self._update_count_label)
        self.email_proxy.layoutChanged.connect(self._update_count_label)

        # Sizing Headers
        tdoc_header = self.tdoc_view.horizontalHeader()
        tdoc_header.setSectionResizeMode(QHeaderView.Interactive)
        tdoc_header.resizeSection(0, 25)
        tdoc_header.setSectionResizeMode(1, QHeaderView.Stretch)
        tdoc_header.resizeSection(2, 60)
        tdoc_header.resizeSection(3, 45)

        email_header = self.email_view.horizontalHeader()
        email_header.setSectionResizeMode(QHeaderView.Interactive)
        email_header.resizeSection(0, 30)
        email_header.resizeSection(1, 85)
        email_header.resizeSection(2, 75)
        email_header.resizeSection(3, 120)
        email_header.resizeSection(4, 90)
        email_header.resizeSection(5, 60)
        email_header.resizeSection(6, 70)
        email_header.resizeSection(7, 110)
        email_header.resizeSection(8, 140)
        email_header.setSectionResizeMode(9, QHeaderView.Stretch)

    # -------------------------------------------------------------------------
    # CORE LOGIC & EVENT HANDLERS
    # -------------------------------------------------------------------------
    def _refresh_table(self):
        old_tdoc = self.email_proxy.target_tdoc

        old_email_id = None
        indexes = self.email_view.selectionModel().selectedRows()
        if indexes:
            source_idx = self.email_proxy.mapToSource(indexes[0])
            old_email_id = self.email_model.get_row_data(source_idx.row()).get("id")

        import sqlite3
        with sqlite3.connect(self.db.db_path) as conn:
            conn.row_factory = sqlite3.Row
            data = [dict(row) for row in conn.execute('SELECT * FROM emails ORDER BY date_received DESC').fetchall()]

        starred_tdocs = self.db.get_starred_tdocs()
        followed_ais = self.db.get_followed_ais()

        self.tdoc_model.update_data(data, starred_tdocs, followed_ais)
        self.email_model.update_data(data, starred_tdocs, followed_ais)

        def clean(val):
            return str(val).strip() if val else ""

        def natural_sort_key(s):
            return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

        unique_ais = set(clean(r.get("agenda_item")) for r in data)
        unique_companies = set(clean(r.get("company")) for r in data)
        unique_senders = set(clean(r.get("sender_name")) for r in data)

        self.cb_ai.updateItems(sorted(unique_ais, key=natural_sort_key))
        self.cb_company.updateItems(sorted(unique_companies))
        self.cb_sender.updateItems(sorted(unique_senders))

        self._apply_filters()

        if old_tdoc:
            for r in range(self.tdoc_proxy.rowCount()):
                idx = self.tdoc_proxy.index(r, 0)
                if self.tdoc_proxy.data(idx, Qt.UserRole) == old_tdoc:
                    self.tdoc_view.selectRow(r)
                    break

        if old_email_id:
            for r in range(self.email_proxy.rowCount()):
                idx = self.email_proxy.index(r, 0)
                src_idx = self.email_proxy.mapToSource(idx)
                if self.email_model.get_row_data(src_idx.row()).get("id") == old_email_id:
                    self.email_view.selectRow(r)
                    break

    def _apply_filters(self):
        self.tdoc_proxy.set_filters(
            self.btn_filter_star.isChecked(),
            self.btn_filter_follow.isChecked(),
            self.cb_ai.getCheckedItems(),
            self.tdoc_search_input.text()
        )
        self.email_proxy.set_filters(
            self.search_input.text(),
            self.cb_company.getCheckedItems(),
            self.cb_sender.getCheckedItems()
        )
        self._update_count_label()

    def _on_tdoc_selected(self, selected, deselected):
        indexes = self.tdoc_view.selectionModel().selectedRows()
        if not indexes:
            self.email_proxy.set_target_tdoc(None)
            self.lbl_right_title.setText("📧 Thread History (Select a TDoc from the left)")
            return

        proxy_index = indexes[0]
        tdoc_id = self.tdoc_proxy.data(proxy_index, Qt.UserRole)

        self.email_proxy.set_target_tdoc(tdoc_id)

        count = self.email_proxy.rowCount()
        self.lbl_right_title.setText(f"📧 Thread History: {tdoc_id} ({count} Emails)")
        self._update_count_label()

    def _on_email_selected(self, selected, deselected):
        indexes = self.email_view.selectionModel().selectedRows()
        if not indexes:
            self.reading_pane.clear()
            self.btn_open_msg.setEnabled(False)
            self.btn_toggle_star.setEnabled(False)
            self.btn_toggle_follow.setEnabled(False)
            return

        source_idx = self.email_proxy.mapToSource(indexes[0])
        row_data = self.email_model.get_row_data(source_idx.row())

        self.current_tdoc_id = row_data.get("tdoc_id", "")
        self.current_agenda_item = row_data.get("agenda_item", "")
        self.current_msg_path = row_data.get("msg_path", "")

        self.btn_open_msg.setEnabled(bool(self.current_msg_path))
        self.btn_toggle_star.setEnabled(bool(self.current_tdoc_id))
        self.btn_toggle_follow.setEnabled(bool(self.current_agenda_item))

        is_starred = self.current_tdoc_id in self.email_model.starred_tdocs
        self.btn_toggle_star.setText(
            f"❌ Unstar {self.current_tdoc_id}" if is_starred else f"⭐ Star {self.current_tdoc_id}")

        is_followed = self.current_agenda_item in self.email_model.followed_ais
        ai_display = self.current_agenda_item or "Unknown AI"
        self.btn_toggle_follow.setText(f"❌ Unfollow {ai_display}" if is_followed else f"👀 Follow {ai_display}")

        real_email = row_data.get('sender_email', '')
        html = f"""
        <h2 style='color:#005A9E; margin-bottom: 2px;'>{row_data.get('subject', '')}</h2>
        <p style='color:#555; margin-top:0px;'><b>From:</b> {row_data.get('sender_name')} &lt;{real_email}&gt; ({row_data.get('company')}) | <b>TDoc:</b> {row_data.get('tdoc_id')} | <b>Rev:</b> {row_data.get('revisions_mentioned')}</p>
        <hr>
        <h3 style='color:#D83B01;'>Short Text / Comments</h3>
        <p style='background-color:#FFF3E0; padding:10px; border-left: 4px solid #FFB74D;'>{row_data.get('short_text', '').replace(chr(10), '<br>')}</p>
        <h3>Free Text</h3>
        <p style='color:#333;'>{row_data.get('free_text', '').replace(chr(10), '<br>')}</p>
        """
        self.reading_pane.setHtml(html)

    def _update_count_label(self):
        self.lbl_count.setText(f"Showing {self.tdoc_proxy.rowCount()} Threads | {self.email_proxy.rowCount()} Emails")

    # -------------------------------------------------------------------------
    # INTERACTIONS (Clicks, Syncs, Stars)
    # -------------------------------------------------------------------------
    def _on_table_clicked(self, index):
        if not index.isValid(): return
        source_idx = self.email_proxy.mapToSource(index)
        col_name = self.email_model._headers[index.column()]
        row_data = self.email_model.get_row_data(source_idx.row())

        if col_name == "Sender":
            email = row_data.get("sender_email", "")
            if email: webbrowser.open(f"mailto:{email}")

        elif col_name == "Rev":
            revs = row_data.get("revisions_mentioned", "")
            if revs:
                first_rev = revs.split(',')[0].strip()
                self.tdoc_open_requested.emit(first_rev)
                self.lbl_status.setText(f"Requesting open for: {first_rev}")

    def _show_context_menu(self, pos):
        index = self.email_view.indexAt(pos)
        if not index.isValid(): return
        col_name = self.email_model._headers[index.column()]
        if col_name == "Sender":
            row_data = self.email_model.get_row_data(self.email_proxy.mapToSource(index).row())
            email = row_data.get("sender_email", "")
            menu = QMenu(self)
            copy_action = menu.addAction("📋 Copy Email Address")
            action = menu.exec_(self.email_view.viewport().mapToGlobal(pos))
            if action == copy_action and email:
                QApplication.clipboard().setText(email)
                self.lbl_status.setText(f"Copied {email} to clipboard.")

    def _toggle_current_star(self):
        if not getattr(self, "current_tdoc_id", ""): return
        is_starred = self.current_tdoc_id in self.email_model.starred_tdocs
        new_status = not is_starred
        self.btn_toggle_star.setText(
            f"❌ Unstar {self.current_tdoc_id}" if new_status else f"⭐ Star {self.current_tdoc_id}")

        self.db.toggle_tdoc_star(self.current_tdoc_id, new_status)
        self._refresh_table()

    def _toggle_current_ai_follow(self):
        if not getattr(self, "current_agenda_item", ""): return
        is_followed = self.current_agenda_item in self.email_model.followed_ais
        new_status = not is_followed
        ai_display = self.current_agenda_item or "Unknown AI"
        self.btn_toggle_follow.setText(f"❌ Unfollow {ai_display}" if new_status else f"👀 Follow {ai_display}")

        self.db.toggle_ai_follow(self.current_agenda_item, new_status)
        self._refresh_table()

    def _open_current_msg(self):
        if self.current_msg_path and Path(self.current_msg_path).exists():
            try:
                os.startfile(self.current_msg_path)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not open file: {e}")

    # -------------------------------------------------------------------------
    # THREADING (Sync, Move, Scan, Stats)
    # -------------------------------------------------------------------------
    def _run_sync(self):
        if not self.source_folder:
            QMessageBox.warning(self, "Setup Required", "Please configure the Outlook Source Folder first.")
            return
        self._set_buttons_enabled(False)
        self.btn_sync.setText("⏳ Syncing...")
        sd = self.dt_start.date().toString(Qt.ISODate)
        ed = self.dt_end.date().toString(Qt.ISODate)
        self.sync_thread = EmailSyncThread(self.source_folder, self.meeting_dir, self.ai_lookup, self.db, sd, ed)
        self.sync_thread.log_msg.connect(lambda m, _: self.lbl_status.setText(m))
        self.sync_thread.progress_update.connect(
            lambda c, t: self.lbl_status.setText(f"Scanning Source: {c} / {t} items..."))
        self.sync_thread.finished.connect(self._on_sync_finished)
        self.sync_thread.start()

    def _on_sync_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)
        self.btn_sync.setText("🔄 Sync Source")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _run_move(self):
        if not self.target_folder:
            QMessageBox.warning(self, "Setup Required", "Please configure the Outlook Target Folder first.")
            return
        indexes = self.email_view.selectionModel().selectedRows()
        if not indexes:
            QMessageBox.information(self, "Selection Required", "Please select at least one email to move.")
            return
        items_to_move = []
        for idx in indexes:
            source_idx = self.email_proxy.mapToSource(idx)
            row_data = self.email_model.get_row_data(source_idx.row())
            if row_data.get("outlook_location") == "Source":
                items_to_move.append((row_data.get("id"), row_data.get("agenda_item")))
        if not items_to_move:
            QMessageBox.information(self, "Notice", "Selected emails have already been moved to the Target.")
            return
        self._set_buttons_enabled(False)
        self.btn_move.setText("⏳ Moving...")
        self.move_thread = EmailMoveThread(items_to_move, self.target_folder, self.db)
        self.move_thread.progress_update.connect(lambda c, t: self.lbl_status.setText(f"Moving {c}/{t}..."))
        self.move_thread.finished.connect(self._on_move_finished)
        self.move_thread.start()

    def _run_move_all(self):
        if not self.target_folder:
            QMessageBox.warning(self, "Setup Required", "Please configure the Outlook Target Folder first.")
            return
        items_to_move = []
        for row_data in self.email_model._data:
            if row_data.get("outlook_location") == "Source":
                items_to_move.append((row_data.get("id"), row_data.get("agenda_item")))
        if not items_to_move:
            QMessageBox.information(self, "Notice", "There are no emails in the Source folder to move.")
            return
        reply = QMessageBox.question(self, 'Confirm Move All',
                                     f"Move ALL {len(items_to_move)} source emails to the target?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.No: return
        self._set_buttons_enabled(False)
        self.btn_move_all.setText("⏳ Moving All...")
        self.move_thread = EmailMoveThread(items_to_move, self.target_folder, self.db)
        self.move_thread.progress_update.connect(lambda c, t: self.lbl_status.setText(f"Moving {c}/{t}..."))
        self.move_thread.finished.connect(self._on_move_all_finished)
        self.move_thread.start()

    def _on_move_all_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)
        self.btn_move_all.setText("⏭️ Move All")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _on_move_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)
        self.btn_move.setText("➡️ Move Selected")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _run_target_rescan(self):
        if not self.target_folder:
            QMessageBox.warning(self, "Setup Required", "Please configure the Outlook Target Folder first.")
            return
        self._set_buttons_enabled(False)
        self.btn_rescan.setText("⏳ Scanning...")
        sd = self.dt_start.date().toString(Qt.ISODate)
        ed = self.dt_end.date().toString(Qt.ISODate)
        self.rescan_thread = EmailTargetRescanThread(self.target_folder, self.meeting_dir, self.ai_lookup, self.db, sd,
                                                     ed)
        self.rescan_thread.log_msg.connect(lambda m, _: self.lbl_status.setText(m))
        self.rescan_thread.progress_update.connect(
            lambda c, t: self.lbl_status.setText(f"Scanning Target: {c} / {t} items..."))
        self.rescan_thread.finished.connect(self._on_rescan_finished)
        self.rescan_thread.start()

    def _on_rescan_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)
        self.btn_rescan.setText("🔁 Scan Target")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _generate_statistics(self):
        all_emails = self.email_model._data
        if not all_emails:
            QMessageBox.warning(self, "No Data", "There are no emails loaded to analyze.")
            return

        self._set_buttons_enabled(False)
        self.btn_stats.setText("⏳ Generating...")

        # Load the stats parameters saved from the configuration dialog
        current_stats = getattr(self, 'stats_config', self._load_config().get("stats_config", {}))

        # Import the new Modularized Exporter Thread
        meeting_name = self.meeting_dir.name if self.meeting_dir else "Meeting"

        self.stats_thread = EmailStatsExporterThread(self.meeting_dir, all_emails, meeting_name, current_stats)
        self.stats_thread.finished.connect(self._on_stats_finished)
        self.stats_thread.start()

    def _on_stats_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)
        self.btn_stats.setText("📊 Statistics")
        if success:
            self.lbl_status.setText("✅ Analytics Report Generated.")
            if hasattr(os, 'startfile'):
                os.startfile(msg)
            else:
                webbrowser.open(f"file:///{msg}")
        else:
            self.lbl_status.setText("❌ Analytics Generation Failed.")
            QMessageBox.warning(self, "Error", f"Could not generate statistics:\n{msg}")

    def _set_buttons_enabled(self, state: bool):
        self.btn_sync.setEnabled(state)
        self.btn_move.setEnabled(state)
        self.btn_move_all.setEnabled(state)
        self.btn_rescan.setEnabled(state)
        self.btn_stats.setEnabled(state)