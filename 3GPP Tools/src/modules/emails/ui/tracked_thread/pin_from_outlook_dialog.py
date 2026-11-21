"""Outlook folder discovery and pinning dialog for untracked meeting email threads."""

from datetime import datetime, timedelta
from pathlib import Path
import re
from typing import Dict, List, Optional

from PyQt5.QtCore import QSettings, Qt
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from modules.emails.core.tracked_threads_service import (
    classify_discovered_threads,
    load_meeting_discovery_config,
    pin_thread_by_subject_or_cid,
    save_meeting_discovery_config,
)
from modules.emails.core.tracked_threads_workers import OutlookFolderScanWorker
from modules.emails.ui.sync_progress_bar import AsyncSyncProgressBar
from core.ui.theme import C


def _infer_meeting_tag(meeting_id: str, wg: str) -> str:
    raw = str(meeting_id or "").strip()
    if not raw:
        return str(wg or "").strip()
    if "#" in raw:
        return raw
    m = re.match(r"^([A-Za-z0-9]+)\s+(\d+.*)$", raw)
    if m:
        return f"{m.group(1)}#{m.group(2)}"
    if wg and raw.isdigit():
        return f"{wg}#{raw}"
    return raw


class PinFromOutlookDialog(QDialog):
    """Browse, classify, ignore, and pin untracked email threads from configured WG Outlook folders."""

    def __init__(
        self,
        db: TrackedThreadsDatabase,
        meeting_id: str,
        wg: str,
        parent: Optional[QWidget] = None,
        agenda_map: Optional[Dict[str, object]] = None,
        valid_wis: Optional[List[str]] = None,
        meeting_dir: Optional[Path] = None,
        general_email_db=None,
        valid_tdocs: Optional[set] = None,
        mtg_start_date: str = "",
        mtg_end_date: str = "",
        auto_scan: bool = True,
        tdoc_meta_map: Optional[Dict[str, dict]] = None,
    ):
        super().__init__(parent)
        self.db = db
        self.meeting_id = meeting_id
        self.wg = wg
        self._agenda_map = agenda_map
        self._valid_wis = valid_wis
        self._meeting_dir = Path(meeting_dir) if meeting_dir else None
        self._general_email_db = general_email_db
        self._valid_tdocs = set(valid_tdocs) if valid_tdocs else set()
        self._tdoc_meta_map = tdoc_meta_map or {}
        self.pinned_thread_ids: List[str] = []
        self._raw_threads: List[dict] = []
        self._threads: List[dict] = []
        self._scan_worker: Optional[OutlookFolderScanWorker] = None
        self._loading_config = False

        default_tag = _infer_meeting_tag(meeting_id, wg)
        self._config = load_meeting_discovery_config(
            self._meeting_dir,
            default_tag=default_tag,
            default_start=str(mtg_start_date or "").strip()[:10],
            default_end=str(mtg_end_date or "").strip()[:10],
        )

        self.setWindowFlags(Qt.Window)
        self.setSizeGripEnabled(True)
        self.setWindowTitle(f"Discover & Pin Email Threads — Configured {wg} Outlook Folders")
        self.setMinimumSize(620, 420)
        self.resize(1120, 680)
        self._settings = QSettings("3GPP_Tools", "PinFromOutlookDialog")
        saved_geom = self._settings.value("geometry")
        if saved_geom:
            try:
                self.restoreGeometry(saved_geom)
            except Exception:
                pass
        self._build_ui()
        if auto_scan:
            self._load_threads()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)

        # Row 1: Responsive FlowLayout for Coverage Filter, Search & Rescan
        top_widget = QWidget(self)
        top_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        top_sp.setHeightForWidth(True)
        top_widget.setSizePolicy(top_sp)
        top_row = FlowLayout(top_widget, margin=0, h_spacing=8, v_spacing=4)
        top_row.addWidget(QLabel("<b>Coverage:</b>"))
        self.coverage_combo = QComboBox()
        self.coverage_combo.addItem("🔍 Untracked & Body-Only (🔴 + 🟡)", "UNTRACKED_AND_BODY")
        self.coverage_combo.addItem("🔴 Strictly Untracked Only (No TDoc)", "STRICT_UNTRACKED")
        self.coverage_combo.addItem("🟡 All Unpinned (Includes Subject TDoc)", "ALL_UNPINNED")
        self.coverage_combo.addItem("📋 All Discovered (Includes Pinned)", "ALL")
        self.coverage_combo.addItem("🚫 Ignored Threads", "IGNORED")
        cov_val = str(self._config.get("coverage_filter") or "UNTRACKED_AND_BODY").upper()
        cov_idx = max(0, self.coverage_combo.findData(cov_val))
        self.coverage_combo.setCurrentIndex(cov_idx)
        self.coverage_combo.currentIndexChanged.connect(self._on_config_ui_changed)
        top_row.addWidget(self.coverage_combo)

        self.search_edit = QLineEdit()
        self.search_edit.setMinimumWidth(220)
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.setPlaceholderText("🔍 Search subject, sender, TDoc, AI, or WI...")
        self.search_edit.textChanged.connect(self._apply_filter)
        top_row.addWidget(self.search_edit)

        self.reload_btn = QPushButton("🔄 Rescan Folders")
        self.reload_btn.clicked.connect(self._load_threads)
        top_row.addWidget(self.reload_btn)
        layout.addWidget(top_widget)

        # Row 2: Configurable Subject Tag & Meeting Date Window (persisted per meeting)
        cfg_frame = QFrame(self)
        cfg_frame.setObjectName("DiscoveryCfgBar")
        cfg_frame.setStyleSheet(
            f"#DiscoveryCfgBar {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; }}"
        )
        cfg_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        cfg_sp.setHeightForWidth(True)
        cfg_frame.setSizePolicy(cfg_sp)
        cfg_row = FlowLayout(cfg_frame, margin=8, h_spacing=8, v_spacing=4)

        self.require_tag_check = QCheckBox("Require Subject Tag:")
        self.require_tag_check.setChecked(bool(self._config.get("require_subject_tag", False)))
        self.require_tag_check.setToolTip(
            "Only show threads whose Outlook subject contains this substring (configurable & saved per meeting)."
        )
        self.require_tag_check.toggled.connect(self._on_config_ui_changed)
        cfg_row.addWidget(self.require_tag_check)

        self.tag_edit = QLineEdit(str(self._config.get("subject_tag") or ""))
        self.tag_edit.setPlaceholderText("e.g. SA2#178")
        self.tag_edit.setMaximumWidth(150)
        self.tag_edit.setClearButtonEnabled(True)
        self.tag_edit.textChanged.connect(self._on_config_ui_changed)
        cfg_row.addWidget(self.tag_edit)

        self.use_date_check = QCheckBox("Date Window:")
        self.use_date_check.setChecked(bool(self._config.get("use_date_filter", False)))
        self.use_date_check.setToolTip("Restrict results to the meeting date window ± buffer days.")
        self.use_date_check.toggled.connect(self._on_config_ui_changed)
        cfg_row.addWidget(self.use_date_check)

        self.start_date_edit = QLineEdit(str(self._config.get("start_date") or ""))
        self.start_date_edit.setPlaceholderText("YYYY-MM-DD")
        self.start_date_edit.setMaximumWidth(95)
        self.start_date_edit.textChanged.connect(self._on_config_ui_changed)
        cfg_row.addWidget(self.start_date_edit)

        cfg_row.addWidget(QLabel("to"))

        self.end_date_edit = QLineEdit(str(self._config.get("end_date") or ""))
        self.end_date_edit.setPlaceholderText("YYYY-MM-DD")
        self.end_date_edit.setMaximumWidth(95)
        self.end_date_edit.textChanged.connect(self._on_config_ui_changed)
        cfg_row.addWidget(self.end_date_edit)

        cfg_row.addWidget(QLabel("−"))
        self.days_before_spin = QSpinBox()
        self.days_before_spin.setRange(0, 90)
        self.days_before_spin.setValue(int(self._config.get("days_before", 14) or 0))
        self.days_before_spin.setSuffix("d")
        self.days_before_spin.setToolTip("Days before start date to include.")
        self.days_before_spin.valueChanged.connect(self._on_config_ui_changed)
        cfg_row.addWidget(self.days_before_spin)

        cfg_row.addWidget(QLabel("+"))
        self.days_after_spin = QSpinBox()
        self.days_after_spin.setRange(0, 90)
        self.days_after_spin.setValue(int(self._config.get("days_after", 7) or 0))
        self.days_after_spin.setSuffix("d")
        self.days_after_spin.setToolTip("Days after end date to include.")
        self.days_after_spin.valueChanged.connect(self._on_config_ui_changed)
        cfg_row.addWidget(self.days_after_spin)

        layout.addWidget(cfg_frame)

        self.scan_bar = AsyncSyncProgressBar(self)
        self.scan_bar.cancel_requested.connect(self._on_cancel_scan)
        layout.addWidget(self.scan_bar)

        self.summary_label = QLabel("")
        self.summary_label.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
        layout.addWidget(self.summary_label)

        # Multi-column discovery table
        self.table = QTableWidget(0, 6, self)
        self.table.setHorizontalHeaderLabels([
            "Coverage",
            "Msgs (Unread)",
            "Subject / Thread Title",
            "Suggested AI / WI / TDocs",
            "Last Sender",
            "Last Reply",
        ])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.Interactive)
        hdr.resizeSection(0, 145)
        hdr.resizeSection(1, 95)
        hdr.setSectionResizeMode(2, QHeaderView.Stretch)
        hdr.resizeSection(3, 195)
        hdr.resizeSection(4, 165)
        hdr.resizeSection(5, 120)
        self.table.itemSelectionChanged.connect(self._update_action_buttons)
        layout.addWidget(self.table, 1)

        btn_row = QHBoxLayout()
        self.ignore_btn = QPushButton("🚫 Ignore Selected")
        self.ignore_btn.setToolTip("Hide selected thread(s) from future untracked discovery for this meeting.")
        self.ignore_btn.clicked.connect(self._on_toggle_ignore_selected)
        btn_row.addWidget(self.ignore_btn)

        btn_row.addStretch()
        self.pin_btn = QPushButton("📌 Pin Selected Thread(s)")
        self.pin_btn.setStyleSheet(
            f"QPushButton {{ background-color: {C.ACCENT_ALT}; color: white; font-weight: bold; "
            "padding: 6px 14px; border-radius: 5px; border: none; } "
            f"QPushButton:hover {{ background-color: {C.INFO}; }} "
            f"QPushButton:disabled {{ background-color: {C.BORDER_STRONG}; color: {C.TEXT_MUTED}; }}"
        )
        self.pin_btn.clicked.connect(self._on_pin_selected)
        btn_row.addWidget(self.pin_btn)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)
        self._update_action_buttons()

    def _on_config_ui_changed(self, *_args):
        if self._loading_config:
            return
        self._config["coverage_filter"] = self.coverage_combo.currentData() or "UNTRACKED_AND_BODY"
        self._config["require_subject_tag"] = bool(self.require_tag_check.isChecked())
        self._config["subject_tag"] = self.tag_edit.text().strip()
        self._config["use_date_filter"] = bool(self.use_date_check.isChecked())
        self._config["start_date"] = self.start_date_edit.text().strip()
        self._config["end_date"] = self.end_date_edit.text().strip()
        self._config["days_before"] = int(self.days_before_spin.value())
        self._config["days_after"] = int(self.days_after_spin.value())
        save_meeting_discovery_config(self._meeting_dir, self._config)
        self._update_action_buttons()
        self._apply_filter()

    def _on_cancel_scan(self):
        if self._scan_worker and self._scan_worker.isRunning():
            self._scan_worker.request_cancel()

    def _stop_scan_worker(self):
        if getattr(self, "_scan_worker", None) and self._scan_worker.isRunning():
            self._scan_worker.request_cancel()
            self._scan_worker.wait(1500)

    def _save_geometry(self):
        try:
            if hasattr(self, "_settings"):
                self._settings.setValue("geometry", self.saveGeometry())
        except Exception:
            pass

    def accept(self):
        self._save_geometry()
        super().accept()

    def reject(self):
        self._save_geometry()
        self._stop_scan_worker()
        super().reject()

    def closeEvent(self, event):
        self._save_geometry()
        self._stop_scan_worker()
        super().closeEvent(event)

    def _load_threads(self):
        if self._scan_worker and self._scan_worker.isRunning():
            return
        self.reload_btn.setEnabled(False)
        self.reload_btn.setText("⏳ Scanning...")
        self.scan_bar.start_busy(f"Scanning {self.wg} Outlook folders & Sent Items...")

        worker = OutlookFolderScanWorker(self.wg, parent=self)
        worker.progress.connect(self.scan_bar.update_progress)
        worker.finished_scan.connect(self._on_scan_finished)
        self._scan_worker = worker
        worker.start()

    def _reclassify_threads(self):
        ignored_set = set(self._config.get("ignored_thread_keys") or [])
        self._threads = classify_discovered_threads(
            self._raw_threads,
            self.db,
            self._general_email_db,
            meeting_id=self.meeting_id,
            valid_tdocs=self._valid_tdocs,
            agenda_map=self._agenda_map,
            valid_wis=self._valid_wis,
            ignored_keys=ignored_set,
            tdoc_meta_map=self._tdoc_meta_map,
        )

    def _on_scan_finished(self, ok: bool, threads: list, err: str):
        self.reload_btn.setEnabled(True)
        self.reload_btn.setText("🔄 Rescan Folders")
        self._raw_threads = list(threads or [])
        self._reclassify_threads()
        self._apply_filter()
        if ok:
            self.scan_bar.finish_ok()
        else:
            self.scan_bar.show_error(err or "Folder scan did not finish cleanly.")

    @staticmethod
    def _parse_iso_date(val: str) -> Optional[datetime]:
        txt = str(val or "").strip()[:10]
        if not txt:
            return None
        try:
            return datetime.strptime(txt, "%Y-%m-%d")
        except Exception:
            return None

    def _matches_tag_and_date(self, t: dict) -> bool:
        if self.require_tag_check.isChecked():
            tag = self.tag_edit.text().strip().lower()
            if tag:
                subj_blob = f"{t.get('raw_subject', '')} {t.get('title', '')}".lower()
                if tag not in subj_blob:
                    return False

        if self.use_date_check.isChecked():
            last_dt = self._parse_iso_date(t.get("last_date", ""))
            start_dt = self._parse_iso_date(self.start_date_edit.text())
            end_dt = self._parse_iso_date(self.end_date_edit.text())
            if last_dt is not None:
                if start_dt is not None:
                    min_dt = start_dt - timedelta(days=int(self.days_before_spin.value()))
                    if last_dt < min_dt:
                        return False
                if end_dt is not None:
                    max_dt = end_dt + timedelta(days=int(self.days_after_spin.value()))
                    if last_dt > max_dt:
                        return False
        return True

    def _apply_filter(self):
        cov_mode = str(self.coverage_combo.currentData() or "UNTRACKED_AND_BODY").upper()
        q = self.search_edit.text().strip().lower()

        scoped: List[dict] = []
        counts = {"STRICT_UNTRACKED": 0, "BODY_TDOC": 0, "SUBJECT_TDOC": 0, "PINNED": 0, "IGNORED": 0}
        for t in self._threads:
            if not self._matches_tag_and_date(t):
                continue
            is_ignored = bool(t.get("is_ignored"))
            cov = str(t.get("coverage") or "STRICT_UNTRACKED")
            if is_ignored and cov != "PINNED":
                counts["IGNORED"] += 1
            elif cov in counts:
                counts[cov] += 1

            if cov_mode == "IGNORED":
                if not is_ignored:
                    continue
            else:
                if is_ignored and cov_mode != "ALL":
                    continue
                if cov_mode == "UNTRACKED_AND_BODY" and cov not in ("STRICT_UNTRACKED", "BODY_TDOC"):
                    continue
                if cov_mode == "STRICT_UNTRACKED" and cov != "STRICT_UNTRACKED":
                    continue
                if cov_mode == "ALL_UNPINNED" and cov == "PINNED":
                    continue

            if q:
                hay = " ".join([
                    str(t.get("title") or ""),
                    str(t.get("last_sender") or ""),
                    " ".join(t.get("suggested_tdocs") or []),
                    " ".join(t.get("suggested_ais") or []),
                    " ".join(t.get("suggested_wis") or []),
                ]).lower()
                if q not in hay:
                    continue
            scoped.append(t)

        self.table.setRowCount(len(scoped))
        for r, t in enumerate(scoped):
            cov = str(t.get("coverage") or "")
            cov_item = QTableWidgetItem(str(t.get("coverage_label") or cov))
            cov_item.setData(Qt.UserRole, t)
            if cov == "STRICT_UNTRACKED" and not t.get("is_ignored"):
                cov_item.setForeground(QColor(C.DANGER))
            elif cov == "BODY_TDOC" and not t.get("is_ignored"):
                cov_item.setForeground(QColor(C.WARNING))
            elif cov == "SUBJECT_TDOC":
                cov_item.setForeground(QColor(C.SUCCESS))
            elif cov == "PINNED":
                cov_item.setForeground(QColor(C.INFO))
            else:
                cov_item.setForeground(QColor(C.TEXT_MUTED))

            tot_cnt = int(t.get("count", 1) or 0)
            unread_cnt = int(t.get("unread_count", 0) or 0)
            msg_str = f"{tot_cnt} ({unread_cnt} new)" if unread_cnt > 0 else f"{tot_cnt}"
            msg_item = QTableWidgetItem(msg_str)
            msg_item.setTextAlignment(Qt.AlignCenter)
            if unread_cnt > 0:
                font = msg_item.font()
                font.setBold(True)
                msg_item.setFont(font)
                msg_item.setForeground(QColor(C.INFO))

            title_item = QTableWidgetItem(str(t.get("title") or ""))
            if t.get("raw_subject") and t.get("raw_subject") != t.get("title"):
                title_item.setToolTip(str(t.get("raw_subject")))

            sugg_parts = []
            if t.get("suggested_ais"):
                sugg_parts.append("AI: " + ", ".join(t["suggested_ais"][:3]))
            if t.get("suggested_wis"):
                sugg_parts.append("WI: " + ", ".join(t["suggested_wis"][:2]))
            if t.get("suggested_tdocs"):
                sugg_parts.append("TDoc: " + ", ".join(t["suggested_tdocs"][:3]))
            sugg_item = QTableWidgetItem(" · ".join(sugg_parts) if sugg_parts else "—")

            sender_item = QTableWidgetItem(str(t.get("last_sender") or ""))
            date_item = QTableWidgetItem(str(t.get("last_date") or ""))

            self.table.setItem(r, 0, cov_item)
            self.table.setItem(r, 1, msg_item)
            self.table.setItem(r, 2, title_item)
            self.table.setItem(r, 3, sugg_item)
            self.table.setItem(r, 4, sender_item)
            self.table.setItem(r, 5, date_item)

        self.summary_label.setText(
            f"Showing {len(scoped)} thread(s)  •  "
            f"🔴 {counts['STRICT_UNTRACKED']} no TDoc  |  "
            f"🟡 {counts['BODY_TDOC']} body-only  |  "
            f"🟢 {counts['SUBJECT_TDOC']} subject TDoc  |  "
            f"📌 {counts['PINNED']} pinned  |  "
            f"🚫 {counts['IGNORED']} ignored"
        )
        self._update_action_buttons()

    def _selected_threads(self) -> List[dict]:
        rows = sorted({idx.row() for idx in self.table.selectionModel().selectedRows()}) if self.table.selectionModel() else []
        out = []
        for r in rows:
            item = self.table.item(r, 0)
            data = item.data(Qt.UserRole) if item else None
            if isinstance(data, dict):
                out.append(data)
        return out

    def _update_action_buttons(self):
        cov_mode = str(self.coverage_combo.currentData() or "UNTRACKED_AND_BODY").upper()
        selected = self._selected_threads()
        if cov_mode == "IGNORED":
            self.ignore_btn.setText("↩️ Un-ignore Selected")
            self.ignore_btn.setToolTip("Restore selected thread(s) back into untracked discovery.")
        else:
            self.ignore_btn.setText("🚫 Ignore Selected")
            self.ignore_btn.setToolTip("Hide selected thread(s) from future untracked discovery for this meeting.")
        self.ignore_btn.setEnabled(bool(selected))
        self.pin_btn.setEnabled(bool(selected))

    def _on_toggle_ignore_selected(self):
        selected = self._selected_threads()
        if not selected:
            return
        cov_mode = str(self.coverage_combo.currentData() or "UNTRACKED_AND_BODY").upper()
        ignored = set(self._config.get("ignored_thread_keys") or [])
        for t in selected:
            nkey = str(t.get("norm_key") or "").strip().lower()
            if not nkey:
                continue
            if cov_mode == "IGNORED":
                ignored.discard(nkey)
            else:
                ignored.add(nkey)
        self._config["ignored_thread_keys"] = sorted(ignored)
        save_meeting_discovery_config(self._meeting_dir, self._config)
        self._reclassify_threads()
        self._apply_filter()

    def _on_pin_selected(self):
        selected = self._selected_threads()
        if not selected:
            QMessageBox.information(self, "No Thread Selected", "Please select at least one email thread to pin.")
            return
        ignored = set(self._config.get("ignored_thread_keys") or [])
        changed_ignored = False
        for t in selected:
            nkey = str(t.get("norm_key") or "").strip().lower()
            if nkey and nkey in ignored:
                ignored.discard(nkey)
                changed_ignored = True
            tid = pin_thread_by_subject_or_cid(
                self.db,
                title=t["title"],
                normalized_subject=t["normalized_subject"],
                conversation_id=t.get("conversation_id", ""),
                meeting_id=self.meeting_id,
                wg=self.wg,
                agenda_map=self._agenda_map,
                valid_wis=self._valid_wis,
                tdoc_meta_map=self._tdoc_meta_map,
            )
            self.pinned_thread_ids.append(tid)
        if changed_ignored:
            self._config["ignored_thread_keys"] = sorted(ignored)
            save_meeting_discovery_config(self._meeting_dir, self._config)
        self.accept()

