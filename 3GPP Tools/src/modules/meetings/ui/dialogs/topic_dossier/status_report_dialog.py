"""Modeless Corpus Status & Index Refresh report window."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Dict, List

from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QColor, QDesktopServices, QFont
from PyQt5.QtWidgets import (
    QDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from core.ui.font_scale import scale_px
from core.ui.theme import C, R
from core.utils.folder_opener import open_folder as _open_folder
from core.utils.paths import strip_extended_path, to_extended_path
from modules.meetings.ui.dialogs.topic_dossier.styles import (
    STYLE_PRIMARY_BTN,
    STYLE_SECONDARY_BTN,
)


def format_grouped_changed_meetings(changed_meetings: List[str]) -> str:
    """Groups `'WG/MeetingFolder'` keys by WG for compact display."""
    grouped: Dict[str, List[str]] = {}
    ungrouped: List[str] = []
    for raw in changed_meetings or []:
        entry = str(raw or "").strip()
        if not entry:
            continue
        if "/" in entry:
            wg, mtg = entry.split("/", 1)
            grouped.setdefault(wg.strip().upper(), []).append(mtg.strip())
        else:
            ungrouped.append(entry)

    parts = [f"<b>{wg}:</b> {', '.join(mtgs)}" for wg, mtgs in grouped.items()]
    if ungrouped:
        parts.append(", ".join(ungrouped))
    return " &nbsp;·&nbsp; ".join(parts)


def _pill_qss(bg: str, fg: str, border: str) -> str:
    return (
        f"background-color: {bg}; color: {fg}; border: 1px solid {border}; "
        f"border-radius: {R.SM}px; padding: 2px 8px; font-size: 11px; font-weight: bold;"
    )


class _NumericTableItem(QTableWidgetItem):
    """QTableWidgetItem sorting numerically by UserRole while displaying formatted text."""

    def __init__(self, text: str, num_val: int):
        super().__init__(text)
        self.setData(Qt.UserRole, int(num_val))
        self.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

    def __lt__(self, other: QTableWidgetItem) -> bool:
        v1 = self.data(Qt.UserRole)
        v2 = other.data(Qt.UserRole) if isinstance(other, QTableWidgetItem) else None
        if isinstance(v1, int) and isinstance(v2, int):
            return v1 < v2
        return super().__lt__(other)


class CorpusStatusReportDialog(QDialog):
    """Modeless themed report dialog summarizing Topic Dossier checkup results."""

    def __init__(self, parent=None):
        super().__init__(None)
        self.setWindowFlags(Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)
        self.setWindowModality(Qt.NonModal)
        self.setAttribute(Qt.WA_DeleteOnClose, False)
        self.setWindowTitle("Corpus Status & Index Refreshed")
        self.resize(scale_px(680), scale_px(470))
        self._target_dir: str = ""
        self._init_ui()

    def _init_ui(self) -> None:
        self.setStyleSheet(f"QDialog {{ background-color: {C.BG_CANVAS}; }}")
        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        # 1. Header Card
        hdr = QFrame()
        hdr.setObjectName("StatusReportHeader")
        hdr.setStyleSheet(
            f"QFrame#StatusReportHeader {{ background-color: {C.BG_CARD}; "
            f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: {R.LG}px; }}"
            f"QFrame#StatusReportHeader QLabel {{ background: transparent; border: none; }}"
        )
        hdr_layout = QVBoxLayout(hdr)
        hdr_layout.setContentsMargins(14, 12, 14, 12)
        hdr_layout.setSpacing(4)

        top_row = QHBoxLayout()
        top_row.setSpacing(8)
        self.lbl_topic_name = QLabel("Topic Corpus")
        self.lbl_topic_name.setStyleSheet(
            f"color: {C.TEXT_PRIMARY}; font-size: 15px; font-weight: bold;"
        )
        self.lbl_topic_name.setWordWrap(True)
        top_row.addWidget(self.lbl_topic_name, 1)

        self.lbl_elapsed_pill = QLabel("Refreshed")
        self.lbl_elapsed_pill.setObjectName("StatusElapsedPill")
        self.lbl_elapsed_pill.setStyleSheet(_pill_qss(C.INFO_SOFT, C.INFO_DARK, C.INFO_BORDER))
        top_row.addWidget(self.lbl_elapsed_pill, 0, Qt.AlignTop)
        hdr_layout.addLayout(top_row)

        self.lbl_target_dir = QLabel("")
        self.lbl_target_dir.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px;")
        self.lbl_target_dir.setWordWrap(True)
        hdr_layout.addWidget(self.lbl_target_dir)
        root.addWidget(hdr)

        # 2. KPI Strip (4 Metric Cards)
        kpi_grid = QGridLayout()
        kpi_grid.setSpacing(8)
        (
            self.kpi_meetings_val,
            self.kpi_meetings_sub,
            card_mtg,
        ) = self._make_kpi_card("Meetings", C.TEXT_PRIMARY)
        (
            self.kpi_decisions_val,
            self.kpi_decisions_sub,
            card_dec,
        ) = self._make_kpi_card("Decisions", C.INFO_DARK)
        (
            self.kpi_docs_val,
            self.kpi_docs_sub,
            card_doc,
        ) = self._make_kpi_card("Active Docs", C.TEXT_PRIMARY)
        (
            self.kpi_stage1_val,
            self.kpi_stage1_sub,
            card_s1,
        ) = self._make_kpi_card("Stage-1 Analyses", C.SUCCESS_DARK)
        kpi_grid.addWidget(card_mtg, 0, 0)
        kpi_grid.addWidget(card_dec, 0, 1)
        kpi_grid.addWidget(card_doc, 0, 2)
        kpi_grid.addWidget(card_s1, 0, 3)
        root.addLayout(kpi_grid)

        # 3. Stage-2 Synthesis Card
        s2_card = QFrame()
        s2_card.setObjectName("StatusStage2Card")
        s2_card.setStyleSheet(
            f"QFrame#StatusStage2Card {{ background-color: {C.BG_CARD}; "
            f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: {R.LG}px; }}"
            f"QFrame#StatusStage2Card QLabel {{ background: transparent; border: none; }}"
        )
        s2_layout = QVBoxLayout(s2_card)
        s2_layout.setContentsMargins(14, 12, 14, 12)
        s2_layout.setSpacing(6)

        s2_top = QHBoxLayout()
        s2_top.setSpacing(8)
        s2_title = QLabel("<b>Stage-2 Cross-WG Synthesis</b>")
        s2_title.setStyleSheet(f"color: {C.TEXT_PRIMARY}; font-size: 12px;")
        s2_top.addWidget(s2_title)

        self.lbl_s2_pill = QLabel("PENDING")
        self.lbl_s2_pill.setObjectName("StatusStage2Pill")
        s2_top.addWidget(self.lbl_s2_pill)

        self.lbl_s2_val_pill = QLabel("")
        self.lbl_s2_val_pill.setObjectName("StatusStage2ValPill")
        self.lbl_s2_val_pill.setVisible(False)
        s2_top.addWidget(self.lbl_s2_val_pill)
        s2_top.addStretch()
        s2_layout.addLayout(s2_top)

        self.lbl_s2_pkg = QLabel("")
        self.lbl_s2_pkg.setStyleSheet(f"color: {C.TEXT_BODY}; font-size: 11px;")
        self.lbl_s2_pkg.setWordWrap(True)
        s2_layout.addWidget(self.lbl_s2_pkg)

        self.lbl_s2_changed = QLabel("")
        self.lbl_s2_changed.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
        self.lbl_s2_changed.setWordWrap(True)
        self.lbl_s2_changed.setVisible(False)
        s2_layout.addWidget(self.lbl_s2_changed)
        root.addWidget(s2_card)

        # 4. Highlights Strip (Newly Promoted / Stubs / Stale / Pending)
        self.highlights_frame = QFrame()
        self.highlights_frame.setObjectName("StatusHighlightsFrame")
        self.highlights_frame.setStyleSheet(
            f"QFrame#StatusHighlightsFrame {{ background-color: {C.BG_SURFACE}; "
            f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: {R.MD}px; }}"
            f"QFrame#StatusHighlightsFrame QLabel {{ background: transparent; border: none; }}"
        )
        hl_layout = QVBoxLayout(self.highlights_frame)
        hl_layout.setContentsMargins(12, 8, 12, 8)
        self.lbl_highlights = QLabel("")
        self.lbl_highlights.setWordWrap(True)
        self.lbl_highlights.setStyleSheet(f"color: {C.TEXT_BODY}; font-size: 11px;")
        hl_layout.addWidget(self.lbl_highlights)
        self.highlights_frame.setVisible(False)
        root.addWidget(self.highlights_frame)

        # 5. Expandable Meeting Breakdown
        self.btn_toggle_breakdown = QPushButton("▸ Show Meeting Breakdown (0)")
        self.btn_toggle_breakdown.setCheckable(True)
        self.btn_toggle_breakdown.setChecked(False)
        self.btn_toggle_breakdown.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_toggle_breakdown.toggled.connect(self._on_toggle_breakdown)
        root.addWidget(self.btn_toggle_breakdown, 0, Qt.AlignLeft)

        self.table_breakdown = QTableWidget(0, 7)
        self.table_breakdown.setHorizontalHeaderLabels(
            ["WG", "Meeting", "Lifecycle", "Stage-1", "Decisions", "Docs", "Stubs"]
        )
        self.table_breakdown.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table_breakdown.setSelectionBehavior(QTableWidget.SelectRows)
        self.table_breakdown.setAlternatingRowColors(True)
        self.table_breakdown.setSortingEnabled(True)
        self.table_breakdown.verticalHeader().setVisible(False)
        self.table_breakdown.verticalHeader().setDefaultSectionSize(scale_px(26))
        self.table_breakdown.setStyleSheet(
            f"QTableWidget {{ background-color: {C.BG_CARD}; alternate-background-color: {C.BG_SURFACE}; "
            f"border: 1px solid {C.BORDER_STRONG}; border-radius: {R.MD}px; "
            f"gridline-color: {C.BORDER_ROW}; font-size: 11px; }}"
            f"QHeaderView::section {{ background-color: {C.BG_SURFACE}; color: {C.TEXT_SECONDARY}; "
            f"font-weight: bold; font-size: 11px; border: none; "
            f"border-bottom: 1px solid {C.BORDER_SUBTLE}; padding: 5px 6px; }}"
            f"QTableWidget::item {{ padding: 3px 6px; }}"
        )
        hdr_view = self.table_breakdown.horizontalHeader()
        hdr_view.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hdr_view.setSectionResizeMode(1, QHeaderView.Stretch)
        hdr_view.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        hdr_view.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        hdr_view.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        hdr_view.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        hdr_view.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        self.table_breakdown.setVisible(False)
        self.table_breakdown.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        root.addWidget(self.table_breakdown, 1)

        # 6. Footer Note & Actions
        self.lbl_footer_note = QLabel(
            "Refreshed: index.md · company_positions_e2e_prompt.md · per-meeting briefs · .export_manifest.json"
        )
        self.lbl_footer_note.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px;")
        self.lbl_footer_note.setWordWrap(True)
        root.addWidget(self.lbl_footer_note)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        self.btn_open_index = QPushButton("📄 Open index.md")
        self.btn_open_index.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_open_index.clicked.connect(self._on_open_index)
        btn_row.addWidget(self.btn_open_index)

        self.btn_open_folder = QPushButton("📂 Open Folder")
        self.btn_open_folder.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_open_folder.clicked.connect(self._on_open_folder)
        btn_row.addWidget(self.btn_open_folder)

        btn_row.addStretch()

        self.btn_close = QPushButton("Close")
        self.btn_close.setStyleSheet(STYLE_PRIMARY_BTN)
        self.btn_close.clicked.connect(self.close)
        btn_row.addWidget(self.btn_close)
        root.addLayout(btn_row)

    def _make_kpi_card(self, title: str, accent_color: str):
        card = QFrame()
        card.setObjectName("StatusKpiCard")
        card.setStyleSheet(
            f"QFrame#StatusKpiCard {{ background-color: {C.BG_CARD}; "
            f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: {R.MD}px; }}"
            f"QFrame#StatusKpiCard QLabel {{ background: transparent; border: none; }}"
        )
        lay = QVBoxLayout(card)
        lay.setContentsMargins(10, 8, 10, 8)
        lay.setSpacing(2)

        lbl_title = QLabel(title)
        lbl_title.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px; font-weight: bold;")
        lay.addWidget(lbl_title)

        lbl_val = QLabel("0")
        lbl_val.setStyleSheet(f"color: {accent_color}; font-size: 18px; font-weight: bold;")
        lay.addWidget(lbl_val)

        lbl_sub = QLabel("")
        lbl_sub.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
        lbl_sub.setWordWrap(True)
        lay.addWidget(lbl_sub)
        return lbl_val, lbl_sub, card

    def update_report(self, res: Dict[str, Any]) -> None:
        """Populates all KPI cards, Stage-2 synthesis details, highlights, and breakdown rows."""
        self._target_dir = strip_extended_path(str(res.get("target_dir") or ""))
        topic_name = str(res.get("topic_name") or "Topic Dossier")
        self.lbl_topic_name.setText(f"🌐 {topic_name}")
        self.lbl_target_dir.setText(self._target_dir)
        self.lbl_target_dir.setToolTip(self._target_dir)

        elapsed_ms = int(res.get("elapsed_ms", 0) or 0)
        if elapsed_ms >= 1000:
            self.lbl_elapsed_pill.setText(f"Refreshed in {elapsed_ms / 1000.0:.1f}s")
        elif elapsed_ms > 0:
            self.lbl_elapsed_pill.setText(f"Refreshed in {elapsed_ms} ms")
        else:
            self.lbl_elapsed_pill.setText("Refreshed")

        total_mtgs = int(res.get("total_meetings", 0) or 0)
        no_act = int(res.get("no_activity", 0) or 0)
        analyzed = int(res.get("analyzed", 0) or 0)
        stale = int(res.get("stale", 0) or 0)
        pending = int(res.get("pending", 0) or 0)
        total_decs = int(res.get("total_decisions", 0) or 0)
        total_docs = int(res.get("total_active_docs", 0) or 0)
        total_stubs = int(res.get("total_stubs", 0) or 0)
        promo = int(res.get("newly_promoted", 0) or 0)

        self.kpi_meetings_val.setText(f"{total_mtgs}")
        self.kpi_meetings_sub.setText(f"{no_act} no-activity" if no_act else "all active")

        self.kpi_decisions_val.setText(f"{total_decs:,}")
        self.kpi_decisions_sub.setText("unified decisions")

        self.kpi_docs_val.setText(f"{total_docs:,}")
        if total_stubs:
            self.kpi_docs_sub.setText(
                f"<span style='color: {C.WARNING_DARK}; font-weight: bold;'>{total_stubs} stub(s)</span>"
            )
        else:
            self.kpi_docs_sub.setText("0 stubs")

        self.kpi_stage1_val.setText(f"{analyzed}")
        s1_parts = [f"{analyzed} analyzed"]
        if stale:
            s1_parts.append(f"<span style='color: {C.WARNING_DARK};'>{stale} stale</span>")
        if pending:
            s1_parts.append(f"{pending} pending")
        self.kpi_stage1_sub.setText(" · ".join(s1_parts))

        # Stage-2 Card
        syn = str(res.get("synthesis_status") or "pending").lower()
        s2_state = res.get("synthesis_state") or {}
        latest_pkg = s2_state.get("latest_package") or {}
        all_pkgs = list(s2_state.get("packages") or [])
        completed_cnt = sum(1 for pk in all_pkgs if pk.get("status") == "completed")

        if syn == "analyzed":
            self.lbl_s2_pill.setText("✅ ANALYZED")
            self.lbl_s2_pill.setStyleSheet(_pill_qss(C.SUCCESS_SOFT_ALT, C.SUCCESS_DARK, C.SUCCESS_BORDER))
        elif syn == "stale":
            self.lbl_s2_pill.setText("🔄 STALE")
            self.lbl_s2_pill.setStyleSheet(_pill_qss(C.WARNING_SOFT_ALT, C.WARNING_DARK, C.WARNING_BORDER))
        elif syn == "in_progress":
            self.lbl_s2_pill.setText("🏗️ IN PROGRESS")
            self.lbl_s2_pill.setStyleSheet(_pill_qss(C.INFO_SOFT, C.INFO_DARK, C.INFO_BORDER))
        else:
            self.lbl_s2_pill.setText("⏳ PENDING")
            self.lbl_s2_pill.setStyleSheet(_pill_qss(C.BG_CANVAS, C.TEXT_MUTED, C.BORDER_STRONG))

        if latest_pkg:
            s2_folder = latest_pkg.get("folder_name") or "(root)"
            pkg_id = latest_pkg.get("package_id") or ""
            as_of = latest_pkg.get("as_of") or "N/A"
            rev = latest_pkg.get("revision") or ""
            runs_note = f" · {completed_cnt} runs total" if completed_cnt > 1 else ""
            self.lbl_s2_pkg.setText(
                f"<b>Package:</b> {s2_folder} &nbsp;|&nbsp; <b>ID:</b> {pkg_id} &nbsp;|&nbsp; "
                f"<b>Cut-off:</b> {as_of} {rev}{runs_note}"
            )
            val_passed = latest_pkg.get("validation_passed")
            claims = latest_pkg.get("claims")
            if val_passed is not None or claims is not None:
                claims_txt = f" · {claims} claims" if claims is not None else ""
                if val_passed is False:
                    self.lbl_s2_val_pill.setText(f"⚠️ ISSUES{claims_txt}")
                    self.lbl_s2_val_pill.setStyleSheet(_pill_qss(C.DANGER_SOFT, C.DANGER_DARK, C.DANGER_BORDER))
                else:
                    self.lbl_s2_val_pill.setText(f"✓ PASSED{claims_txt}")
                    self.lbl_s2_val_pill.setStyleSheet(_pill_qss(C.SUCCESS_SOFT, C.SUCCESS_DARK, C.SUCCESS_BORDER))
                self.lbl_s2_val_pill.setVisible(True)
            else:
                self.lbl_s2_val_pill.setVisible(False)
        else:
            self.lbl_s2_val_pill.setVisible(False)
            if syn == "in_progress":
                self.lbl_s2_pkg.setText(f"Active package in progress: <b>{s2_state.get('folder_name') or 'stage2'}</b>")
            else:
                self.lbl_s2_pkg.setText("No completed Stage-2 synthesis package found yet.")

        chg_list = list(s2_state.get("changed_meetings") or [])
        if chg_list:
            grouped_html = format_grouped_changed_meetings(chg_list)
            self.lbl_s2_changed.setText(
                f"<b>Updated since baseline ({len(chg_list)}):</b> {grouped_html}"
            )
            self.lbl_s2_changed.setVisible(True)
        else:
            self.lbl_s2_changed.setVisible(False)

        # Highlights
        records = list(res.get("meeting_records") or [])
        stub_mtgs = [r for r in records if int(r.get("stub_count", 0) or 0) > 0]
        stale_mtgs = [r.get("meeting_folder_id") for r in records if r.get("analysis_status") == "stale"]
        pending_mtgs = [r.get("meeting_folder_id") for r in records if r.get("analysis_status") == "pending"]
        promo_mtgs = list(res.get("newly_promoted_meetings") or [])

        hl_lines: List[str] = []
        if promo:
            promo_names = ", ".join(m.split("/")[-1] for m in promo_mtgs[:6])
            extra_p = f" (+{len(promo_mtgs) - 6} more)" if len(promo_mtgs) > 6 else ""
            hl_lines.append(
                f"<span style='color: {C.SUCCESS_DARK}; font-weight: bold;'>✅ Newly Promoted (+{promo}):</span> {promo_names}{extra_p}"
            )
        if stale_mtgs:
            hl_lines.append(
                f"<span style='color: {C.WARNING_DARK}; font-weight: bold;'>🔄 Stale Stage-1 ({len(stale_mtgs)}):</span> {', '.join(str(x) for x in stale_mtgs[:6])}"
            )
        if pending_mtgs:
            hl_lines.append(
                f"<span style='color: {C.TEXT_SECONDARY}; font-weight: bold;'>⏳ Pending Stage-1 ({len(pending_mtgs)}):</span> {', '.join(str(x) for x in pending_mtgs[:6])}"
            )
        if stub_mtgs:
            stub_desc = ", ".join(f"{r.get('meeting_folder_id')} ({r.get('stub_count')})" for r in stub_mtgs[:6])
            extra_s = f" (+{len(stub_mtgs) - 6} more)" if len(stub_mtgs) > 6 else ""
            hl_lines.append(
                f"<span style='color: {C.WARNING_DARK}; font-weight: bold;'>⚠️ Stubs ({total_stubs} across {len(stub_mtgs)} meeting(s)):</span> {stub_desc}{extra_s}"
            )

        if hl_lines:
            self.lbl_highlights.setText("<br/>".join(hl_lines))
            self.highlights_frame.setVisible(True)
        else:
            self.highlights_frame.setVisible(False)

        # Breakdown Table
        self._populate_breakdown_table(records)
        arrow = "▾ Hide" if self.btn_toggle_breakdown.isChecked() else "▸ Show"
        self.btn_toggle_breakdown.setText(f"{arrow} Meeting Breakdown ({len(records)})")

        has_dir = bool(self._target_dir and os.path.isdir(to_extended_path(self._target_dir)))
        self.btn_open_folder.setEnabled(has_dir)
        self.btn_open_index.setEnabled(
            has_dir and os.path.isfile(to_extended_path(Path(self._target_dir) / "index.md"))
        )

    def _populate_breakdown_table(self, records: List[Dict[str, Any]]) -> None:
        self.table_breakdown.setSortingEnabled(False)
        self.table_breakdown.setRowCount(len(records))
        for row_idx, rec in enumerate(records):
            wg = str(rec.get("wg") or "")
            fid = str(rec.get("meeting_folder_id") or "")
            lc = str(rec.get("lifecycle") or "").lower()
            st = str(rec.get("analysis_status") or "").lower()
            decs = int(rec.get("decision_count", 0) or 0)
            docs = int(rec.get("tdoc_count", 0) or 0)
            stubs = int(rec.get("stub_count", 0) or 0)
            is_promo = bool(rec.get("newly_promoted"))

            lc_label = "✅ Concluded" if lc == "concluded" else ("⚪ No Activity" if lc == "no_activity" else "🧭 Upcoming")
            if st == "analyzed":
                st_label = "✅ Analyzed*" if is_promo else "✅ Analyzed"
            elif st == "stale":
                st_label = "🔄 Stale"
            elif st == "no_activity":
                st_label = "⚪ Skip"
            else:
                st_label = "⏳ Pending"

            items = [
                QTableWidgetItem(wg),
                QTableWidgetItem(fid),
                QTableWidgetItem(lc_label),
                QTableWidgetItem(st_label),
                _NumericTableItem(str(decs), decs),
                _NumericTableItem(str(docs), docs),
                _NumericTableItem(str(stubs), stubs),
            ]
            if stubs > 0 or st == "stale":
                tint = QColor(C.WARNING_SOFT)
                for it in items:
                    it.setBackground(tint)
            elif is_promo:
                tint = QColor(C.SUCCESS_SOFT)
                for it in items:
                    it.setBackground(tint)
            if stubs > 0:
                f = QFont()
                f.setBold(True)
                items[6].setFont(f)
                items[6].setForeground(QColor(C.WARNING_DARK))

            for col_idx, it in enumerate(items):
                self.table_breakdown.setItem(row_idx, col_idx, it)
        self.table_breakdown.setSortingEnabled(True)

    def _on_toggle_breakdown(self, checked: bool) -> None:
        self.table_breakdown.setVisible(checked)
        count = self.table_breakdown.rowCount()
        self.btn_toggle_breakdown.setText(
            f"{'▾ Hide' if checked else '▸ Show'} Meeting Breakdown ({count})"
        )
        if checked:
            self.resize(max(self.width(), scale_px(720)), max(self.height(), scale_px(660)))
        else:
            self.resize(self.width(), scale_px(470))

    def _on_open_index(self) -> None:
        if not self._target_dir:
            return
        idx_path = Path(self._target_dir) / "index.md"
        if os.path.isfile(to_extended_path(idx_path)):
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(idx_path)))

    def _on_open_folder(self) -> None:
        if self._target_dir and os.path.isdir(to_extended_path(self._target_dir)):
            _open_folder(Path(self._target_dir))
