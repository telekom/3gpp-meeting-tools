"""Candidate ranking presentation, Stage 2 deep verification, and TDoc actions mixin."""

import html
from typing import Any, Dict, List

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QHBoxLayout, QLabel, QPushButton, QTableWidgetItem, QWidget

from core.ui.theme import C


class _AttributionResultsMixin:
    """Populates attribution candidates, candidate details, deep fingerprinting, and TDoc actions."""

    def _populate_attribution_view(self, candidates: List[Dict[str, Any]]):
        self._candidate_list = list(candidates)
        if not candidates:
            self.attr_hero_banner.setVisible(False)
            self.attr_table.setRowCount(0)
            self.attr_status_lbl.setText(f"No candidate CRs or p-CRs found for TS {self.spec_number} v{self.version_b}.")
            self.attr_details_browser.setHtml(
                f"<div style='color: {C.TEXT_MUTED}; padding: 12px; font-family: Segoe UI, sans-serif;'>"
                f"No change history records or candidate TDocs could be resolved for "
                f"<b>3GPP TS {html.escape(self.spec_number)} v{html.escape(self.version_b)}</b>.<br/><br/>"
                f"Tip: Ensure specification change history is indexed in <code>3gpp_spec_history.db</code>."
                f"</div>"
            )
            self.btn_attr_deep_verify.setEnabled(False)
            return

        self.btn_attr_deep_verify.setEnabled(True)
        self.attr_status_lbl.setText(
            f"<b>Found {len(candidates)} candidate contribution(s)</b> for Clause {self.clause_number} in v{self.version_b}:"
        )

        # 1. Hero banner for top candidate
        top = candidates[0]
        top_tdoc = top.get("tdoc", "")
        top_source = top.get("source", "")
        top_score = top.get("score", 0)
        top_conf = top.get("confidence", "High")
        top_verified = top.get("is_verified", False)

        if top_verified:
            self.attr_hero_banner.setStyleSheet(f"QFrame {{ background-color: {C.SUCCESS_SOFT}; border: 1px solid {C.SUCCESS_BORDER_MEDIUM}; border-radius: 6px; padding: 6px; }}")
            self.hero_match_title_lbl.setStyleSheet(f"font-size: 13px; color: {C.SUCCESS};")
            self.hero_match_title_lbl.setText(f"🎯 <b>Top Match (Verified):</b> {html.escape(top_tdoc)} {f'({html.escape(top_source)})' if top_source else ''} · {top_score}% Match")
        else:
            self.attr_hero_banner.setStyleSheet(f"QFrame {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; padding: 6px; }}")
            self.hero_match_title_lbl.setStyleSheet(f"font-size: 13px; color: {C.ACCENT_SKY_DARK};")
            self.hero_match_title_lbl.setText(f"🎯 <b>Top Candidate:</b> {html.escape(top_tdoc)} {f'({html.escape(top_source)})' if top_source else ''} · {top_conf} Confidence ({top_score}%)")

        self.hero_doc_title_lbl.setText(top.get("title", "") or "No title available")

        # Clear existing badges
        while self.hero_badges_layout.count():
            item = self.hero_badges_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        for b_text in top.get("evidence_badges", []):
            b_lbl = QLabel(b_text)
            b_lbl.setStyleSheet(f"background-color: {C.SKY_SOFT}; color: {C.ACCENT_SKY_DARK}; border: 1px solid {C.SKY_BORDER}; padding: 2px 8px; border-radius: 10px; font-size: 10px; font-weight: bold;")
            self.hero_badges_layout.addWidget(b_lbl)
        self.hero_badges_layout.addStretch()

        # Connect Hero Action Buttons
        try:
            self.btn_hero_open.clicked.disconnect()
        except Exception:
            pass
        self.btn_hero_open.clicked.connect(lambda: self._open_candidate_tdoc(top_tdoc, top.get("meeting", "")))

        try:
            self.btn_hero_jump.clicked.disconnect()
        except Exception:
            pass
        self.btn_hero_jump.clicked.connect(lambda: self._jump_candidate_to_meetings(top_tdoc))

        try:
            self.btn_hero_copy.clicked.disconnect()
        except Exception:
            pass
        self.btn_hero_copy.clicked.connect(lambda: self._copy_candidate_citation(top_tdoc, top.get("meeting", ""), top.get("title", "")))

        self.attr_hero_banner.setVisible(True)

        # 2. Populate Candidate Table
        self.attr_table.setRowCount(len(candidates))
        for row_idx, cand in enumerate(candidates):
            tdoc_tok = cand.get("tdoc", "")
            cr_num = cand.get("cr_number", "")
            tdoc_display = tdoc_tok if not cr_num else f"{tdoc_tok} (CR {cr_num})"
            src = cand.get("source", "")
            title = cand.get("title", "")
            score = cand.get("score", 0)
            conf = cand.get("confidence", "Low")
            verified = cand.get("is_verified", False)
            badges = cand.get("evidence_badges", [])

            # Col 0: Confidence Pill
            if verified:
                conf_text = f"✅ {score}% Verified"
            elif score >= 80:
                conf_text = f"🟢 {score}% {conf}"
            elif score >= 55:
                conf_text = f"🔵 {score}% {conf}"
            elif score >= 35:
                conf_text = f"🟠 {score}% {conf}"
            else:
                conf_text = f"⚪ {score}% {conf}"
            item_conf = QTableWidgetItem(conf_text)
            item_conf.setFlags(item_conf.flags() & ~Qt.ItemIsEditable)
            self.attr_table.setItem(row_idx, 0, item_conf)

            # Col 1: TDoc
            item_tdoc = QTableWidgetItem(tdoc_display)
            item_tdoc.setFlags(item_tdoc.flags() & ~Qt.ItemIsEditable)
            f_tdoc = QFont("Consolas")
            f_tdoc.setBold(True)
            item_tdoc.setFont(f_tdoc)
            self.attr_table.setItem(row_idx, 1, item_tdoc)

            # Col 2: Source
            item_src = QTableWidgetItem(src)
            item_src.setFlags(item_src.flags() & ~Qt.ItemIsEditable)
            self.attr_table.setItem(row_idx, 2, item_src)

            # Col 3: Title
            item_title = QTableWidgetItem(title)
            item_title.setFlags(item_title.flags() & ~Qt.ItemIsEditable)
            item_title.setToolTip(title)
            self.attr_table.setItem(row_idx, 3, item_title)

            # Col 4: Evidence Signals
            evidence_str = " • ".join(badges)
            item_ev = QTableWidgetItem(evidence_str)
            item_ev.setFlags(item_ev.flags() & ~Qt.ItemIsEditable)
            item_ev.setToolTip(evidence_str)
            self.attr_table.setItem(row_idx, 4, item_ev)

            # Col 5: Actions
            action_widget = QWidget()
            act_layout = QHBoxLayout(action_widget)
            act_layout.setContentsMargins(2, 2, 2, 2)
            act_layout.setSpacing(4)

            btn_open = QPushButton("📄")
            btn_open.setToolTip(f"Open {tdoc_tok} in Microsoft Word")
            btn_open.setStyleSheet(f"padding: 2px 6px; font-size: 10px; font-weight: bold; background-color: {C.BG_CANVAS}; border: 1px solid {C.BORDER_STRONG}; border-radius: 3px;")
            btn_open.clicked.connect(lambda _, t=tdoc_tok, m=cand.get("meeting", ""): self._open_candidate_tdoc(t, m))
            act_layout.addWidget(btn_open)

            btn_jump = QPushButton("🚀")
            btn_jump.setToolTip(f"Jump to Meetings tab for {tdoc_tok}")
            btn_jump.setStyleSheet(f"padding: 2px 6px; font-size: 10px; font-weight: bold; background-color: {C.BG_CANVAS}; border: 1px solid {C.BORDER_STRONG}; border-radius: 3px;")
            btn_jump.clicked.connect(lambda _, t=tdoc_tok: self._jump_candidate_to_meetings(t))
            act_layout.addWidget(btn_jump)

            btn_cp = QPushButton("📋")
            btn_cp.setToolTip(f"Copy citation for {tdoc_tok}")
            btn_cp.setStyleSheet(f"padding: 2px 6px; font-size: 10px; font-weight: bold; background-color: {C.BG_CANVAS}; border: 1px solid {C.BORDER_STRONG}; border-radius: 3px;")
            btn_cp.clicked.connect(lambda _, t=tdoc_tok, m=cand.get("meeting", ""), s=title: self._copy_candidate_citation(t, m, s))
            act_layout.addWidget(btn_cp)

            self.attr_table.setCellWidget(row_idx, 5, action_widget)

        self.attr_table.selectRow(0)
        self._on_candidate_table_selection_changed()

    def _on_candidate_table_selection_changed(self):
        """Displays rich metadata and snippet for currently selected candidate."""
        selected_rows = self.attr_table.selectedIndexes()
        if not selected_rows:
            return
        row = selected_rows[0].row()
        if row < 0 or row >= len(self._candidate_list):
            return
        cand = self._candidate_list[row]

        tdoc = cand.get("tdoc", "")
        cr_num = cand.get("cr_number", "")
        source = cand.get("source", "Unknown")
        meeting = cand.get("meeting", "")
        agenda_item = cand.get("agenda_item", "")
        title = cand.get("title", "")
        abstract = cand.get("abstract", "")
        remarks = cand.get("remarks", "")
        score = cand.get("score", 0)
        confidence = cand.get("confidence", "Low")
        badges = cand.get("evidence_badges", [])
        is_verified = cand.get("is_verified", False)
        match_pct = cand.get("verified_match_pct", 0)
        snippet = cand.get("verified_snippet", "")

        badges_html = " ".join([
            f"<span style='background-color: {C.SKY_SOFT}; color: {C.ACCENT_SKY_DARK}; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: bold;'>{html.escape(b)}</span>"
            for b in badges
        ])

        verified_html = ""
        if is_verified:
            verified_html = f"""
            <div style='background-color: {C.SUCCESS_SOFT_ALT}; border: 1px solid {C.SUCCESS_BORDER_MEDIUM}; border-radius: 4px; padding: 6px 10px; margin: 8px 0;'>
                <span style='color: {C.SUCCESS}; font-weight: bold;'>✅ Deterministic OpenXML Text Match ({match_pct}%):</span><br/>
                <code style='font-family: Consolas, monospace; font-size: 11px; color: {C.SUCCESS_DARK};'>"{html.escape(snippet)}"</code>
            </div>
            """

        details_html = f"""
        <div style='font-family: Segoe UI, sans-serif;'>
            <div style='margin-bottom: 6px;'>
                <b style='font-size: 13px; color: {C.ACCENT_SKY};'>{html.escape(tdoc)}</b>
                {f" <span style='color: {C.TEXT_MUTED};'>(CR #{html.escape(cr_num)})</span>" if cr_num else ""}
                &nbsp;—&nbsp; <span style='font-weight: 600; color: {C.TEXT_BODY};'>{html.escape(source)}</span>
                &nbsp;•&nbsp; <span style='color: {C.TEXT_MUTED};'>Meeting: <b>{html.escape(meeting or 'N/A')}</b></span>
                {f" &nbsp;•&nbsp; <span style='color: {C.TEXT_MUTED};'>AI: <b>{html.escape(agenda_item)}</b></span>" if agenda_item else ""}
                &nbsp;|&nbsp; <span style='color: {C.ACCENT_SKY_DARK}; font-weight: bold;'>Score: {score}% ({confidence})</span>
            </div>
            <div style='margin-bottom: 6px; font-size: 12px; font-weight: 600; color: {C.TEXT_STRONG};'>
                {html.escape(title or 'No title available')}
            </div>
            <div style='margin-bottom: 8px;'>{badges_html}</div>
            {verified_html}
            {f"<div style='margin-top: 6px; color: {C.TEXT_SECONDARY};'><b>Abstract / Summary:</b> {html.escape(abstract)}</div>" if abstract else ""}
            {f"<div style='margin-top: 4px; color: {C.TEXT_MUTED}; font-size: 10px;'><b>Secretary Remarks:</b> {html.escape(remarks)}</div>" if remarks else ""}
        </div>
        """
        self.attr_details_browser.setHtml(details_html)

    def _run_deep_verify(self):
        """Runs targeted Stage 2 OpenXML fingerprinting on top 3 candidates via background thread."""
        if not self._candidate_list or not self._raw_diff_text:
            return

        self.btn_attr_deep_verify.setEnabled(False)
        self.attr_progress_bar.setVisible(True)
        self.attr_progress_bar.setValue(10)
        self.attr_status_lbl.setText("Running targeted Stage 2 OpenXML text fingerprinting...")

        if not hasattr(self, "attr_engine"):
            from modules.specifications.core.spec_clause_attribution import SpecClauseAttributionEngine
            self.attr_engine = SpecClauseAttributionEngine(search_db=self.db)

        from modules.specifications.core.spec_clause_attribution import SpecClauseDeepVerifyThread
        self._deep_verify_worker = SpecClauseDeepVerifyThread(
            engine=self.attr_engine,
            candidates=self._candidate_list,
            diff_text=self._raw_diff_text,
            parent=self,
        )

        def _on_dv_progress(val: int, msg: str):
            self.attr_progress_bar.setValue(val)
            self.attr_status_lbl.setText(f"<b>Deep Verify:</b> {html.escape(msg)}")

        def _on_dv_finished(updated_cands: List[Dict[str, Any]]):
            self.attr_progress_bar.setVisible(False)
            self.btn_attr_deep_verify.setEnabled(True)
            self._cached_attribution_results = updated_cands
            self._populate_attribution_view(updated_cands)

        def _on_dv_error(err: str):
            self.attr_progress_bar.setVisible(False)
            self.btn_attr_deep_verify.setEnabled(True)
            self.attr_status_lbl.setText(f"<span style='color: {C.DANGER};'><b>Deep Verify Error:</b> {html.escape(err)}</span>")

        self._deep_verify_worker.progress.connect(_on_dv_progress)
        self._deep_verify_worker.finished.connect(_on_dv_finished)
        self._deep_verify_worker.error.connect(_on_dv_error)
        self._deep_verify_worker.start()

    def _open_candidate_tdoc(self, tdoc_id: str, meeting: str = ""):
        from modules.specifications.core.spec_tdoc_opener import SpecTDocOpener
        SpecTDocOpener.open_tdoc(tdoc_id, meeting_hint=meeting, parent=self)

    def _jump_candidate_to_meetings(self, tdoc_id: str):
        from modules.specifications.core.spec_tdoc_opener import SpecTDocOpener
        SpecTDocOpener.jump_to_meetings_tab(tdoc_id)

    def _copy_candidate_citation(self, tdoc_id: str, meeting: str = "", title: str = ""):
        from modules.specifications.core.spec_tdoc_opener import SpecTDocOpener
        SpecTDocOpener.copy_tdoc_citation(
            tdoc_id=tdoc_id,
            spec_number=self.spec_number,
            version=self.version_b,
            meeting=meeting,
            subject=title,
            parent=self,
        )
