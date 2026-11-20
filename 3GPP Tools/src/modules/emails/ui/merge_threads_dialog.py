# --- File: src/modules/emails/ui/merge_threads_dialog.py ---
import re
from typing import Any, Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QFont
from PyQt5.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QVBoxLayout,
)

from core.ui.font_scale import scale_px
from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY

_SUBJ_PREFIX_RE = re.compile(
    r"^(?:\s*(?:re|fwd?|aw|wg|sv|antw|回复|答复|转发)\s*[:：]\s*)+",
    re.IGNORECASE,
)
_SUBJ_TAG_RE = re.compile(r"^\s*\[[^\]]{2,80}\]\s*")
_STOP_WORDS = frozenset({
    "the", "and", "for", "with", "from", "into", "about", "regarding",
    "draft", "update", "updated", "revision", "rev", "interim", "conclusion",
    "conclusions", "agreement", "agreements", "discussion", "email", "thread",
    "baseline", "part", "new", "version", "sa2", "sa3", "ct1", "ran", "3gpp",
})


def _clean_subject_key(text: str) -> str:
    s = str(text or "").strip()
    changed = True
    while changed:
        changed = False
        ns = _SUBJ_PREFIX_RE.sub("", s).strip()
        if ns != s:
            s = ns
            changed = True
        ns = _SUBJ_TAG_RE.sub("", s).strip()
        if ns != s:
            s = ns
            changed = True
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def _subject_tokens(text: str) -> Set[str]:
    clean = _clean_subject_key(text)
    tokens = set(re.findall(r"[a-z0-9]{3,}(?:[._#-][a-z0-9]+)*", clean))
    return {t for t in tokens if t not in _STOP_WORDS}


def _thread_item_key(item: Dict[str, Any]) -> str:
    tid = str(item.get("thread_id") or "").strip()
    if tid:
        return f"THR::{tid}"
    eid = str(item.get("id") or "").strip()
    return f"ITEM::{eid}" if eid else ""


def _thread_msg_count(item: Dict[str, Any]) -> int:
    for k in ("message_count", "_total_count"):
        val = item.get(k)
        if val is not None:
            try:
                return int(val)
            except (TypeError, ValueError):
                pass
    return 1


def _thread_last_date(item: Dict[str, Any]) -> str:
    return str(item.get("last_message_date") or item.get("date_str") or item.get("date_sent") or "").strip()


def score_thread_similarity(
    target: Dict[str, Any],
    candidate: Dict[str, Any],
) -> Tuple[int, List[str]]:
    """
    Computes a deterministic similarity score + human-readable reasons between
    `target` and `candidate` to surface likely split threads at the top.
    """
    score = 0
    reasons: List[str] = []

    t_subj = _clean_subject_key(target.get("normalized_subject") or target.get("title") or target.get("subject") or "")
    c_subj = _clean_subject_key(candidate.get("normalized_subject") or candidate.get("title") or candidate.get("subject") or "")
    if t_subj and c_subj:
        if t_subj == c_subj:
            score += 100
            reasons.append("Exact subject match")
        elif t_subj in c_subj or c_subj in t_subj:
            score += 75
            reasons.append("Matching subject phrase")
        else:
            t_toks = _subject_tokens(t_subj)
            c_toks = _subject_tokens(c_subj)
            if t_toks and c_toks:
                shared = t_toks & c_toks
                union = t_toks | c_toks
                jaccard = len(shared) / max(1, len(union))
                if jaccard >= 0.35 or len(shared) >= 2:
                    score += int(round(60 * jaccard)) + min(20, 6 * len(shared))
                    reasons.append(f"Similar topic ({', '.join(sorted(shared)[:3])})")

    t_tdocs = {str(x).strip().upper() for x in (target.get("linked_tdocs") or []) if str(x).strip()}
    c_tdocs = {str(x).strip().upper() for x in (candidate.get("linked_tdocs") or []) if str(x).strip()}
    shared_tdocs = sorted(t_tdocs & c_tdocs)
    if shared_tdocs:
        score += 35 * min(3, len(shared_tdocs))
        reasons.append(f"Same TDoc: {', '.join(shared_tdocs[:2])}")

    t_subs = {str(x).strip() for x in (target.get("linked_sub_items") or []) if str(x).strip()}
    c_subs = {str(x).strip() for x in (candidate.get("linked_sub_items") or []) if str(x).strip()}
    shared_subs = sorted(t_subs & c_subs)
    if shared_subs:
        score += 25 * min(2, len(shared_subs))
        reasons.append(f"Same Sub-Topic: {', '.join(shared_subs[:2])}")

    t_ais = {str(x).strip() for x in (target.get("linked_ais") or []) if str(x).strip()}
    c_ais = {str(x).strip() for x in (candidate.get("linked_ais") or []) if str(x).strip()}
    shared_ais = sorted(t_ais & c_ais)
    if shared_ais:
        score += 15 * min(2, len(shared_ais))
        reasons.append(f"Same AI: {', '.join(shared_ais[:2])}")

    t_wis = {str(x).strip() for x in (target.get("linked_wis") or []) if str(x).strip()}
    c_wis = {str(x).strip() for x in (candidate.get("linked_wis") or []) if str(x).strip()}
    shared_wis = sorted(t_wis & c_wis)
    if shared_wis:
        score += 10 * min(2, len(shared_wis))
        reasons.append(f"Same WI: {', '.join(shared_wis[:2])}")

    return score, reasons


def choose_default_primary_thread(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Selects the best default primary target thread from a list of selected threads/items."""
    if not items:
        return {}

    def _rank(it: Dict[str, Any]) -> Tuple[int, int, int, int, str]:
        is_tracked = 1 if (it.get("in_tracked_db") or str(it.get("thread_id") or "").startswith("THR-")) else 0
        msgs = _thread_msg_count(it)
        status = str(it.get("user_status") or "").strip().lower()
        has_ann = int(bool(
            str(it.get("user_notes") or "").strip()
            or (status and status not in ("neutral", "⚪ neutral"))
            or (it.get("user_tags") or [])
        ))
        links = len(it.get("linked_tdocs") or []) + len(it.get("linked_ais") or []) + len(it.get("linked_wis") or [])
        return (is_tracked, msgs, has_ann, links, _thread_last_date(it))

    return max(items, key=_rank)


class MergeThreadsDialog(QDialog):
    """
    Searchable, multi-select dialog for merging Outlook-split email threads.

    Supports two workflows:
    1. Search-and-merge into an already-open target thread (`target_thread` + `candidates`).
    2. Direct select-and-merge from multi-selection (`selected_items`, with optional `candidates`).
    """

    def __init__(
        self,
        *,
        target_thread: Optional[Dict[str, Any]] = None,
        candidates: Optional[List[Dict[str, Any]]] = None,
        selected_items: Optional[List[Dict[str, Any]]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self._selected_items: List[Dict[str, Any]] = [dict(x) for x in (selected_items or [])]
        self._all_candidates: List[Dict[str, Any]] = [dict(x) for x in (candidates or [])]
        self._is_multi_select_mode = len(self._selected_items) >= 2

        if self._is_multi_select_mode:
            self._target_item = dict(target_thread) if target_thread else dict(choose_default_primary_thread(self._selected_items))
        else:
            self._target_item = dict(target_thread or (self._selected_items[0] if self._selected_items else {}))

        self._checked_keys: Set[str] = set()
        if self._is_multi_select_mode:
            t_key = _thread_item_key(self._target_item)
            for it in self._selected_items:
                k = _thread_item_key(it)
                if k and k != t_key:
                    self._checked_keys.add(k)

        self._items_by_key: Dict[str, Dict[str, Any]] = {}
        self._selected_keys: Set[str] = {_thread_item_key(it) for it in self._selected_items if _thread_item_key(it)}

        self.setWindowTitle(
            f"🔗 Merge Selected Threads ({len(self._selected_items)})"
            if self._is_multi_select_mode
            else "🔗 Merge Split Threads"
        )
        self.resize(scale_px(680), scale_px(560))
        self.setMinimumSize(scale_px(500), scale_px(400))

        self._build_ui()
        self._refresh_target_summary()
        self._populate_list()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        # Top Target Card
        target_card = QFrame(self)
        target_card.setObjectName("MergeTargetCard")
        target_card.setStyleSheet(
            f"#MergeTargetCard {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_STRONG}; "
            f"border-radius: 6px; padding: 6px; }}"
        )
        t_lay = QVBoxLayout(target_card)
        t_lay.setContentsMargins(10, 8, 10, 8)
        t_lay.setSpacing(4)

        header_lbl = QLabel(
            "Primary Target Thread (kept — all checked threads below merge into it):"
            if self._is_multi_select_mode
            else "Target Thread (kept — select one or more split threads to merge into it):"
        )
        header_lbl.setStyleSheet(f"font-weight: 600; color: {C.TEXT_SECONDARY}; font-size: {scale_px(12)}px;")
        t_lay.addWidget(header_lbl)

        if self._is_multi_select_mode:
            self.target_combo = QComboBox(target_card)
            self.target_combo.setMinimumHeight(scale_px(28))
            t_key = _thread_item_key(self._target_item)
            default_idx = 0
            for idx, it in enumerate(self._selected_items):
                title = str(it.get("title") or it.get("subject") or it.get("thread_id") or "(No Subject)")
                tid = str(it.get("thread_id") or "").strip()
                prefix = f"[{tid}] " if tid else "[Unpinned] "
                msgs = _thread_msg_count(it)
                self.target_combo.addItem(f"{prefix}{title} ({msgs} msgs)", _thread_item_key(it))
                if _thread_item_key(it) == t_key:
                    default_idx = idx
            self.target_combo.setCurrentIndex(default_idx)
            self.target_combo.currentIndexChanged.connect(self._on_target_combo_changed)
            t_lay.addWidget(self.target_combo)
        else:
            self.target_combo = None

        self.target_info_lbl = QLabel("")
        self.target_info_lbl.setWordWrap(True)
        self.target_info_lbl.setStyleSheet(f"color: {C.TEXT_PRIMARY}; font-size: {scale_px(12)}px;")
        t_lay.addWidget(self.target_info_lbl)
        root.addWidget(target_card)

        # Search + filter bar
        filter_row = QHBoxLayout()
        filter_row.setContentsMargins(0, 0, 0, 0)
        filter_row.setSpacing(8)

        self.search_edit = QLineEdit(self)
        self.search_edit.setPlaceholderText("🔍 Search title, thread ID, TDoc, AI, Sub-Topic, WI, or note...")
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.setMinimumHeight(scale_px(28))
        self.search_edit.textChanged.connect(self._populate_list)
        filter_row.addWidget(self.search_edit, 1)

        if self._is_multi_select_mode and self._all_candidates:
            self.show_other_chk = QCheckBox("Include other meeting threads", self)
            self.show_other_chk.setToolTip("Show additional tracked threads in this meeting so you can merge more than just the current selection.")
            self.show_other_chk.toggled.connect(self._populate_list)
            filter_row.addWidget(self.show_other_chk)
            self.likely_only_chk = None
        else:
            self.show_other_chk = None
            self.likely_only_chk = QCheckBox("✨ Likely matches only", self)
            self.likely_only_chk.setToolTip("Show only threads with matching title, subject words, TDocs, AIs, Sub-Topics, or WIs.")
            self.likely_only_chk.toggled.connect(self._populate_list)
            filter_row.addWidget(self.likely_only_chk)

        root.addLayout(filter_row)

        # Candidate list
        self.candidates_list = QListWidget(self)
        self.candidates_list.setAlternatingRowColors(True)
        self.candidates_list.setWordWrap(True)
        self.candidates_list.setStyleSheet(
            f"QListWidget {{ background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; }} "
            f"QListWidget::item {{ padding: 6px 8px; border-bottom: 1px solid {C.BORDER_ROW}; }} "
            f"QListWidget::item:selected {{ background-color: {C.ACCENT_SOFT}; color: {C.TEXT_PRIMARY}; }}"
        )
        self.candidates_list.itemChanged.connect(self._on_list_item_changed)
        self.candidates_list.itemDoubleClicked.connect(self._on_list_item_double_clicked)
        root.addWidget(self.candidates_list, 1)

        # Footer action bar
        footer = QHBoxLayout()
        footer.setContentsMargins(0, 0, 0, 0)
        footer.setSpacing(8)

        self.summary_lbl = QLabel("")
        self.summary_lbl.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: {scale_px(12)}px;")
        footer.addWidget(self.summary_lbl, 1)

        self.btn_cancel = QPushButton("Cancel", self)
        self.btn_cancel.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_cancel.setMinimumHeight(scale_px(28))
        self.btn_cancel.clicked.connect(self.reject)
        footer.addWidget(self.btn_cancel)

        self.btn_merge = QPushButton("🔗 Merge Threads", self)
        self.btn_merge.setStyleSheet(
            f"QPushButton {{ background-color: {C.ACCENT_PRIMARY}; color: {C.TEXT_ON_ACCENT}; font-weight: bold; "
            f"padding: 4px 12px; border-radius: 4px; border: none; }} "
            f"QPushButton:hover {{ background-color: {C.ACCENT_HOVER}; }} "
            f"QPushButton:disabled {{ background-color: {C.BORDER_STRONG}; color: {C.TEXT_MUTED}; }}"
        )
        self.btn_merge.setMinimumHeight(scale_px(28))
        self.btn_merge.setDefault(True)
        self.btn_merge.clicked.connect(self.accept)
        footer.addWidget(self.btn_merge)
        root.addLayout(footer)

    def _on_target_combo_changed(self):
        if self.target_combo is None:
            return
        new_key = str(self.target_combo.currentData() or "")
        old_key = _thread_item_key(self._target_item)
        for it in self._selected_items:
            if _thread_item_key(it) == new_key:
                self._target_item = dict(it)
                break
        if old_key and old_key != new_key:
            self._checked_keys.add(old_key)
        self._checked_keys.discard(new_key)
        self._refresh_target_summary()
        self._populate_list()

    def _refresh_target_summary(self):
        it = self._target_item
        title = str(it.get("title") or it.get("subject") or "(No Subject)")
        tid = str(it.get("thread_id") or "").strip()
        badge = tid if tid else "Unpinned (will pin automatically)"
        msgs = _thread_msg_count(it)
        ais = ", ".join(it.get("linked_ais") or []) or "—"
        subs = ", ".join(it.get("linked_sub_items") or []) or "—"
        tdocs = ", ".join((it.get("linked_tdocs") or [])[:4]) or "—"
        if self._is_multi_select_mode:
            self.target_info_lbl.setText(
                f"<b>{badge}</b> · {msgs} msgs · AI: {ais} · Sub-Topic: {subs} · TDocs: {tdocs}"
            )
        else:
            self.target_info_lbl.setText(
                f"<b>{title}</b><br/>"
                f"<span style='color:{C.TEXT_MUTED};'>{badge} · {msgs} msgs · AI: {ais} · Sub-Topic: {subs} · TDocs: {tdocs}</span>"
            )

    def _gather_pool(self) -> List[Dict[str, Any]]:
        t_key = _thread_item_key(self._target_item)
        pool_map: Dict[str, Dict[str, Any]] = {}
        for it in self._selected_items:
            k = _thread_item_key(it)
            if k and k != t_key:
                pool_map[k] = dict(it)
        include_extra = (
            not self._is_multi_select_mode
            or bool(self.show_other_chk and self.show_other_chk.isChecked())
            or bool(self.search_edit.text().strip())
        )
        if include_extra:
            for it in self._all_candidates:
                k = _thread_item_key(it)
                if k and k != t_key and k not in pool_map:
                    pool_map[k] = dict(it)
        self._items_by_key.update(pool_map)
        return list(pool_map.values())

    @staticmethod
    def _matches_search(it: Dict[str, Any], query_terms: List[str]) -> bool:
        if not query_terms:
            return True
        haystack = " ".join([
            str(it.get("thread_id") or ""),
            str(it.get("title") or ""),
            str(it.get("subject") or ""),
            str(it.get("normalized_subject") or ""),
            " ".join(it.get("linked_ais") or []),
            " ".join(it.get("linked_sub_items") or []),
            " ".join(it.get("linked_wis") or []),
            " ".join(it.get("linked_tdocs") or []),
            " ".join(it.get("user_tags") or []),
            " ".join(it.get("_companies") or []),
            str(it.get("company") or ""),
            str(it.get("user_notes") or ""),
        ]).lower()
        return all(term in haystack for term in query_terms)

    def _populate_list(self):
        query_terms = [t for t in self.search_edit.text().strip().lower().split() if t]
        likely_only = bool(self.likely_only_chk and self.likely_only_chk.isChecked())

        pool = self._gather_pool()
        scored_rows: List[Tuple[int, int, int, str, Dict[str, Any], List[str]]] = []
        for it in pool:
            if not self._matches_search(it, query_terms):
                continue
            k = _thread_item_key(it)
            score, reasons = score_thread_similarity(self._target_item, it)
            if likely_only and score <= 0 and k not in self._checked_keys:
                continue
            in_selection = 1 if k in self._selected_keys else 0
            scored_rows.append((
                in_selection,
                score,
                _thread_msg_count(it),
                _thread_last_date(it),
                it,
                reasons,
            ))

        scored_rows.sort(key=lambda r: (r[0], r[1], r[2], r[3]), reverse=True)

        self.candidates_list.blockSignals(True)
        self.candidates_list.clear()
        for in_sel, score, msgs, last_dt, it, reasons in scored_rows:
            k = _thread_item_key(it)
            tid = str(it.get("thread_id") or "").strip()
            title = str(it.get("title") or it.get("subject") or tid or "(No Subject)")
            ais = ", ".join(it.get("linked_ais") or [])
            subs = ", ".join(it.get("linked_sub_items") or [])
            tdocs = ", ".join((it.get("linked_tdocs") or [])[:4])
            meta_bits = [tid if tid else "Unpinned (will auto-pin)", f"{msgs} msgs"]
            if last_dt:
                meta_bits.append(last_dt[:16])
            if ais:
                meta_bits.append(f"AI: {ais}")
            if subs:
                meta_bits.append(f"Sub: {subs}")
            if tdocs:
                meta_bits.append(f"TDocs: {tdocs}")

            line1 = f"{'📌 ' if in_sel else ''}{'✨ ' if score > 0 else ''}{title}"
            line2 = " · ".join(meta_bits)
            if reasons:
                line2 += f"\n   Match: {' · '.join(reasons)}"

            item = QListWidgetItem(f"{line1}\n   {line2}")
            item.setData(Qt.UserRole, k)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            item.setCheckState(Qt.Checked if k in self._checked_keys else Qt.Unchecked)
            if score >= 75:
                item.setBackground(QColor(C.ACCENT_SOFT))
                item_font = item.font()
                item_font.setWeight(QFont.DemiBold)
                item.setFont(item_font)
            elif score > 0:
                item.setBackground(QColor(C.INFO_SOFT))
            self.candidates_list.addItem(item)

        if not scored_rows:
            empty = QListWidgetItem("No matching threads found for the current filter.")
            empty.setFlags(Qt.NoItemFlags)
            empty.setForeground(QColor(C.TEXT_MUTED))
            self.candidates_list.addItem(empty)

        self.candidates_list.blockSignals(False)
        self._update_summary()

    def _on_list_item_changed(self, item: QListWidgetItem):
        key = str(item.data(Qt.UserRole) or "")
        if not key:
            return
        if item.checkState() == Qt.Checked:
            self._checked_keys.add(key)
        else:
            self._checked_keys.discard(key)
        self._update_summary()

    def _on_list_item_double_clicked(self, item: QListWidgetItem):
        if not (item.flags() & Qt.ItemIsUserCheckable):
            return
        item.setCheckState(Qt.Unchecked if item.checkState() == Qt.Checked else Qt.Checked)

    def _update_summary(self):
        selected = self.selected_source_items()
        cnt = len(selected)
        extra_msgs = sum(_thread_msg_count(it) for it in selected)
        t_label = str(self._target_item.get("thread_id") or self._target_item.get("title") or "target")
        if cnt:
            self.summary_lbl.setText(f"Will merge {cnt} thread(s) (+{extra_msgs} msgs) into {t_label}")
            self.btn_merge.setText(f"🔗 Merge {cnt + 1} Threads" if self._is_multi_select_mode else f"🔗 Merge ({cnt})")
            self.btn_merge.setEnabled(True)
        else:
            self.summary_lbl.setText("Check at least one source thread to merge into the target.")
            self.btn_merge.setText("🔗 Merge Threads")
            self.btn_merge.setEnabled(False)

    def target_item(self) -> Dict[str, Any]:
        return dict(self._target_item)

    def target_thread_id(self) -> str:
        return str(self._target_item.get("thread_id") or "").strip()

    def selected_source_items(self) -> List[Dict[str, Any]]:
        t_key = _thread_item_key(self._target_item)
        out: List[Dict[str, Any]] = []
        for k in self._checked_keys:
            if k != t_key and k in self._items_by_key:
                out.append(dict(self._items_by_key[k]))
        return out

    def selected_source_ids(self) -> List[str]:
        t_id = self.target_thread_id()
        out: List[str] = []
        for it in self.selected_source_items():
            tid = str(it.get("thread_id") or "").strip()
            if tid and tid != t_id:
                out.append(tid)
        return out
