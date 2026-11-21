"""Filter QSettings persistence, responsive layout orientation, column fitting, and lifecycle mixin."""

from __future__ import annotations

import json
from pathlib import Path
import re
from typing import Any, List, Optional

from PyQt5.QtCore import Qt, QTimer

from core.ui.font_scale import scale_px


class InboxLayoutAndSettingsMixin:
    """Manages filter settings persistence, splitter layout, responsive column sizing, and session state."""

    def _meeting_settings_prefix(self) -> str:
        raw = self.meeting_id or "_global"
        safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", raw).strip("_") or "_global"
        return f"meetings/{safe}"

    @staticmethod
    def _parse_json_str_list(raw_val: Any) -> List[str]:
        if isinstance(raw_val, list):
            return [str(x).strip() for x in raw_val if str(x).strip()]
        text = str(raw_val or "").strip()
        if not text:
            return []
        try:
            parsed = json.loads(text)
            if isinstance(parsed, list):
                return [str(x).strip() for x in parsed if str(x).strip()]
        except Exception:
            pass
        return []

    def _load_filter_settings(self):
        prefix = self._meeting_settings_prefix()
        mode = str(self._settings.value(f"{prefix}/active_mode", "UNREAD") or "UNREAD").strip().upper()
        self._active_mode = mode if mode in ("UNREAD", "FLAGGED", "RECENT", "ALL") else "UNREAD"

        raw_vm = self._settings.value(f"{prefix}/view_mode", None)
        if raw_vm is None:
            raw_vm = self._settings.value("view_mode", "MESSAGES")
        self._view_mode = "THREADS" if str(raw_vm or "").strip().upper() == "THREADS" else "MESSAGES"

        src = str(self._settings.value(f"{prefix}/source_filter", "") or "").strip().upper()
        self._saved_source_filter = src if src in ("", "NEEDS_TRIAGE", "THREAD", "UNPINNED") else ""
        self._saved_search_text = str(self._settings.value(f"{prefix}/search_text", "") or "")

        comp_filtered = self._settings.value(f"{prefix}/company_filtered", False, type=bool)
        comp_list = self._parse_json_str_list(self._settings.value(f"{prefix}/company_checked", "[]"))
        self._pending_company_filter = comp_list if (comp_filtered and comp_list) else None

        ai_filtered = self._settings.value(f"{prefix}/ai_filtered", False, type=bool)
        ai_list = self._parse_json_str_list(self._settings.value(f"{prefix}/ai_checked", "[]"))
        self._pending_ai_filter = ai_list if (ai_filtered and ai_list) else None

        status_filtered = self._settings.value(f"{prefix}/status_filtered", False, type=bool)
        status_list = self._parse_json_str_list(self._settings.value(f"{prefix}/status_checked", "[]"))
        self._pending_status_filter = status_list if (status_filtered and status_list) else None

    def _save_filter_settings(self):
        if getattr(self, "_restoring_filters", False):
            return
        try:
            prefix = self._meeting_settings_prefix()
            self._settings.setValue(f"{prefix}/active_mode", self._active_mode)
            self._settings.setValue(f"{prefix}/view_mode", self._view_mode)
            if hasattr(self, "source_combo"):
                self._settings.setValue(f"{prefix}/source_filter", str(self.source_combo.currentData() or ""))
            if hasattr(self, "search_edit"):
                self._settings.setValue(f"{prefix}/search_text", self.search_edit.text())
            if hasattr(self, "company_combo"):
                comp_filtered = self.company_combo.isFiltered()
                comp_checked = (
                    [str(c) for c in self.company_combo.getCheckedItems() if str(c).strip()]
                    if comp_filtered
                    else []
                )
                self._settings.setValue(f"{prefix}/company_filtered", bool(comp_filtered and comp_checked))
                self._settings.setValue(f"{prefix}/company_checked", json.dumps(comp_checked))
            if hasattr(self, "ai_combo"):
                ai_filtered = self.ai_combo.isFiltered()
                ai_checked = (
                    [str(a) for a in self.ai_combo.getCheckedItems() if str(a).strip()]
                    if ai_filtered
                    else []
                )
                self._settings.setValue(f"{prefix}/ai_filtered", bool(ai_filtered and ai_checked))
                self._settings.setValue(f"{prefix}/ai_checked", json.dumps(ai_checked))
            if hasattr(self, "status_combo"):
                st_filtered = self.status_combo.isFiltered()
                st_checked = (
                    [str(s) for s in self.status_combo.getCheckedItems() if str(s).strip()]
                    if st_filtered
                    else []
                )
                self._settings.setValue(f"{prefix}/status_filtered", bool(st_filtered and st_checked))
                self._settings.setValue(f"{prefix}/status_checked", json.dumps(st_checked))
            self._settings.sync()
        except Exception:
            pass

    def _apply_layout_orientation(self, save_current: bool = True):
        if not hasattr(self, "splitter"):
            return
        if save_current:
            prev_key = "splitter_state_v" if self.splitter.orientation() == Qt.Vertical else "splitter_state_h"
            try:
                self._settings.setValue(prev_key, self.splitter.saveState())
            except Exception:
                pass

        is_vertical = self._layout_orientation == "V"
        if is_vertical:
            self.splitter.setOrientation(Qt.Vertical)
            if hasattr(self, "_left_widget"):
                self._left_widget.setMinimumWidth(0)
                self._left_widget.setMinimumHeight(140)
            if hasattr(self, "_right_widget"):
                self._right_widget.setMinimumWidth(0)
                self._right_widget.setMinimumHeight(200)
            self.btn_layout_toggle.setText("◫ Side-by-Side")
            self.btn_layout_toggle.setToolTip("Switch back to Side-by-Side layout (best for wide monitors).")
            saved = self._settings.value("splitter_state_v")
            default_sizes = [260, 440]
        else:
            self.splitter.setOrientation(Qt.Horizontal)
            if hasattr(self, "_left_widget"):
                self._left_widget.setMinimumHeight(0)
                self._left_widget.setMinimumWidth(240)
            if hasattr(self, "_right_widget"):
                self._right_widget.setMinimumHeight(0)
                self._right_widget.setMinimumWidth(280)
            self.btn_layout_toggle.setText("⬒ List on Top")
            self.btn_layout_toggle.setToolTip("Switch to Top/Bottom layout so email subjects span full window width.")
            saved = self._settings.value("splitter_state_h")
            default_sizes = [580, 660]

        restored = False
        if saved:
            try:
                restored = bool(self.splitter.restoreState(saved))
            except Exception:
                restored = False
        if not restored:
            self.splitter.setSizes(default_sizes)
        QTimer.singleShot(0, self._fit_columns_to_viewport)

    def _toggle_layout_orientation(self):
        self._layout_orientation = "H" if self._layout_orientation == "V" else "V"
        self._settings.setValue("layout_orientation", self._layout_orientation)
        self._apply_layout_orientation(save_current=True)

    @staticmethod
    def compute_fitted_column_widths(
        avail_w: int,
        base_widths: List[int],
        mins: Optional[List[int]] = None,
        flex_weights: tuple = (0.03, 0.20, 0.03, 0.48, 0.22, 0.04),
    ) -> List[int]:
        """Proportionally fits interactive column widths to fill `avail_w` without starving Subject or Sender."""
        if avail_w <= 0 or len(base_widths) != 6:
            return list(base_widths)
        min_widths = mins or [
            scale_px(64),
            scale_px(110),
            scale_px(68),
            scale_px(180),
            scale_px(110),
            scale_px(108),
        ]
        clamped = [max(m, int(w)) for w, m in zip(base_widths, min_widths)]
        total_base = sum(clamped)
        if avail_w >= total_base:
            extra = avail_w - total_base
            out = [int(w + round(extra * wt)) for w, wt in zip(clamped, flex_weights)]
        else:
            min_sum = sum(min_widths)
            if avail_w <= min_sum:
                return list(min_widths)
            shrinkable = [max(0, w - m) for w, m in zip(clamped, min_widths)]
            total_shrink = sum(shrinkable)
            if total_shrink <= 0:
                return list(min_widths)
            need = total_base - avail_w
            out = [
                max(m, int(round(w - (need * sk / total_shrink))))
                for w, m, sk in zip(clamped, min_widths, shrinkable)
            ]
        diff = avail_w - sum(out)
        if diff != 0:
            out[3] = max(min_widths[3], out[3] + diff)
        return out

    def _fit_columns_to_viewport(self, avail_width: Optional[int] = None):
        if not hasattr(self, "table"):
            return
        if avail_width is None:
            if not self.isVisible():
                return
            avail_w = int(self.table.viewport().width())
        else:
            avail_w = int(avail_width)
        if avail_w < 300:
            return
        hdr = self.table.horizontalHeader()
        base = list(getattr(self, "_col_base_widths", []) or [hdr.sectionSize(c) for c in range(6)])
        fitted = self.compute_fitted_column_widths(avail_w, base)
        self._restoring_header = True
        try:
            for col_i, col_w in enumerate(fitted):
                hdr.resizeSection(col_i, col_w)
        finally:
            self._restoring_header = False

    def showEvent(self, event):
        super().showEvent(event)
        QTimer.singleShot(0, self._fit_columns_to_viewport)

    def _on_header_section_resized(self, *_args):
        if getattr(self, "_restoring_header", False) or getattr(self, "_restoring_filters", False):
            return
        try:
            hdr = self.table.horizontalHeader()
            self._col_base_widths = [max(40, hdr.sectionSize(c)) for c in range(6)]
            prefix = self._meeting_settings_prefix()
            self._settings.setValue(f"{prefix}/table_header_state", hdr.saveState())
        except Exception:
            pass

    def closeEvent(self, event):
        try:
            self._flush_pending_note_save(reload_view=False)
            if getattr(self, "_sync_worker", None) and self._sync_worker.isRunning():
                self._sync_worker.request_cancel()
                self._sync_worker.wait(1500)
            self._settings.setValue("geometry", self.saveGeometry())
            self._settings.setValue("layout_orientation", self._layout_orientation)
            if hasattr(self, "splitter"):
                key = "splitter_state_v" if self.splitter.orientation() == Qt.Vertical else "splitter_state_h"
                self._settings.setValue(key, self.splitter.saveState())
            if hasattr(self, "table"):
                prefix = self._meeting_settings_prefix()
                self._settings.setValue(f"{prefix}/table_header_state", self.table.horizontalHeader().saveState())
            if hasattr(self, "preview_browser"):
                self._settings.setValue("zoom_steps", int(self.preview_browser.property("_reader_zoom_steps") or 0))
            self._save_filter_settings()
        except Exception:
            pass
        super().closeEvent(event)

    def get_session_state(self) -> dict:
        """Return serializable MeetingEmailInboxDialog context."""
        meeting_dir = ""
        try:
            gdb = getattr(self, "general_email_db", None)
            db_path = getattr(gdb, "db_path", None)
            if db_path:
                meeting_dir = str(Path(db_path).parent.parent)
        except Exception:
            pass
        return {
            "meeting_id": str(getattr(self, "meeting_id", "") or ""),
            "meeting_dir": meeting_dir,
        }
