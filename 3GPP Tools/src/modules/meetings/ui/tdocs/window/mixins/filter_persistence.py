# --- File: src/modules/meetings/ui/tdocs/window/mixins/filter_persistence.py ---
"""Per-meeting TDocs filter state persistence and restoration helpers."""

import json
import re

_TDOC_COMBO_KEYS = (
    ("company", "company_combo"),
    ("type", "type_combo"),
    ("ai", "ai_combo"),
    ("sub_item", "sub_item_combo"),
    ("wi", "wi_combo"),
    ("status", "status_combo"),
    ("my_status", "my_status_combo"),
    ("tags", "tags_combo"),
)


class FilterPersistenceMixin:
    """QSettings load/restore/save helpers for `TDocsWindow` filters."""

    def _tdoc_filter_settings_prefix(self) -> str:
        """Build a stable per-meeting QSettings prefix."""
        info = getattr(self, "mtg_info", {}) or {}
        wg = str(info.get("wg_name", "") or "").strip()
        num = str(info.get("meeting_number", "") or "").strip()
        if wg and num:
            raw = f"{wg}_{num}"
        else:
            raw = str(info.get("mtg_id", "") or "").strip() or (
                self.meeting_dir.name if hasattr(self, "meeting_dir") else "default"
            )
        safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", raw).strip("_") or "default"
        return f"meetings/{safe}"

    @staticmethod
    def _parse_filter_str_list(raw_val) -> list:
        """Decode a persisted JSON string array without dropping empty strings."""
        if isinstance(raw_val, list):
            return [str(x).strip() for x in raw_val]
        text = str(raw_val or "").strip()
        if not text:
            return []
        try:
            parsed = json.loads(text)
            if isinstance(parsed, list):
                return [str(x).strip() for x in parsed]
        except Exception:
            pass
        return []

    def _load_filter_settings(self):
        """Load persisted per-meeting filters before building filter widgets."""
        settings = getattr(self, "_filter_settings", None)
        self._saved_filters = {}
        self._pending_combo_filters = {}
        if settings is None:
            return

        prefix = self._tdoc_filter_settings_prefix()
        raw_order = settings.value(f"{prefix}/order_mode", None)
        email_state = str(settings.value(f"{prefix}/email_state", "ALL") or "ALL").strip().upper()
        if email_state not in ("ALL", "UNREAD", "FLAGGED"):
            email_state = "ALL"

        self._saved_filters = {
            "search_text": str(settings.value(f"{prefix}/search_text", "") or ""),
            "followed_wis": settings.value(f"{prefix}/followed_wis", False, type=bool),
            "starred_wis": settings.value(f"{prefix}/starred_wis", False, type=bool),
            "agreements_only": settings.value(f"{prefix}/agreements_only", False, type=bool),
            "email_state": email_state,
            "no_comments": settings.value(f"{prefix}/no_comments", False, type=bool),
            "order_mode": str(raw_order).strip().upper() if raw_order not in (None, "") else None,
            "row_mode": str(settings.value(f"{prefix}/row_mode", "ALL") or "ALL").strip().upper(),
        }

        for key, _attr in _TDOC_COMBO_KEYS:
            if settings.value(f"{prefix}/{key}_filtered", False, type=bool):
                self._pending_combo_filters[key] = self._parse_filter_str_list(
                    settings.value(f"{prefix}/{key}_checked", "[]")
                )

    def _restore_scalar_filters(self):
        """Populate non-combo filter controls from loaded settings without emitting change signals."""
        saved = getattr(self, "_saved_filters", None) or {}
        if not saved:
            return

        if hasattr(self, "search_input"):
            self.search_input.blockSignals(True)
            self.search_input.setText(saved.get("search_text", ""))
            self.search_input.blockSignals(False)

        for attr, flag_key in (
            ("btn_followed_wis", "followed_wis"),
            ("btn_starred_wis", "starred_wis"),
            ("btn_agreements_filter", "agreements_only"),
        ):
            btn = getattr(self, attr, None)
            if btn is not None:
                btn.blockSignals(True)
                btn.setChecked(bool(saved.get(flag_key, False)))
                btn.blockSignals(False)

        email_state = saved.get("email_state", "ALL")
        for btn_attr, target_state in (
            ("btn_unread_emails", "UNREAD"),
            ("btn_flagged_emails", "FLAGGED"),
        ):
            btn = getattr(self, btn_attr, None)
            if btn is not None:
                btn.blockSignals(True)
                btn.setChecked(email_state == target_state)
                btn.blockSignals(False)

        if hasattr(self, "chk_no_comments"):
            self.chk_no_comments.blockSignals(True)
            self.chk_no_comments.setChecked(bool(saved.get("no_comments", False)))
            self.chk_no_comments.blockSignals(False)

        order_mode = saved.get("order_mode")
        if order_mode and hasattr(self, "order_combo"):
            idx = self.order_combo.findData(order_mode)
            if idx >= 0:
                self.order_combo.blockSignals(True)
                self.order_combo.setCurrentIndex(idx)
                self.order_combo.blockSignals(False)

        row_mode = saved.get("row_mode", "ALL")
        if hasattr(self, "row_mode_combo"):
            idx = self.row_mode_combo.findData(row_mode)
            if idx >= 0:
                self.row_mode_combo.blockSignals(True)
                self.row_mode_combo.setCurrentIndex(idx)
                self.row_mode_combo.blockSignals(False)

        self._update_tracked_wi_button_styles()

    def _save_filter_settings(self):
        """Persist current per-meeting TDocs filters to QSettings."""
        if getattr(self, "_restoring_filters", False):
            return
        settings = getattr(self, "_filter_settings", None)
        if settings is None:
            return

        try:
            prefix = self._tdoc_filter_settings_prefix()
            if hasattr(self, "search_input"):
                settings.setValue(f"{prefix}/search_text", self.search_input.text())
            if hasattr(self, "btn_followed_wis"):
                settings.setValue(f"{prefix}/followed_wis", bool(self.btn_followed_wis.isChecked()))
            if hasattr(self, "btn_starred_wis"):
                settings.setValue(f"{prefix}/starred_wis", bool(self.btn_starred_wis.isChecked()))
            if hasattr(self, "btn_agreements_filter"):
                settings.setValue(f"{prefix}/agreements_only", bool(self.btn_agreements_filter.isChecked()))

            email_state = "ALL"
            if hasattr(self, "btn_unread_emails") and self.btn_unread_emails.isChecked():
                email_state = "UNREAD"
            elif hasattr(self, "btn_flagged_emails") and self.btn_flagged_emails.isChecked():
                email_state = "FLAGGED"
            settings.setValue(f"{prefix}/email_state", email_state)

            if hasattr(self, "chk_no_comments"):
                settings.setValue(f"{prefix}/no_comments", bool(self.chk_no_comments.isChecked()))
            if hasattr(self, "order_combo"):
                settings.setValue(f"{prefix}/order_mode", str(self.order_combo.currentData() or "AGENDA"))
            if hasattr(self, "row_mode_combo"):
                settings.setValue(f"{prefix}/row_mode", str(self.row_mode_combo.currentData() or "ALL"))

            pending = getattr(self, "_pending_combo_filters", {}) or {}
            for key, attr in _TDOC_COMBO_KEYS:
                combo = getattr(self, attr, None)
                if combo is not None and combo.model() and combo.model().rowCount() > 1:
                    is_filtered = combo.isFiltered()
                    checked_vals = [str(v).strip() for v in combo.getCheckedItems()] if is_filtered else []
                    settings.setValue(f"{prefix}/{key}_filtered", bool(is_filtered))
                    settings.setValue(f"{prefix}/{key}_checked", json.dumps(checked_vals))
                elif key in pending:
                    settings.setValue(f"{prefix}/{key}_filtered", True)
                    settings.setValue(f"{prefix}/{key}_checked", json.dumps([str(v).strip() for v in pending[key]]))
                else:
                    settings.setValue(f"{prefix}/{key}_filtered", False)
                    settings.setValue(f"{prefix}/{key}_checked", "[]")

            settings.sync()
        except Exception:
            pass
