# --- File: modules/meetings/core/compare_manager.py ---
import logging
from typing import Optional, Tuple
from PyQt5.QtCore import QObject, pyqtSignal

logger = logging.getLogger(__name__)


class ComparisonManager(QObject):
    _instance = None
    # Emits whenever the cart changes: (slot_a_dict, slot_b_dict)
    cart_updated = pyqtSignal(dict, dict)

    @classmethod
    def get_instance(cls) -> "ComparisonManager":
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def __init__(self):
        super().__init__()
        self.slot_a = {}
        self.slot_b = {}
        self._next_slot = "A"

    def add_to_cart(
        self,
        tdoc_name: str,
        file_path: str,
        slot: Optional[str] = None,
        extra_meta: Optional[dict] = None,
    ):
        """
        Pushes a document into the specified slot ('A' or 'B'), or into the first
        empty slot if one is free, otherwise alternating A and B.
        """
        item = {"name": tdoc_name, "path": str(file_path)}
        if extra_meta:
            item.update(extra_meta)

        target = (slot or "").strip().upper()
        if target not in ("A", "B"):
            if not self.slot_a:
                target = "A"
            elif not self.slot_b:
                target = "B"
            else:
                target = self._next_slot

        if target == "A":
            self.slot_a = item
            self._next_slot = "B"
        else:
            self.slot_b = item
            self._next_slot = "A"

        self.cart_updated.emit(self.slot_a, self.slot_b)

    def set_slot_a(self, name: str, path: str, **meta):
        self.add_to_cart(name, path, slot="A", extra_meta=meta)

    def set_slot_b(self, name: str, path: str, **meta):
        self.add_to_cart(name, path, slot="B", extra_meta=meta)

    def clear_slot(self, slot: str):
        s = (slot or "").strip().upper()
        if s == "A":
            self.slot_a = {}
            self._next_slot = "A"
        elif s == "B":
            self.slot_b = {}
            self._next_slot = "B"
        self.cart_updated.emit(self.slot_a, self.slot_b)

    def swap_slots(self):
        """Swap Slot A (Original) and Slot B (Revised) to invert the redline direction."""
        if not self.slot_a and not self.slot_b:
            return
        self.slot_a, self.slot_b = self.slot_b, self.slot_a
        self.cart_updated.emit(self.slot_a, self.slot_b)

    def clear_cart(self):
        self.slot_a = {}
        self.slot_b = {}
        self._next_slot = "A"
        self.cart_updated.emit(self.slot_a, self.slot_b)

    def stage_active_word_document(self, slot: Optional[str] = None) -> Optional[dict]:
        """
        Detects the currently active Microsoft Word document (snapshotting unsaved edits
        if needed) and stages it into the comparison cart.
        """
        from modules.word_tools.core.document_opener import get_active_word_document

        active_doc = get_active_word_document(allow_unsaved_snapshot=True)
        if not active_doc or not active_doc.get("path"):
            return None

        display_name = active_doc.get("display_name") or f"🪟 {active_doc['name']}"
        self.add_to_cart(
            display_name,
            active_doc["path"],
            slot=slot,
            extra_meta={
                "is_open_word": True,
                "word_identifier": active_doc.get("full_name") or active_doc.get("name"),
                "raw_name": active_doc.get("name"),
                "is_snapshot": bool(active_doc.get("is_snapshot")),
            },
        )
        return active_doc

    def stage_open_word_document(self, identifier: str, slot: Optional[str] = None) -> Optional[dict]:
        """
        Stages a specific open Microsoft Word document (by FullName or Name) into the cart.
        """
        from modules.word_tools.core.document_opener import get_active_word_document

        doc_info = get_active_word_document(
            allow_unsaved_snapshot=True,
            target_identifier=identifier,
        )
        if not doc_info or not doc_info.get("path"):
            return None

        display_name = doc_info.get("display_name") or f"🪟 {doc_info['name']}"
        self.add_to_cart(
            display_name,
            doc_info["path"],
            slot=slot,
            extra_meta={
                "is_open_word": True,
                "word_identifier": doc_info.get("full_name") or doc_info.get("name"),
                "raw_name": doc_info.get("name"),
                "is_snapshot": bool(doc_info.get("is_snapshot")),
            },
        )
        return doc_info

    def resolve_comparison_paths(self) -> Tuple[Optional[str], Optional[str]]:
        """
        Returns (path_a, path_b) ready for WordComparatorThread.
        If either slot references a live open Word document, refreshes its snapshot
        from Word so any edits made after staging are included in the diff.
        """
        from modules.word_tools.core.document_opener import get_active_word_document

        def _resolve_slot(slot_dict: dict) -> Optional[str]:
            if not slot_dict:
                return None
            if slot_dict.get("is_open_word"):
                ident = slot_dict.get("word_identifier") or slot_dict.get("raw_name")
                fresh = get_active_word_document(
                    allow_unsaved_snapshot=True,
                    target_identifier=ident,
                )
                if fresh and fresh.get("path"):
                    slot_dict["path"] = fresh["path"]
                    return fresh["path"]
            return slot_dict.get("path")

        return _resolve_slot(self.slot_a), _resolve_slot(self.slot_b)
