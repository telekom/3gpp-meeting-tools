# --- File: src/modules/meetings/ui/delegate_owner_picker.py ---
"""
Checkbox-based multi-delegate owner picker for the Topic Ownership Matrix.

Replaces the former single-choice ``QComboBox`` + "Custom Multi-Delegate..." dialog
round-trip with one in-place popup:

    [ ● Josep, Dieter ▾ ]
      ┌─────────────────────────┐
      │ ☐ 👥 All Team           │
      │ ─────────────────────── │
      │ ☐ (DI) Dieter           │
      │ ☑ (JO) Josep (Me)       │
      │ ☐ (PI) Piotr            │
      │ ─────────────────────── │
      │ ✕ Clear (Unassigned)    │
      └─────────────────────────┘

Checkbox clicks keep the popup open (``QWidgetAction``), so several delegates can be
ticked in one go. The resulting owner string is committed once, when the popup closes,
which avoids re-rendering the whole session plan on every single tick.

Owner string semantics mirror ``MultiDelegateSelectDialog.get_owner_string``:
``""`` (unassigned), ``"ALL"`` (whole delegation), or ``"Name1, Name2"``.
Owners that are not (or no longer) in the roster — e.g. from an imported plan — are
preserved untouched.
"""
from typing import Dict, List, Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import QCheckBox, QMenu, QToolButton, QWidgetAction

from modules.meetings.core.session_plan.session_plan_models import DelegationMember, compute_unique_initials
from core.ui.theme import C

_ALL_TOKENS = ("ALL", "ALL TEAM", "ALL DELEGATES", "👥 ALL TEAM", "👥 ALL")
_NOT_FOLLOWED_TOKENS = (
    "NOT_FOLLOWED",
    "NOT FOLLOWED",
    "NOT_FOLLOW",
    "NOT FOLLOW",
    "⛔ NOT FOLLOWED",
    "⛔ NOT_FOLLOWED",
    "⚪ NOT FOLLOWED",
    "EXCLUDE",
    "IGNORED",
)


class DelegateOwnerPicker(QToolButton):
    """Compact drop-down button with checkable delegates (multi-select) for topic ownership."""

    ownerChanged = pyqtSignal(str)

    def __init__(
        self,
        members: List[DelegationMember],
        owner_string: str = "",
        initials: Optional[Dict[str, str]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self._members = list(members)
        self._initials = initials or compute_unique_initials(m.name for m in self._members)
        self._owner_string = (owner_string or "").strip()
        self._extra_owners: List[str] = []
        self._is_not_followed = False
        self._block_sync = False

        self.setPopupMode(QToolButton.InstantPopup)
        self.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumWidth(130)

        self._menu = QMenu(self)
        self._menu.setStyleSheet(
            f"QMenu {{ background: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; padding: 4px; }}"
            f"QMenu::separator {{ height: 1px; background: {C.BORDER_SUBTLE}; margin: 4px 6px; }}"
            "QMenu::item { padding: 4px 12px; border-radius: 4px; }"
            f"QMenu::item:selected {{ background: {C.INFO_SOFT}; color: {C.INFO_DARK}; }}"
        )
        self.setMenu(self._menu)

        # "All Team" master checkbox
        self.chk_all = QCheckBox("👥 All Team")
        self.chk_all.setStyleSheet(f"QCheckBox {{ font-weight: 700; color: {C.INFO_DARK}; padding: 4px 10px; }}")
        self.chk_all.toggled.connect(self._on_all_toggled)
        self._add_widget_action(self.chk_all)
        self._menu.addSeparator()

        # Individual delegates
        self._checkboxes: Dict[str, QCheckBox] = {}
        for m in self._members:
            ini = self._initials.get(m.name, m.initials)
            chk = QCheckBox(f"({ini})  {m.name}")
            chk.setStyleSheet(f"QCheckBox {{ color: {m.color}; font-weight: 600; padding: 4px 10px; }}")
            chk.toggled.connect(self._on_individual_toggled)
            self._checkboxes[m.name] = chk
            self._add_widget_action(chk)

        self._menu.addSeparator()
        act_not_followed = self._menu.addAction("⛔  Mark as Not Followed")
        act_not_followed.triggered.connect(self._on_not_followed)
        act_clear = self._menu.addAction("✕  Clear (Unassigned)")
        act_clear.triggered.connect(self._on_clear)

        self._menu.aboutToHide.connect(self._commit)
        self._load_owner_string(self._owner_string)
        self._refresh_button()

    # --- Public API -------------------------------------------------------

    def owner_string(self) -> str:
        """Returns the current owner string ('' / 'ALL' / 'A, B')."""
        return self._compute_owner_string()

    def currentText(self) -> str:  # noqa: N802 - QComboBox-compatible name used by table filtering
        """QComboBox-compatible accessor so existing text filters keep working."""
        return self.text().replace(" ▾", "")

    def set_owner_string(self, owner_string: str) -> None:
        """Programmatically sets the selection without emitting ``ownerChanged``."""
        self._owner_string = (owner_string or "").strip()
        self._load_owner_string(self._owner_string)
        self._refresh_button()

    def selected_delegates(self) -> List[str]:
        """Returns the individually selected roster names (plus preserved external owners)."""
        if self.chk_all.isChecked():
            return [m.name for m in self._members]
        return [n for n, c in self._checkboxes.items() if c.isChecked()] + list(self._extra_owners)

    # --- Internals --------------------------------------------------------

    def _add_widget_action(self, widget) -> None:
        act = QWidgetAction(self._menu)
        act.setDefaultWidget(widget)
        self._menu.addAction(act)

    def _load_owner_string(self, owner_string: str) -> None:
        self._block_sync = True
        raw_clean = (owner_string or "").strip()
        self._is_not_followed = raw_clean.upper() in _NOT_FOLLOWED_TOKENS
        is_all = (not self._is_not_followed) and (raw_clean.upper() in _ALL_TOKENS)
        names = [] if (is_all or self._is_not_followed) else [p.strip() for p in raw_clean.split(",") if p.strip()]
        roster = set(self._checkboxes)
        self._extra_owners = [n for n in names if n not in roster]
        for n, chk in self._checkboxes.items():
            chk.setChecked(is_all or n in names)
        self.chk_all.setChecked(is_all)
        self._block_sync = False

    def _on_all_toggled(self, checked: bool) -> None:
        if self._block_sync:
            return
        self._block_sync = True
        self._is_not_followed = False
        for chk in self._checkboxes.values():
            chk.setChecked(checked)
        self._block_sync = False
        self._refresh_button(preview=True)

    def _on_individual_toggled(self) -> None:
        if self._block_sync:
            return
        self._block_sync = True
        self._is_not_followed = False
        all_checked = len(self._checkboxes) > 1 and all(c.isChecked() for c in self._checkboxes.values())
        self.chk_all.setChecked(all_checked and not self._extra_owners)
        self._block_sync = False
        self._refresh_button(preview=True)

    def _reset_selection(self, *, not_followed: bool) -> None:
        self._block_sync = True
        self.chk_all.setChecked(False)
        for chk in self._checkboxes.values():
            chk.setChecked(False)
        self._extra_owners = []
        self._is_not_followed = not_followed
        self._block_sync = False
        # QMenu emits aboutToHide before QAction.triggered; commit explicitly here
        # so menu actions update owner_string and emit ownerChanged immediately.
        self._commit()

    def _on_not_followed(self) -> None:
        self._reset_selection(not_followed=True)

    def _on_clear(self) -> None:
        self._reset_selection(not_followed=False)

    def _compute_owner_string(self) -> str:
        if self._is_not_followed:
            return "NOT_FOLLOWED"
        if self.chk_all.isChecked():
            return "ALL"
        selected = self.selected_delegates()
        if not selected:
            return ""
        if (
            not self._extra_owners
            and len(self._members) > 1
            and len(selected) == len(self._members)
        ):
            return "ALL"
        return ", ".join(selected)

    def _commit(self) -> None:
        new_owner = self._compute_owner_string()
        if new_owner != self._owner_string:
            self._owner_string = new_owner
            self._refresh_button()
            self.ownerChanged.emit(new_owner)

    def _refresh_button(self, preview: bool = False) -> None:
        owner = self._compute_owner_string() if preview else self._owner_string
        members_by_name = {m.name: m for m in self._members}
        is_nf = self._is_not_followed or (owner and owner.upper() in _NOT_FOLLOWED_TOKENS)
        if is_nf:
            text, fg, bg, bdr, weight = "⛔ Not Followed", C.TEXT_MUTED, C.BG_CANVAS, C.BORDER_STRONG, 600
        elif not owner:
            text, fg, bg, bdr, weight = "— Unassigned —", C.TEXT_MUTED, C.BG_CARD, C.BORDER_STRONG, 500
        elif owner.upper() in _ALL_TOKENS:
            text, fg, bg, bdr, weight = "👥 All Team", C.INDIGO_DARK, C.WG_SA_BG, C.WG_SA_BORDER, 700
        else:
            names = [p.strip() for p in owner.split(",") if p.strip()]
            first = members_by_name.get(names[0]) if names else None
            fg = first.color if first else C.TEXT_BODY
            bg = first.bg_tint if first else C.BG_SURFACE
            bdr = first.border_tint if first else C.BORDER_STRONG
            weight = 700
            text = "● " + ", ".join(names) if len(names) <= 2 else f"● {names[0]} +{len(names) - 1}"
        self.setText(f"{text} ▾")
        disp_status = "Not Followed" if is_nf else (owner or "Unassigned")
        self.setToolTip(
            "Click to (multi-)select the topic owner(s).\n"
            f"Current: {disp_status}\n"
            "Changes are applied when the menu closes."
        )
        self.setStyleSheet(
            f"QToolButton {{ text-align: left; padding: 3px 8px; font-size: 11px; font-weight: {weight}; "
            f"color: {fg}; background: {bg}; border: 1px solid {bdr}; border-radius: 6px; }}"
            "QToolButton::menu-indicator { image: none; width: 0px; }"
        )
