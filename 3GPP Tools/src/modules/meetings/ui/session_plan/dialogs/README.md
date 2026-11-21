# Session Plan Dialogs Package (`session_plan/dialogs/`)

Modular decomposition of the 10 SA2 Session Plan modal/modeless dialogs previously housed in the
1,441-line `session_plan_dialogs.py`. `session_plan_dialogs.py` remains a backward-compatible
re-export shim.

## Submodule Navigation Map

| Module | Dialogs / Constants | Responsibility |
|---|---|---|
| [`core_dialogs.py`](core_dialogs.py) | `TIMEZONE_PRESETS`, `SessionPlanTimezoneDialog`, `SessionPlanImportDialog`, `SessionPlanDiffDialog` | Venue timezone selection, URL/File/Paste manual ingestion, and snapshot diff changelog |
| [`slot_and_topic_dialogs.py`](slot_and_topic_dialogs.py) | `SessionPlanNoteDialog`, `SessionPlanRoomLocationsDialog`, `SessionPlanTrackedTopicsDialog` | Slot annotations/badges, `meeting_locations.json` room notes, and tracked topic checklist |
| [`delegation_roster_dialogs.py`](delegation_roster_dialogs.py) | `DelegationMemberDialog`, `DelegationImportDialog` | Delegate member CRUD with palette picker, and Delegation Plan email/JSON import |
| [`delegation_dispatch_dialogs.py`](delegation_dispatch_dialogs.py) | `SessionDispatchDialog`, `MultiDelegateSelectDialog` | Real-time availability-ranked session dispatch and multi-delegate checkboxes |
| [`__init__.py`](__init__.py) | `logger` + all 11 dialog/preset exports | Package exports |

## Invariants
- **Design Tokens:** All stylesheets use `core.ui.theme.C` (`0` raw `#RRGGBB` hex literals across the package and shim).
- Keep every file under **490 lines**.
