# Session Plan Studio window (`window/`)

`SessionPlanWindow` is composed of feature **mixins**. Each mixin holds verbatim methods of the
former 5,300-line `session_plan_window.py`; they share state via `self` (attributes are created in
`main_window.py → SessionPlanWindow.__init__`). The old module path
`modules.meetings.ui.session_plan.session_plan_window` is kept as a re-export shim.

> **Agents:** open only the file for your concern (table below). Read `main_window.py.__init__`
> if you need the list of shared `self.*` state attributes.

## Where to edit what

| Concern | File |
|---|---|
| Constructor, shared state, Qt lifecycle (`show/close/resize`), `_load_data`, view dispatch, search, timezone / room-location / tracked-topics dialogs | `main_window.py` |
| Constants (`DAY_NAMES`, `QUARTER_*`), `GridRowDescriptor`, snapshot label formatting | `constants.py` |
| Custom-painted grid headers (time-first vertical header, day header, NOW highlight) | `grid_headers.py` |
| Drag & drop, Outlook attachments, import via URL/file/paste, download, diff dialog | `mixins/drop_import.py` |
| Building widgets: top/navigation ribbons, Overview/NOW/Agenda/Drafting/Tracked pages | `mixins/ui_setup_ribbons.py` |
| Building widgets: Team/Delegation page | `mixins/ui_setup_team.py` |
| Grid rows, period timings, day filters, row/column visibility, AI-number helpers, jump to slot | `mixins/grid_model.py` |
| Grid cell cards, slot tooltip, delegate badges, break cells | `mixins/slot_cells.py` |
| Rendering Overview grid, Agenda gap mode, Drafting radar, Tracked topics | `mixins/overview_render.py` |
| Delegation logic: effective delegates, collisions, availability, focus buttons, member CRUD, dispatch, delegation import/export | `mixins/delegation.py` |
| Rendering Team view: roster chips, topic ownership, status cells, clash board | `mixins/team_render.py` |
| NOW view state: live clock, timeline units, context resolution, navigation | `mixins/now_context.py` |
| NOW view rendering: stream cards, upcoming card, compact timeline | `mixins/now_cards.py` |
| Double-click handlers and context menus (grid / drafting / tracked) | `mixins/interactions.py` |
| Exports: mobile HTML, ICS, delegate ICS, e-mail summary | `mixins/exports.py` |

## Rules
- **GUI/HTML parity:** UI changes in `slot_cells.py`, `grid_headers.py`, `now_cards.py` must be mirrored in
  `core/session_plan/session_plan_exporter.py → export_mobile_html`.
- New methods go into the mixin that owns the concern; keep each file below ~600 lines
  (split into a new mixin and add it to `SessionPlanWindow`'s bases otherwise).
- Method names must stay unique across mixins (no MRO overriding between mixins).
- Qt event overrides that call `super()` stay safe because all mixins precede `QWidget` in the bases.
