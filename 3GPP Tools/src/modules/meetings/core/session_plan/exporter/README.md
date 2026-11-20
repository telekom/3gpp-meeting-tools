# Session Plan Exporter (`exporter/`)

Modular replacement for the former 3,726-line `session_plan_exporter.py`.
`modules.meetings.core.session_plan.session_plan_exporter` remains a thin re-export shim exposing `SessionPlanExporter`.

> **Agents:** open only the module for the format or section you need (table below).

## Where to edit what

| Concern | File |
|---|---|
| Public facade (`SessionPlanExporter` methods/signatures) | `facade.py` |
| Shared pure helpers (delegate resolvers, tracked/unassigned checks, period times, locations fallback) | `helpers.py` |
| HTML data prep (`MobileHtmlContext`: timezone, day/quarter maps, unassigned AI catalog, JSON payloads) | `html_data.py` |
| HTML Overview Grid table, slot cells, tooltips, delegate filter buttons | `html_overview.py` |
| HTML Mobile Day Cards view | `html_cards.py` |
| HTML document synthesis (`export_mobile_html`) | `html_exporter.py` |
| Embedded CSS: root tokens, header, dispatch panel, filter strips | `css_layout.py` |
| Embedded CSS: Overview Grid table, sticky headers, slot cards, cell contents | `css_grid.py` |
| Embedded CSS: Mobile Day Cards, Conflicts/Unassigned/Now views, drawer, print | `css_views.py` |
| Embedded JS: UI state, view switching, filters, inspector drawer, client `.ics` | `js_core.py` |
| Embedded JS: clash detection, Conflicts board, Unassigned AIs view | `js_conflicts_unassigned.py` |
| Embedded JS: live "Now" view, dispatch board, dual-clock pill, startup init | `js_now_live.py` |
| RFC 5545 UTC-anchored `.ics` calendar export (`export_ics`) | `ics_exporter.py` |
| Outlook e-mail summaries & delegation plan JSON serialize/parse | `email_exporter.py` |

## Invariants
- **GUI ↔ HTML parity:** UI changes in `ui/session_plan/window/` (`slot_cells.py`, `grid_headers.py`, `now_cards.py`) must be mirrored in `html_overview.py` / `html_cards.py` / `css_*.py` / `js_*.py` and verified with `tests/test_session_plan_ui_and_html.py`.
- **Solid `#RRGGBB` tints only:** never emit 8-digit `#RRGGBBAA` hex codes in HTML/CSS/emails.
- **Python `>= 3.8` f-strings:** keep backslashes out of f-string `{...}` expressions (e.g. `safe_target` hoisted in `ics_exporter.py`).
- **JS escaping:** `js_*.py` strings MUST remain non-raw triple-quoted `"""..."""` (NOT `r"""..."""`) so `\n`, `\r`, and regex escapes preserve their exact runtime characters.
- Keep every file under **600 lines**.
