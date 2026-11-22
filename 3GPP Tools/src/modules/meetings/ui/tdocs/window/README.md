# TDocs window (`tdocs/window/`)

`TDocsWindow` is composed of feature **mixins**. Each mixin holds verbatim methods of the former
~4,000-line `tdocs_window.py`; they share state via `self` (attributes are created in
`main_window.py → TDocsWindow.__init__` and the `_setup_*` methods). The old module path
`modules.meetings.ui.tdocs_window` is kept as a re-export shim.

> **Agents:** open only the file for your concern (table below). Read `main_window.py.__init__`
> for the shared `self.*` state; widgets are created in `ui_header.py`, `filter_setup.py`, `table_layout.py`.

## Where to edit what

| Concern | File |
|---|---|
| Constructor, shared state, `show/resize/close`, cache setup, `_save_user_data`, count label, context-menu entry, help | `main_window.py` |
| Header/toolbar widgets, resources menu, instant fetch / compare with Word, routing indicator, venue clock | `mixins/ui_header.py` |
| Filter bar widgets, filter-change handlers, filter toggle buttons (agreements, emails, followed WIs) | `mixins/filter_setup.py` |
| Tracked / starred Work Item filter buttons and their context menus | `mixins/filter_tracked_wis.py` |
| Facet counts, combobox population, AI tooltips, clear all filters | `mixins/filter_facets.py` |
| Table construction, row heights, column width fitting / font scaling, header sorting, copy selection | `mixins/table_layout.py` |
| Per-TDoc actions: action/related menus, cell popup, notes, open/reveal/scroll to TDoc, agreements, draft CR, triage | `mixins/tdoc_actions.py` |
| Compose/send TDoc e-mails, email manager, per-TDoc emails, full date-range dialog | `mixins/email_compose.py` |
| Unified email sync (`InboxOutlookSyncWorker`), `🟢 Live Auto-Sync (2m)`, and `📧 Emails ▾` menu builder | `mixins/email_sync.py` |
| Tracked e-mail threads: pin from Outlook/.msg, inbox dialog, table sync, TDoc meta map | `mixins/tracked_threads.py` |
| Reports / statistics, Markdown & LLM export, package & Excel export, export folders | `mixins/exports.py` |
| Refresh Excel / revisions / agenda, WI enrichment, agenda & chair-notes delegates, Session Plan integration | `mixins/meeting_data.py` |
| Drag & drop of files / Outlook items | `mixins/drag_drop.py` |

## Rules
- New methods go into the mixin that owns the concern; keep files below ~600 lines
  (`ui_header.py` is slightly above because `_setup_header` is one 350-line method — split it first if it grows).
- Method names must stay unique across mixins (no overriding between mixins).
- **Tests:** patch module globals in the mixin module that uses them
  (e.g. `modules.meetings.ui.tdocs.window.mixins.exports._open_folder`), not in `tdocs_window`.
