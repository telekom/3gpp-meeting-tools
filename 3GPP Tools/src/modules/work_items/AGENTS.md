# Work Items Guide (`src/modules/work_items`)

Developer and AI agent guide for the 3GPP Work Items database, parallel portal scrapers, personal watchlist tracking, and clipboard exports.

---

## 1. Storage & Scrapers

- **Database (`3gpp_data.db`):** Stores Work Items across `work_items`, `wi_group_map`, and `wi_remarks`. Never run `git checkout`/`restore`/`reset`/`clean` against `src/db/3gpp_data.db`.
- **Parallel & Targeted Scrapers:**
  - Full parallel scraper syncs across all TSGs/WGs concurrently from dynamic reports.
  - `AddWorkItemDialog` + `ManualWIFetcherThread` scrape individual WIs by ID or 3GPP Portal URL (`extract_wi_code`, `fetch_wi_details_from_portal`, `PORTAL_TO_WG_MAP`, atomic `WorkItemsDatabase.upsert_single_work_item()`).
  - `TargetedWIScraperThread` updates a specific list of WI codes without triggering a full multi-WG scrape.
- **Personal Tracking (`src/config/work_items_config.json`, `WorkItemsSettings`):** Decoupled from `3gpp_data.db` with two independent flags:
  - `starred`: Personal interest / watchlist.
  - `following`: Assigned delegation responsibility (team work-split).

## 2. UI, Filtering & Clipboard (`ui_tabs.py`, `work_items_delegates.py`, `work_items_model.py`, `work_items_dialogs.py`)

- **Modular UI:** `WorkItemsTab` separates `TrackingDelegate`, `work_items_model.py`, filter chips (`⭐ Starred`, `🎯 Following`), `➕ Add WI`, batch context actions, and `WorkItemInfoDialog`. `WorkItemsDatabase` optimizes tracking filters with SQLite primary-key `IN` clauses and safe `1=0` empty-set handling.
- **Quick Copy & Word-Compatible Clipboard:** Context menu, `WorkItemInfoDialog` `📋 Quick Copy ▾`, and `Ctrl+C` (`copy_wis_default`) reuse `_copy_to_clipboard` (`QMimeData` plain text + rich HTML) for WI Acronym links, Portal URLs, Latest WID 3GU Portal links (`core.utils.portal_utils` `contributionUid=...`), citations, Markdown links, and multi-row batch lists.
- **Cross-Tab Navigation:** Preserve `WorkItemsTab.focus_work_item` used by `CommandPaletteDialog` (`Ctrl+K`) and deep links.
