# SA2 Session Plan Studio, Parsing Engine & Delegation Model

Deep architectural reference for SA2 Session Plan management, parsing quirks, 3-state delegation distribution, clash/dispatch board, and 5-view offline HTML/ICS/email interchange (`src/modules/meetings/core/session_plan/`, `src/modules/meetings/ui/session_plan/`).

> [!IMPORTANT]
> **Session Plan GUI ↔ Interactive HTML Export Parity Rule:** `SessionPlanWindow` and `SessionPlanExporter.export_mobile_html` are parallel views of the same meeting timetable. Whenever improving Session Plan visuals, time prominence, conflict indicators, room location notes, or slot card hierarchy in the desktop GUI, keep the standalone interactive HTML export in `session_plan_exporter.py` consistent and verify parity with `tests/test_session_plan_ui_and_html.py`.

---

## 1. Ingestion, Timezones & Snapshot Caching

- **Problem Solved:** SA2 meeting session plans ("Draft time allocation") are published as living TDocs with frequent intra-day revisions (often reusing the same TDoc number e.g. `S2-2610195`), 3 parallel streams, 45-minute half-quarter Time Units (TUs), and ad-hoc drafting sessions (08:00 morning, lunch, 18:00 evening).
- **100% Manual Ingestion & Collision-Safe Caching:**
  - Staged under `<meeting_dir>/Agenda/SessionPlans/` (content-hash naming `S2-2610195_snap2_...docx` prevents overwrites across updates sharing the same TDoc ID) and persisted in SQLite at `<meeting_dir>/Agenda/session_plans.db`.
  - **5 Entry Points:** (1) `🗓️ Session Plan` toolbar button in `TDocsWindow`; (2) `SessionPlanImportDialog` (`http://`/`https://` URL including `10.10.10.10/ftp/...` downloaded via `NetworkSession` with URL-decoded Windows-sanitized filenames into `SessionPlans/Downloaded/`, or TDoc# via `URLRouter`); (3) File browser (`.docx`, `.doc`, `.docm`, `.zip`, `.htm`, `.html`, `.csv`, `.txt`); (4) Drag & Drop onto `SessionPlanWindow` (`DropOverlayWidget` forwarding events, Explorer files, automatic `.zip` decompression into `SessionPlans/Extracted/`, virtual Microsoft Outlook OLE streams `FileGroupDescriptorW`/`FileContents` into `SessionPlans/Dropped/`, and email body / TOHRU table text); (5) TDoc row context menu **`🗓️ Load {tdoc_id} into Session Plan...`**.
- **Timezone & Corporate Laptop Synchronization (`SessionPlanTimezone`, `WorldClockService`):**
  - Strictly anchors internal timestamps in true UTC to survive IT-locked corporate OS clocks (`Europe/Vienna`).
  - **3-Tier Resolution:** (1) Per-meeting `<meeting_dir>/meeting_timezone.json` & `meetings_config.json` `venue_overrides` (`manual_override` locked against auto-overwrite); (2) Public Web API verification (`WebTimezoneSyncWorker` via Open-Meteo Geocoding + `timeapi.io` + physical Geo-IP via `ipwho.is`) and Agenda Deadlines table parsing (`0900 Local time | 0700 UTC`); (3) Expanded offline 3GPP city/country dictionary backed by Python `zoneinfo` + `SessionPlanTimezoneDialog` / `WorldClockDialog` manual overrides and `-1h`/`+1h` DST nudges.
  - Live dual clock pill in toolbar (`clock_lbl` / `#liveClockPill`) updates Venue vs. Laptop time every second.

---

## 2. Fault-Tolerant Session Plan Parsing Engine
(`session_plan_parser.py`, `session_plan_models.py`, `session_plan_db.py`)

Core real-world chair-formatting quirks and their architectural solutions:
1. **Physical Room Overrides (`In main room`, `In Breakout 1`, `In Breakout 2`):** `ROOM_OVERRIDE_RE` isolates and normalizes inline room directives, stores them in `slot.room_override` (SQLite migration `ALTER TABLE session_slots ADD COLUMN room_override TEXT`), and strips them from topic tokens. All UI surfaces and exports resolve `s.room_override or stream_room_name`.
2. **Breakout vs. Coffee Break Disambiguation:** `_detect_break_period` enforces strict word boundary matching (`\bbreak\b`) + quarter context so room header rows (`| Main Room | Breakout 1 | ...`) never spawn phantom coffee breaks.
3. **OpenXML `<w:vMerge>` Deduplication:** Tracks underlying XML cell pointers (`seen_tcs = set()`) across stream rows and break rows so vertically merged plenary cells (`Q0..Q2`, `Q4..Q5`) never spawn phantom duplicate sessions in Streams 2/3, leak `Q0` into `Break 1`, or trigger false delegate double-bookings.
4. **Markdown/HTML `colspan` Row Expansion:** `_process_day_cells` expands `<td colspan=2/>` and 5-cell rows into 10 discrete half-quarter slots, maintaining exact Monday–Friday alignment.
5. **Multi-Topic Slots, Unicode Dash Ranges & Multi-KI Lists:** `_extract_agenda_items`, `_extract_room_override_and_tokens`, and shared `expand_numeric_ai_range()` normalize Unicode dashes (`\u2010`–`\u2015`, `\u2212`, `-`), join unclosed-parenthesis continuation lines, extract **every** `KI#` item and `KI#A–KI#B` range (`(20.6.12 – 16)`, `KI#3, KI#5, KI#11, KI#17?/KI#22?`), populate `ki_map` across ranges, and allocate range slots to the union of delegates following any constituent AI. `SessionPlanDatabase._repair_stale_slot_items()` idempotently repairs legacy cached `slot_agenda_items` rows (like stale `'20.6'`) on snapshot load.
6. **Dynamic `ki_map` Cache:** Populates meeting-wide `KI#N` $\to$ `20.6.N` mappings during traversal to resolve late-week implicit `KI#` references.
7. **Parenthesis Balancing (`_balance_parentheses`):** Strips unmatched outer delimiters after extracting AI numbers while preserving valid topic brackets (`(Net Sharing)`).
8. **Administrative Note Filtering (`ADMIN_NOTE_RE`):** Suppresses logistical scheduling notes (`Drafting 18:30-20:30`, `Session ends at 17:30`, `Revisions (tbd)`) from technical AI lists.
9. **Single-Digit Plenary Items (`single_ai_matches`):** Splits `Opening (1), Agenda (2), Reports (3)` into independent agenda items.
10. **OpenXML Track Changes (`_extract_xml_text`):** Recursively walks `<w:p>` and `<w:tc>` XML trees to capture `<w:ins>` and `<w:moveTo>` insertions while skipping `<w:del>` and `<w:moveFrom>` deletions in revised `r1`/`r2` plans.
11. **Document-Specified Quarter Times (`_resolve_quarter_times`):** Parses the document's time column (including Track Changes) and dynamically computes half-quarter `TU 1` / `TU 2` midpoints (e.g., shifted Q2 `1230 - 1400` $\to$ `12:30–13:15` / `13:15–14:00`), propagating to row headers, tooltips, and HTML exports.
12. **`.x` / `.X` Wildcard Agenda Items (`30.x`, `5.x`, `20.6.X`):** Parses both parenthesized and standalone `.x`/`.X` tokens (including multi-group `Rel-21 (30.x) and Rel-20 (20.x)`), normalizing to lowercase `.x`. Shared `expand_ai_number()` and `expand_slot_ai_numbers()` expand wildcards into concrete `agenda_map` child AIs (preferring strict `<prefix>.*` children over bare chapter headers when children exist), powering `_matching_topic_ownerships()`, `resolve_slot_effective_delegates()`, `get_slot_unassigned_ais()`, `is_slot_not_followed()`, `build_slot_ai_bullets_html()`, `_is_slot_tracked()`, tooltips, context menus, and HTML `UNASSIGNED_AIS` gap analysis; `SessionPlanWindow._combined_ai_keys()` surfaces scheduled wildcard items even when absent from `agenda.csv`.

---

## 3. Desktop `SessionPlanWindow` Architecture & Ergonomics
(`session_plan/window/` package, re-export shim `ui/session_plan_window.py`)

- **Resizable Shell (`minimumSizeHint().width() <= 960 px`):** `FlowLayout` ribbons (`_setup_top_ribbon`, `_setup_navigation_ribbon`, `room_strip`, `delegate_chips_layout`, `ctrl_card`, Team Allocations toolbars) and `_format_compact_snapshot_label` prevent clipping at half-screen width.
- **6 Integrated Views:**
  1. **🗓️ Overview Grid:** 10-column ($5 \text{ days} \times 2 \text{ TU}$) graphical layout with `SessionPlanHorizontalHeader` (2-tier Day Banners + `1st Half`/`2nd Half`, alternating day tints, emerald `Today`), `SessionPlanVerticalHeader` (bold venue-only `Q1 · 09:00 – 10:30`, stream dot, `format_break_period_label()` breaks, live `🟢 NOW`), `SessionPlanGridTable` (`2px` day/quarter dividers, `1px` dashed TU splits), 23 stream/break rows (`☕ Show Breaks`), `Days:` pills (`All 5`, `Today+`, `Mon`–`Fri`), quiet muted `—` for empty `"Drafting:"` (`is_empty_drafting`), unified `QFrame#slotCard` slot cards with top-aligned right-rail initials pills (`JO`, `DI`, `PI`, `❓`, `⛔`, `⚠`), subtle ~7% `DelegationMember.compute_tint` backgrounds, inline `🕒` pills only on half-quarter/custom-timed slots, deferred `resizeRowsToContents()`, and stale `cellWidget` reset (`clearContents()`, `clearSpans()`, `setRowCount(0)`) before rebuilding.
  2. **🔴 Now Live & Upcoming Dashboard (`_setup_now_view`, `_resolve_now_context`):** Auto-follows live venue clock (or `[◀ Prev]`, Slot selector, `[Next ▶]`, `[🔴 Live Now]`), showing side-by-side **Happening Now** stream/room cards with swapped-room badges (`⇄ Moved from ...`), custom location notes (`meeting_locations.json`, `room_locations`, `slot_locations`, `SessionPlanRoomLocationsDialog`, `resolve_effective_room_and_location()` with inline `✏️`), convener/TOHRU metadata, resolved AIs, `2nd Half (TU 2)` lookahead, **⏭️ Up Next Today**, and **📅 Rest of Day Timeline**.
  3. **📋 Agenda Mode (`_render_agenda_mode`, `user_tracked_topics`):** Scheduled vs. unscheduled Gap Analysis with `⭐` 1-click tracking toggle (`SessionPlanTrackedTopicsDialog` + `#TrackedTopicsHeader`); resolves topic titles via `_resolve_ai_description(ai_num, scheduled_ais)` (1. `agenda_map`, 2. matching `AgendaSlotItem.topic_label`, 3. cleaned slot lines). Persistent cell notes/priorities (`user_session_annotations`) are keyed by `(meeting, day, quarter, sub_slot, stream)` across snapshot updates.
  4. **🛠️ Drafting Radar** & **🎯 Delegate Agenda** (`[👥 Whole Delegation | ● Anna | ...]`).
  5. **👥 Team Allocations:** Topic Ownership Matrix (`📅 Scheduled Only (40/222)` toggle backed by `Qt.UserRole` $O(1)$ count, calibrated unclipped columns Col 0 `65px`..Col 4 `185px`, `36px` row height) + 7-column **⚡ Clash & Dispatch Board** (`⚠️ Conflicts (N)` ribbon quick-jump, `🟢 Free Colleagues`, `⚡ Quick-Appoint <Free Colleague> → Cover Str X`, double-click jump to Grid).
- **Window Lifecycle (`closeEvent` / `showEvent`):** Manages `self.clock_timer` and auto-dismisses child `QDialog`s; synchronized with `TDocsWindow.closeEvent`, `MeetingsTab.shutdown()`, and `DragDropUI.closeEvent`.

---

## 4. Delegation Topic Distribution, Three-State Model & Exports

- **Egalitarian Roster & Two-Tier Persistence:** No "myself" tag. **Tier 1 (Default Topic Ownership):** `default_delegation_roster` + `default_topic_owners` in `session_plans.db` synced with `<app_data_root>/delegation_defaults.json` across meetings. **Tier 2 (Session Overrides):** `session_appointments` per meeting/slot, surfaced via `⚡ N session override(s)` and resettable via `↺ Reset Session Overrides` (`clear_session_appointments`).
- **Three-State Delegation Model (`Assigned` / `Unassigned` / `Not Followed`):**
  - `NOT_FOLLOWED_TOKENS` (`NOT_FOLLOWED`, `⛔ NOT FOLLOWED`, etc.) in `session_plan_models.py`; `TopicOwnership.is_not_followed()`, `is_slot_not_followed(slot, plan)`.
  - `DelegateOwnerPicker` (inline `QToolButton` + `QWidgetAction` checkboxes committing on `aboutToHide` and immediately on `_reset_selection` $\to$ `_commit`) supports multi-delegate, `👥 All Team` (`ALL`), `⛔ Mark as Not Followed` (muted slate `#64748B`/`#F1F5F9`), and `✕ Clear (Unassigned)`. Overview Grid right-click provides `👤 Assign Topic Owner (AI X.Y.Z)...`.
  - Not-followed items persist as `owner_name = "NOT_FOLLOWED"`, render `⚪ Not Followed` in the matrix and dashed muted cells in the Grid, and are excluded from `❓ Unassigned AIs` and `#nowDispatchPanel` gap warnings.
- **Boundary-Aware Resolution & Half-Quarter Clash Precision:**
  - `resolve_slot_effective_delegates` & `ai_number_in_text` use `(?<![\d.])` / `(?!\d|\.\d)` lookarounds so `20.6.2` never falsely matches `20.6.24`.
  - `_slots_time_overlap` models `sub_slot=1` (`TU 1`), `sub_slot=2` (`TU 2`), and `0` (full quarter), preventing false clashes on sequential `TU 1` $\to$ `TU 2` slots, while `_compute_delegate_availability` computes real-time `🟢 Free Colleagues`.
- **Solid Pastel Tint Architecture (`DelegationMember.compute_tint`):** 10 WCAG AA hues in `DELEGATE_PALETTE` blended as solid `#RRGGBB` (`bg_tint` 12%, `border_tint` 35% on `#FFFFFF`). **Strictly forbid 8-digit `#RRGGBBAA` hex codes** (Qt QSS parses `#AARRGGBB` and Outlook ignores them).
- **Multi-Format Exports (`SessionPlanExporter`):**
  - **5-View Offline HTML (`SA2_{mtg}_Session_Plan_Mobile.html`, <120 KB, zero CDN):** Sticky 2-row `<header>` on desktop (un-stuck `position: static` on `<= 768px` so filter controls scroll cleanly) + `#liveVenueBanner`, `flex-direction: column` `#overviewView` across all viewports (`100vh` flex above `768px`, dynamic day-count `min-width` + `108px` mobile sticky column + `word-break: normal` so `#liveVenueBanner` always stacks vertically above the grid and filtered single-day columns never squish text vertically) with 2D sticky headers (`#cardsView` auto-default on `<= 768px` with sticky `.day-tabs`), `#nowView` + `#nowDispatchPanel`, `#conflictsView` + `.clash-pill-ring`, `#unassignedView`, `#slotInspectorDrawer` slide-in/bottom sheet, interactive delegate strip (`data-del` attribute binding), client-side Blob `.ics` export, and strict JS/regex escaping (`/[\\r\\n]+/g`, `'\\r\\n'`, `json.dumps()`).
  - **Strict UTC `.ics` (`📅 Export Delegate .ics...`) & Outlook Email Interchange:** `generate_delegation_email_summary` embeds `<!-- 3GPP-DELEGATION-PLAN: {...} -->` JSON for 1-click import via `DelegationImportDialog`.
  - **Version Diffing (`SessionPlanDiffEngine`, `SessionPlanDiffDialog`):** Compares any two snapshots slot-by-slot (`🟢 Added`, `🔴 Removed`, `🟡 Rescheduled/Moved`, `🔵 Volume Changes`).

Agenda rendering shares slot matches across the summary and rows. Tracking changes update existing agenda stars in place while refreshing the dependent views.
