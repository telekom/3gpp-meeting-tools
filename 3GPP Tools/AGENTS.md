# AGENTS.md — 3GPP Tools Developer & AI Agent Guide

Authoritative root guide for human developers and AI coding agents working on **3GPP Meeting Tools & Diagram Converter** (`3GPP Tools/`). Cross-cutting environment rules and architectural constraints live here; module specifics and deep feature specs live in subsystem `AGENTS.md` files and [`docs/`](docs/) so this file stays concise and never truncates in agent context windows.

---

## 1. Project Overview

**3GPP Tools** is a Python/PyQt5 desktop engineering environment for telecommunications engineers and 3GPP delegates:
- **PlantUML & Office Integration:** Live PlantUML IDE (`plantuml.jar`) and conversion to native Microsoft Visio (`.vsdx`) and PowerPoint shapes.
- **Specification Search & Evolution:** SQLite FTS5 trigram search across 3GPP releases, "First Added" tracking, OpenXML multi-part parsing, and Change History / Clause-to-TDoc Attribution.
- **Protocol Inspection & Evolution:** Message & IE diffing for 5GS/EPS NAS, 5GS UE Policies (TS 24.526/24.588/24.555/24.514/24.578), NR/LTE RRC, NGAP/XnAP/F1AP/S1AP, GTPv1-U, PDU Session UP, and PFCP.
- **Meetings, Emails & Work Items:** Offline 3GPP meeting/TDoc sync, SA2 Session Plan Studio, Cross-WG Topic Dossiers, Tracked Email Threads & Unified Inbox, and Work Item tracking.
- **Local AI Intelligence & MCP:** 100% offline Ollama integration (`127.0.0.1:11434`) plus stdio MCP server (`src/mcp_server.py`) & headless JSON CLI (`src/cli_tools.py`).

---

## 2. Environment, Execution & Updates

- **OS & Runtime:** Windows (primary, Win32 COM & shell hooks), Python `>= 3.8` (64-bit Python 3.10–3.14 user-space install, zero admin rights required), Java 11+ (or legacy 8) for PlantUML JAR, optional Microsoft Office (Visio/Word/PowerPoint/Outlook) and local Ollama (`127.0.0.1:11434`).
- **Startup Dependency Check (`src/core/utils/dependency_checker.py`):** All entry points (`main_tools.py`, `cli_tools.py`, `mcp_server.py`) verify `requirements.txt` via `importlib.util.find_spec` before heavy imports (GUI shows clipboard + 1-click pip repair dialog; headless/MCP logs to `sys.stderr` without polluting `stdout`).
- **Manual Continuous GitHub `master` Updater (`src/core/updater/`, `src/core/ui/update_dialog.py`, `src/Update.cmd`):** Never polls automatically at startup; triggered explicitly via Settings / Help / `Ctrl+K` or `.\src\Update.cmd`. Fetches `commits/master/3GPP%20Tools.atom` via `NetworkSession`, stages `<sha>.zip` with ZipSlip checks, `3GPP Tools/` subtree filtering, and `\\?\` extended-length paths, preserves `src/db/*.db*`, `src/config/*_config.json`, `*_user.json`, `useful_links.json`, `export_config.json`, `plantuml.jar`, templates, and `output/`, supports `zip_tracked` (`.app_version.json`) / `zip_manual` / `git` (with `"zip_override": true` fallback), and applies via a detached pure-stdlib `apply_worker.py`.

```powershell
# Launch GUI (auto-detects Python 3.14 / 3.x):
.\src\Launch.cmd

# Direct Python launch:
$env:PYTHONPATH = ".;.\src;$env:PYTHONPATH"
python -m main_tools

# Standalone updater outside GUI:
.\src\Update.cmd
```

---

## 3. Repository Layout & Documentation Map

```text
3GPP Tools/
├── AGENTS.md                      # This root guide (keep <18 KB)
├── README.md                      # User & architectural documentation
├── requirements.txt               # Direct runtime dependencies
├── docs/                          # Deep feature & subsystem specifications
│   ├── LEGACY_FUNCTIONALITIES.md
│   ├── UI_SHELL_AND_STYLING.md
│   ├── SPEC_HISTORY_AND_ATTRIBUTION.md
│   ├── MEETINGS_TDOCS_AND_EXPORTERS.md
│   ├── MEETINGS_SESSION_PLAN.md
│   └── MEETINGS_TOPIC_DOSSIERS.md
├── output/                        # Default export directory
└── src/
    ├── Launch.cmd / Launch_MCP.cmd / Update.cmd
    ├── main_tools.py              # GUI entry point, logging, startup faulthandler
    ├── main_window.py             # Re-export shim -> core/ui/main_window/ (DragDropUI shell mixins; see README.md)
    ├── mcp_server.py              # JSON-RPC 2.0 stdio MCP server (12 domain tools)
    ├── cli_tools.py               # Headless JSON CLI for external agents
    ├── assets/                    # PlantUML JAR, multi-res .ico, bundled assets
    ├── config/                    # JSON preferences, presets, hot-reloadable prompts/
    ├── core/                      # ai/, config/, network/, ui/, updater/, utils/, queue_manager.py
    ├── db/                        # Persistent SQLite DBs (3gpp_data.db, etc.)
    └── modules/                   # Feature subsystems (each with focused AGENTS.md)
```

| Subsystem / Topic | Guide Location |
|---|---|
| UI Shell, Typography Scaling & Styling | [`docs/UI_SHELL_AND_STYLING.md`](docs/UI_SHELL_AND_STYLING.md) |
| Visio & PlantUML Studio (`src/modules/puml2visio`) | [`src/modules/puml2visio/AGENTS.md`](src/modules/puml2visio/AGENTS.md) |
| Word Tools, CR/pCR, Sanitizer & MIP Labels (`src/modules/word_tools`) | [`src/modules/word_tools/AGENTS.md`](src/modules/word_tools/AGENTS.md) |
| Specifications Catalogue (`src/modules/specifications`) | [`src/modules/specifications/AGENTS.md`](src/modules/specifications/AGENTS.md) |
| Spec Search & Evolution Matrix (`src/modules/spec_search`) | [`src/modules/spec_search/AGENTS.md`](src/modules/spec_search/AGENTS.md) |
| Spec Change History, Clause Attribution & Origin CRs | [`docs/SPEC_HISTORY_AND_ATTRIBUTION.md`](docs/SPEC_HISTORY_AND_ATTRIBUTION.md) |
| Protocol Matrix & Inspector (`src/modules/nas`) | [`src/modules/nas/AGENTS.md`](src/modules/nas/AGENTS.md) & [`src/modules/nas/core/parsing/AGENTS.md`](src/modules/nas/core/parsing/AGENTS.md) |
| Meetings & TDocs (`src/modules/meetings`) | [`src/modules/meetings/AGENTS.md`](src/modules/meetings/AGENTS.md) + `docs/MEETINGS_*.md` |
| Emails, Tracked Threads & Inbox (`src/modules/emails`) | [`src/modules/emails/AGENTS.md`](src/modules/emails/AGENTS.md) |
| Research Assistant, MCP & External CLIs (`src/modules/assistant`) | [`src/modules/assistant/AGENTS.md`](src/modules/assistant/AGENTS.md) |
| Work Items (`src/modules/work_items`) | [`src/modules/work_items/AGENTS.md`](src/modules/work_items/AGENTS.md) |

---

## 4. Global Architectural Rules & Constraints

### 4.1. Threading & UI Responsiveness
- **Never block the Qt GUI thread:** Run network requests, scraping, large document parsing, COM automation, and LLM inference in `QThread` subclasses or `QueueManager`.
- **Cooperative Cancellation:** Periodically check `self.isInterruptionRequested()` or `CancellationToken`. Never call `QThread.terminate()` during normal work (it corrupts SQLite locks and C runtime state).
- **Safe Signal Delivery:** Pass results back to UI widgets strictly via Qt signals (`pyqtSignal`).

### 4.2. Win32 COM Automation Safety & `faulthandler` Boundary
- **Apartment Lifecycle:** Background threads using `win32com.client` must call `pythoncom.CoInitialize()` before COM calls and `pythoncom.CoUninitialize()` in `finally`. Always clean up (`app.Quit()`, `del app`) in `try...finally` to avoid orphaned `WINWORD.EXE` / `VISIO.EXE`. Prefer direct `lxml` reading over Word COM when layout rendering is not required.
- **`faulthandler` VEH Pitfall on Windows:** `faulthandler.enable()` installs a Vectored Exception Handler that intercepts SEH exceptions before Python `try...except`, turning handled COM rejections (`0x80010001` `RPC_E_CALL_REJECTED`, `0x8001010A` `RPC_E_SERVERCALL_RETRYLATER`, `0xE06D7363`) into false-positive fatal crash dumps. `faulthandler.enable()` and `dump_traceback_later()` are strictly confined to startup in `main_tools.py`; immediately after `window.show()` succeeds, `faulthandler.disable()` **must** be called on Windows (`os.name == 'nt'`).
- **Busy State, Snapshots & Polling:** Wrap live COM property accesses (`Documents.Count`, `ActiveDocument`, `doc.Name`, `FullName`, `Saved`) in `try...except`, returning `None` or logging `DEBUG` on modal rejections. Handle `Document.SaveCopyAs()` OLE error `0x800a1704` by falling back to `shutil.copy2` (or `doc.SaveAs2(..., AddToRecentFiles=False)` for headless instances). Background polling (`_poll_active_word`) must set `allow_unsaved_snapshot=False` and be gated by `self.isVisible()`.

### 4.3. Local AI & Corporate Proxy Bypassing
- **Strictly Offline & Local:** Never add cloud LLM fallbacks; all in-app inference targets local Ollama (`127.0.0.1:11434`).
- **Proxy Bypass:** Use `core.network.session.get_ai_session()` (sets `trust_env=False` for RFC 1918 loopback via `is_local_address()`) so corporate proxies (Zscaler/BlueCoat) never intercept local traffic.
- **Hot-Reloadable Prompts:** Store prompt templates in `src/config/prompts/` (`mtime`-reloaded).

### 4.4. Database Integrity & Multi-Store Isolation
Domain data is split across 5 SQLite stores in `src/db/`:
1. `3gpp_data.db`: Meetings, TDocs, specification catalogue, work items, spec user tags/notes.
2. `3gpp_protocol_data.db`: Protocol models, messages, IE hierarchies.
3. `3gpp_spec_search.db`: Extracted clause text + FTS5 trigram index.
4. `3gpp_spec_history.db`: CR/p-CR ledgers, TDoc cross-refs, `tdoc_metadata_cache`.
5. `user_tdocs.db`: Personal TDoc notes, review status, triage archives.

- **CRITICAL — `src/db/3gpp_data.db` Protection:** **NEVER execute `git checkout`, `git restore`, `git reset`, or `git clean` against `src/db/3gpp_data.db` (or any `.db` in `src/db/`).** It holds live runtime synchronizations and has Git `skip-worktree` set (`git update-index --skip-worktree src/db/3gpp_data.db`); never unset `skip-worktree` or force-restore it. Always write backward-compatible programmatic SQLite migrations (`ALTER TABLE` / column inspection).
- **No Cross-DB `ATTACH`:** Do not couple queries across these DBs with SQLite `ATTACH` outside isolated maintenance scripts. Compaction (`PRAGMA wal_checkpoint(TRUNCATE)`, `VACUUM`) lives in `DatabaseMaintenanceDialog`.

### 4.5. Logging & Diagnostics
- Use `logging.getLogger(__name__)` (mirrored to `3gpp_tools.log` and `ConsolePanel` via `GuiLogHandler`). Log lengths/IDs rather than multi-megabyte raw clause payloads. In `mcp_server.py` and `cli_tools.py`, route all logs to `3gpp_tools.log` and `sys.stderr` so `sys.stdout` remains pure JSON.

### 4.6. UI Shell, Scaling, Design Tokens & Hex Governance
Full component architecture lives in [`docs/UI_SHELL_AND_STYLING.md`](docs/UI_SHELL_AND_STYLING.md). Mandatory rules:
- **Design Tokens (`core.ui.theme`):** Use `C`, `T`, `R`, `S`, `B` and shared styles in `core.ui.ui_components` (`BUTTON_STYLE_TOOLBAR_SECONDARY`, `CARD_FRAME_STYLE`, `BADGE_STYLE_*`, `TABLE_STYLE_CLEAN`). Never introduce raw `#RRGGBB` literals in new/migrated files (enforced by `tests/test_ui_hex_lint.py`; preview migrations via `py -3.14 -m core.ui.hex_migrator <path> --dry-run --diff`). Never use 8-digit `#RRGGBBAA` hex codes in Qt QSS or HTML emails (Qt parses `#AARRGGBB` and Outlook ignores them).
- **Logical-Pixel Scaling (`core.ui.font_scale`):** Scale fixed dimensions, row heights, and delegate `QFont`s strictly in logical pixels via `scale_px()`, `scaled_font()`, and `register_scaled_row_height()` — never `setPointSize()` / `scale_pt()` / `QFont(..., pt)`. For bold/family-only static table/tree items, omit `pixelSize` so Qt inherits the live-scaled QSS `font-size`. Prefer `setMinimumHeight(scale_px(...))` over rigid `setFixedHeight(...)`.
- **Scoped Filter Containers & Checkboxes:** Use `.QWidget { background: transparent; border: none; }` (or explicit `#objectName` selectors) so wrapper styles never cascade into `QComboBox` popups (`#checkableComboPopupView` / `CheckableComboItemDelegate`). Define `:checked` and `:checked:hover` on checkable buttons.
- **Responsive Shell Preservation:** Keep `DragDropUI.setMinimumSize(780, 560)` and `self.tabs` (`.tabs.` API) intact; use `notify_success()` / `notify_info()` Smart Toasts instead of modal `QMessageBox.information` success popups.

### 4.7. Qt UI Text & Ampersand Mnemonic Escaping
- **Mnemonic-Aware Controls (`QPushButton`, `QToolButton`, `QCommandLinkButton`, `QCheckBox`, `QRadioButton`, `QGroupBox`, `QAction`, `QMenu`, `QMenuBar`, `QTabBar`, `QTabWidget`):** Qt treats a single `&` as an accelerator prefix. Always escape literal ampersands as `&&` in static strings (`QPushButton("Save && Close")`, `menu.addAction("Tags && Notes")`), and sanitize dynamic external strings with idempotent `escape_qt_mnemonic(text)` from `core.ui.ui_components`.
- **Non-Mnemonic Widgets (`QLabel`, `QLineEdit`, `QTextEdit`, `QTextBrowser`, `QToolTip`, `QStatusBar`, `setWindowTitle`, `setPlaceholderText`, cells, HTML):** Never use `&&` here (it renders literally as `"&&"`). Always use a single `&` (or `&amp;` in HTML).
- **Mandatory AST Check:** Verify all UI changes with `py -3.14 -m unittest tests/test_ui_ampersand_escaping.py`.

### 4.8. Modeless Dialogs, Win32 Owner Pinning & High-DPI Rules
- **No `Qt.WindowStaysOnTopHint`:** Never pin utility/tool dialogs above external apps (Word, Teams, Outlook).
- **Prevent Win32 `GW_OWNER` Pinning (`super().__init__(None)`):** Modeless tool windows (`SpecClauseExcerptDialog`, `CrPcrDialog`, `DatabaseMaintenanceDialog`, `DatabaseInspectorDialog`, `ProcessManagerDialog`, `TDocEmailsDialog`, `StandaloneEmailReaderWindow`, `EmailSummaryDialog`, `TDocTriageDialog`, `CallFlowDialog`, `SessionPlanWindow`, `SpecVersionHistoryDialog`, `SpecTDocReportDialog`, `SpecClauseDiffDialog`, `ContributionReportDialog`, `TopicDossierDialog`, `TopicInputsViewerDialog`, `McpMonitorDialog`, etc.) must initialize with `super().__init__(None)` (connecting `parent.destroyed.connect(self.close)` if needed), set `self.setWindowFlags(Qt.Window)` (not `self.windowFlags() | Qt.Window` since `Qt.Dialog = 0x3`), set `self.setWindowModality(Qt.NonModal)`, launch with `.show()`, `.raise_()`, `.activateWindow()`, and keep persistent references on the spawner (`destroyed.connect` cleanup).
- **Clean Worker Shutdown & Taskbar Icon:** Guard `closeEvent` threads with `getattr(self, "_worker", None)` + `requestInterruption()`, `wait(500)` (falling back to `terminate()` only on shutdown if still running). Keep `src/assets/3gpp_tools.ico` multi-resolution (`16..256` px) and preserve `install_window_icon_manager(app)` (`tests/test_taskbar_icon.py`).
- **High-DPI Sizing:** Never set tight fixed pixel widths (e.g. `setFixedWidth(72)`) on icon+text buttons; use padding, `min-width`, generous action columns (`>= 280px`), and `min-height: 26-28px` / `46-48px` row heights so descenders (`p`, `g`, `y`, `q`, `j`) never clip.

### 4.9. 3GPP Portal & TDoc URI Construction
- Always use `core.utils.portal_utils` (or alias `core.utils.portal_urls`) instead of raw ad-hoc strings.
- `/ngppapp/CreateTdoc.aspx?mode=view` requires `contributionUid=<base_tdoc>` — **never `tdocId`**. Use `extract_base_tdoc()` / `get_portal_tdoc_url(..., base_only=True)` and prefer `open_portal_tdoc()`, `open_portal_meeting()`, `open_portal_create_tdoc()`, `open_portal_work_item()`.

### 4.10. Protected Drop-In Files
- **`src/modules/meetings/core/exporters/tdoc2md.py`** is a standalone stdlib-only core synced with `3gpp-tdoc-markdown-exporter`. Keep Qt/Office COM/app adapters in `tdoc_markdown_converter.py` or calling modules, and preserve 100% stdlib compatibility when enhancing `tdoc2md.py`.

### 4.11. Modularity, File-Size Limits & Anti-God-Class Rules
- **Size Budgets (`tests/test_code_modularity_lint.py`):** Keep new/refactored Python files in `src/` under **500 lines** (hard ratchet: **<= 600 lines** for new files, **<= 1,000 lines** hard ceiling for all non-protected files; never grow entries in `LEGACY_LARGE_FILE_ALLOWLIST`).
- **Prevent God Classes:** Split large dialogs/windows into `dialog.py` + focused `mixins/` (`ui_setup.py`, data/filtering, actions/export) and large engines into acyclic leaf-to-orchestrator packages (`models`, `persistence`, `planning`, `formatters`) with a thin backward-compatible re-export shim + package `README.md`.

---

## 5. Development, Documentation & Local Commit Workflow

1. **Verify Imports & Tests:** Run with `PYTHONPATH=".;src"`:
   - UI ampersand lint: `python -m unittest tests/test_ui_ampersand_escaping.py`
   - UI hex token lint: `python -m unittest tests/test_ui_hex_lint.py`
   - Modularity & `AGENTS.md` size lint: `python -m unittest tests/test_code_modularity_lint.py`
   - Taskbar icon / Session Plan / targeted unit tests as applicable (`tests/test_taskbar_icon.py`, `tests/test_session_plan_ui_and_html.py`, etc.).
   - When modifying GUI tabs/workers, verify `python -m main_tools` starts and closes cleanly without tracebacks or leaked threads.
2. **Preserve Unrelated Comments/Docstrings** in untouched code.
3. **Consult Subsystem Guides First:** Read the target module's `AGENTS.md` (and `docs/*.md` if applicable) before editing.
4. **Docs-as-Code Placement Rule (Prevent Root Bloat):** Synchronize documentation in the same work unit, placing content at the right level:
   - **Root `AGENTS.md` (keep <18 KB):** Only cross-cutting rules, environment/launch changes, or new subsystem links.
   - **Module `AGENTS.md` (keep <18 KB):** Module directory layout, invariants, public APIs, and backward-compat aliases.
   - **`docs/*.md` & `README.md`:** Detailed feature specs, UI walkthroughs, parser quirks, and user documentation.
5. **Mandatory Autonomous Local Commit (Never Wait for Prompt, Never `git push`):**
   - After verification and doc updates, immediately stage the changed files belonging to your work unit and create a clean local conventional commit (`feat(...)`, `fix(...)`, `refactor(...)`, `docs(...)`).
   - **NEVER execute `git push`** — pushing to remote is strictly reserved for the human developer.
