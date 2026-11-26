# AGENTS.md — 3GPP Tools Developer & AI Agent Guide

Authoritative root guide for human developers and AI agents working on **3GPP Tools** (`3GPP Tools/`). Cross-cutting rules live here; module specifics and deep feature specs live in subsystem `AGENTS.md` files and [`docs/`](docs/) so this file stays concise and never truncates in agent context windows.

---

## 1. Project Overview

Python/PyQt5 desktop engineering suite for 3GPP delegates:
- **PlantUML & Office:** Live PlantUML IDE (`plantuml.jar`) and conversion to native Visio (`.vsdx`) and PowerPoint shapes.
- **Specs Search & Evolution:** SQLite FTS5 trigram search across releases, "First Added" tracking, OpenXML parsing, Change History, and Clause-to-TDoc Attribution.
- **Protocol Inspector:** Message & IE diffing for 5GS/EPS NAS, 5GS UE Policies (TS 24.526/24.588/24.555/24.514/24.578), NR/LTE RRC, NGAP/XnAP/F1AP/S1AP, GTPv1-U, PDU Session UP, and PFCP.
- **Meetings, Emails & WIs:** Offline meeting/TDoc sync, SA2 Session Plan Studio, Cross-WG Topic Dossiers, Tracked Email Threads & Unified Inbox, and Work Item tracking.
- **Local AI & MCP:** 100% offline Ollama (`127.0.0.1:11434`), stdio MCP server (`src/mcp_server.py`), and headless JSON CLI (`src/cli_tools.py`).

---

## 2. Environment, Execution & Updates

- **Runtime:** Windows primary (Win32 COM & shell hooks), Python `>= 3.8` (64-bit 3.10–3.14 user-space, no admin required), Java 11+ (or 8) for PlantUML JAR, optional MS Office (Visio/Word/PowerPoint/Outlook) and Ollama (`127.0.0.1:11434`).
- **Startup Dependency Check (`src/core/utils/dependency_checker.py`):** `main_tools.py`, `cli_tools.py`, and `mcp_server.py` verify `requirements.txt` via `importlib.util.find_spec` before heavy imports (GUI offers 1-click pip repair; headless/MCP logs to `sys.stderr` only).
- **Manual GitHub `master` Updater (`src/core/updater/`, `update_dialog.py`, `src/Update.cmd`):** Never auto-polls at startup (triggered via Settings/Help/`Ctrl+K` or `Update.cmd`). Fetches `commits/master/3GPP%20Tools.atom` via `NetworkSession`, stages `<sha>.zip` (ZipSlip checks, `3GPP Tools/` subtree filter, `\\?\` paths), preserves DBs (`src/db/*.db*`), user configs (`*_config.json`, `*_user.json`, `useful_links.json`, `export_config.json`), `plantuml.jar`, templates, and `output/`, supports `zip_tracked` (`.app_version.json`)/`zip_manual`/`git` (`"zip_override": true`), and applies via detached stdlib `apply_worker.py`.

```powershell
.\src\Launch.cmd                                     # Launch GUI (auto-detects Python 3.14 / 3.x)
$env:PYTHONPATH = ".;.\src;$env:PYTHONPATH"; python -m main_tools   # Direct launch
.\src\Update.cmd                                     # Standalone updater
```

---

## 3. Repository Layout & Documentation Map

```text
3GPP Tools/
├── AGENTS.md / README.md / requirements.txt
├── docs/              # LEGACY_FUNCTIONALITIES, UI_SHELL_AND_STYLING, SPEC_HISTORY_AND_ATTRIBUTION, MEETINGS_*.md
├── output/            # Default export directory
└── src/
    ├── Launch.cmd / Launch_MCP.cmd / Update.cmd
    ├── main_tools.py  # GUI entry point, logging, startup faulthandler
    ├── main_window.py # Shim -> core/ui/main_window/ (DragDropUI shell mixins; see README.md)
    ├── mcp_server.py  # JSON-RPC 2.0 stdio MCP server (12 domain tools)
    ├── cli_tools.py   # Headless JSON CLI for external agents
    ├── assets/        # plantuml.jar, multi-res 3gpp_tools.ico
    ├── config/        # JSON prefs, presets, hot-reloadable prompts/
    ├── core/          # ai/, config/, network/, ui/, updater/, utils/, queue_manager.py
    ├── db/            # Persistent SQLite DBs (3gpp_data.db, etc.)
    └── modules/       # Feature subsystems (each with focused AGENTS.md)
```

| Subsystem / Topic | Guide |
|---|---|
| UI Shell, Scaling & Styling | [UI_SHELL_AND_STYLING.md](docs/UI_SHELL_AND_STYLING.md) |
| Visio & PlantUML Studio | [puml2visio/AGENTS.md](src/modules/puml2visio/AGENTS.md) |
| Word Tools, CR/pCR, Sanitizer & MIP Labels | [word_tools/AGENTS.md](src/modules/word_tools/AGENTS.md) |
| Specifications Catalogue | [specifications/AGENTS.md](src/modules/specifications/AGENTS.md) |
| Spec Search & Evolution Matrix | [spec_search/AGENTS.md](src/modules/spec_search/AGENTS.md) |
| Spec Change History, Clause Attribution & Origin CRs | [SPEC_HISTORY_AND_ATTRIBUTION.md](docs/SPEC_HISTORY_AND_ATTRIBUTION.md) |
| Protocol Matrix & Inspector | [nas/AGENTS.md](src/modules/nas/AGENTS.md) & [parsing/AGENTS.md](src/modules/nas/core/parsing/AGENTS.md) |
| Meetings & TDocs | [meetings/AGENTS.md](src/modules/meetings/AGENTS.md) + `docs/MEETINGS_*.md` |
| Emails, Tracked Threads & Inbox | [emails/AGENTS.md](src/modules/emails/AGENTS.md) |
| Research Assistant, MCP & External CLIs | [assistant/AGENTS.md](src/modules/assistant/AGENTS.md) |
| Work Items | [work_items/AGENTS.md](src/modules/work_items/AGENTS.md) |

---

## 4. Global Architectural Rules & Constraints

### 4.1. Threading & UI Responsiveness
- **Never block the Qt GUI thread:** Run network, scraping, large parsing, COM, and LLM work in `QThread` or `QueueManager`; deliver UI updates strictly via `pyqtSignal`.
- **Cooperative Cancellation:** Poll `isInterruptionRequested()` / `CancellationToken`. Never call `QThread.terminate()` during normal execution (corrupts SQLite locks & C runtime).

### 4.2. Win32 COM Safety & `faulthandler` Boundary
- **Apartment Lifecycle:** Background COM threads must call `pythoncom.CoInitialize()` / `CoUninitialize()` and `app.Quit()` / `del app` in `try...finally` to avoid orphaned `WINWORD.EXE` / `VISIO.EXE`. Prefer `lxml` over Word COM when layout rendering isn't needed.
- **`faulthandler` VEH Pitfall:** On Windows, `faulthandler.enable()` intercepts SEH exceptions before Python `try...except`, turning handled COM rejections (`0x80010001` `RPC_E_CALL_REJECTED`, `0x8001010A` `RPC_E_SERVERCALL_RETRYLATER`, `0xE06D7363`) into fatal dumps. Confine `enable()` / `dump_traceback_later()` strictly to startup in `main_tools.py` and call `faulthandler.disable()` immediately after `window.show()` when `os.name == 'nt'`.
- **Busy State & Snapshots:** Guard live COM properties (`Documents.Count`, `ActiveDocument`, `Name`, `FullName`, `Saved`) in `try...except`. On `SaveCopyAs()` OLE `0x800a1704`, fall back to `shutil.copy2` (or `SaveAs2(..., AddToRecentFiles=False)` headless). Background `_poll_active_word` must use `allow_unsaved_snapshot=False` and check `self.isVisible()`.

### 4.3. Local AI & Proxy Bypassing
- **Strictly Offline:** Never add cloud LLM fallbacks; all in-app inference targets local Ollama (`127.0.0.1:11434`).
- **Proxy Bypass & Prompts:** Use `core.network.session.get_ai_session()` (`trust_env=False` for RFC 1918 loopback via `is_local_address()`) to bypass corporate proxies. Store hot-reloadable (`mtime`) prompt templates in `src/config/prompts/`.

### 4.4. Database Integrity & Multi-Store Isolation
5 isolated SQLite stores in `src/db/`:
1. `3gpp_data.db`: Meetings, TDocs, spec catalogue, work items, spec tags/notes.
2. `3gpp_protocol_data.db`: Protocol models, messages, IE hierarchies.
3. `3gpp_spec_search.db`: Extracted clause text + FTS5 trigram index.
4. `3gpp_spec_history.db`: CR/p-CR ledgers, TDoc cross-refs, `tdoc_metadata_cache`.
5. `user_tdocs.db`: Personal TDoc notes, review status, triage archives.

- **CRITICAL — `src/db/*.db*` Protection:** **NEVER run `git checkout`, `git restore`, `git reset`, or `git clean` on `src/db/3gpp_data.db` (or any `.db` in `src/db/`).** It has `skip-worktree` set (`git update-index --skip-worktree src/db/3gpp_data.db`) — never unset or force-restore it. Always use backward-compatible programmatic migrations (`ALTER TABLE` / column inspection).
- **No Cross-DB `ATTACH`:** Keep stores decoupled outside isolated maintenance scripts; compaction (`wal_checkpoint(TRUNCATE)`, `VACUUM`) lives in `DatabaseMaintenanceDialog`.

### 4.5. Logging & Diagnostics
- Use `logging.getLogger(__name__)` (mirrored to `3gpp_tools.log` & `ConsolePanel` via `GuiLogHandler`). Log lengths/IDs, never multi-MB clause payloads. In `mcp_server.py` and `cli_tools.py`, log exclusively to `3gpp_tools.log` and `sys.stderr` so `sys.stdout` stays pure JSON.

### 4.6. UI Shell, Scaling, Design Tokens & Hex Governance
Full spec in [UI_SHELL_AND_STYLING.md](docs/UI_SHELL_AND_STYLING.md). Mandatory rules:
- **Design Tokens (`core.ui.theme`):** Use `C`, `T`, `R`, `S`, `B` and shared `core.ui.ui_components` styles (`BUTTON_STYLE_TOOLBAR_SECONDARY`, `CARD_FRAME_STYLE`, `BADGE_STYLE_*`, `TABLE_STYLE_CLEAN`). Zero raw `#RRGGBB` literals in new/migrated UI files (`tests/test_ui_hex_lint.py`; preview via `py -3.14 -m core.ui.hex_migrator <path> --dry-run --diff`). Never use 8-digit `#RRGGBBAA` in QSS or HTML emails (Qt parses `#AARRGGBB`; Outlook ignores them).
- **Logical-Pixel Scaling (`core.ui.font_scale`):** Use `scale_px()`, `scaled_font()`, and `register_scaled_row_height()` — never `setPointSize()` / `scale_pt()` / `QFont(..., pt)`. Omit `pixelSize` on static bold/family table/tree items so they inherit live-scaled QSS. Prefer `setMinimumHeight(scale_px(...))` over `setFixedHeight(...)`.
- **Scoped Containers & Shell:** Use `.QWidget { background: transparent; border: none; }` (or `#objectName`) so styles never cascade into `QComboBox` popups (`#checkableComboPopupView` / `CheckableComboItemDelegate`). Define `:checked` and `:checked:hover` on checkable buttons. Preserve `DragDropUI.setMinimumSize(780, 560)` and `self.tabs`; use `notify_success()` / `notify_info()` toasts instead of modal success `QMessageBox`es.

### 4.7. Qt Ampersand Mnemonic Escaping
- **Mnemonic Controls** (`QPushButton`, `QToolButton`, `QCommandLinkButton`, `QCheckBox`, `QRadioButton`, `QGroupBox`, `QAction`, `QMenu`, `QMenuBar`, `QTabBar`, `QTabWidget`): Escape literal `&` as `&&` in static text; wrap dynamic strings in idempotent `escape_qt_mnemonic(text)` (`core.ui.ui_components`).
- **Non-Mnemonic Widgets** (`QLabel`, `QLineEdit`, `QTextEdit`, `QTextBrowser`, `QToolTip`, `QStatusBar`, `setWindowTitle`, `setPlaceholderText`, cells, HTML): Never use `&&` — use single `&` (or `&amp;` in HTML). Verify with `tests/test_ui_ampersand_escaping.py`.

### 4.8. Modeless Dialogs, Win32 Owner Pinning & High-DPI Rules
- **No `Qt.WindowStaysOnTopHint`** on utility/tool dialogs.
- **Prevent Win32 `GW_OWNER` Pinning:** All modeless tool windows (clause excerpts, CR/pCR, DB maintenance/inspector, process/MCP monitors, email readers/summaries, triage, CallFlow, Session Plan, spec history/diff/reports, topic dossiers/inputs, etc.) must call `super().__init__(None)` (optionally `parent.destroyed.connect(self.close)`), `setWindowFlags(Qt.Window)` (never `windowFlags() | Qt.Window`, since `Qt.Dialog = 0x3`), `setWindowModality(Qt.NonModal)`, launch via `.show()`/`.raise_()`/`.activateWindow()`, and retain a persistent spawner reference with `destroyed.connect` cleanup.
- **Clean Shutdown & Taskbar Icon:** In `closeEvent`, check `getattr(self, "_worker", None)` → `requestInterruption()` + `wait(500)` (`terminate()` only as shutdown last resort). Preserve multi-res `src/assets/3gpp_tools.ico` (`16..256` px) and `install_window_icon_manager(app)` (`tests/test_taskbar_icon.py`).
- **High-DPI Sizing:** Avoid tight `setFixedWidth(...)` on icon+text buttons; use padding, `min-width`, action columns `>= 280px`, and `min-height: 26-28px` / `46-48px` rows to prevent descender clipping.

### 4.9. 3GPP Portal & TDoc URIs
- Always use `core.utils.portal_utils` (alias `portal_urls`). `/ngppapp/CreateTdoc.aspx?mode=view` requires `contributionUid=<base_tdoc>` (**never `tdocId`**). Use `extract_base_tdoc()`, `get_portal_tdoc_url(..., base_only=True)`, and `open_portal_*()` helpers.

### 4.10. Protected Drop-In Files
- **`src/modules/meetings/core/exporters/tdoc2md.py`** is a standalone stdlib-only core synced with `3gpp-tdoc-markdown-exporter`. Keep Qt/COM/app adapters in `tdoc_markdown_converter.py` and preserve 100% stdlib compatibility.

### 4.11. Modularity, File-Size Limits & Anti-God-Class Rules
- **Size Budgets (`tests/test_code_modularity_lint.py`):** Target `< 500` lines for new/refactored Python files in `src/` (hard ratchet: `<= 600` lines for new files, `<= 1,000` lines hard ceiling for all non-protected files; never grow `LEGACY_LARGE_FILE_ALLOWLIST`).
- **Prevent God Classes:** Split large UI dialogs/windows into `dialog.py` + `mixins/` (`ui_setup.py`, filtering, actions/export) and large engines into acyclic leaf-to-orchestrator packages (`models`, `persistence`, `planning`, `formatters`) with a thin backward-compatible re-export shim + package `README.md`.

---

## 5. Development, Documentation & Local Commit Workflow

1. **Verify Imports & Tests** (`PYTHONPATH=".;src"`):
   - Lint suite: `python -m unittest tests/test_ui_ampersand_escaping.py tests/test_ui_hex_lint.py tests/test_code_modularity_lint.py tests/test_ruff_lint.py`
   - Static check: prefer `python -m ruff check src tests` (configured in `pyproject.toml`) when `ruff` is installed; fall back to `python -m pyflakes src tests`.
   - Run targeted tests as applicable (`tests/test_taskbar_icon.py`, `tests/test_session_plan_ui_and_html.py`, etc.).
   - For GUI tab/worker changes, verify `python -m main_tools` starts and closes cleanly without tracebacks or leaked threads.
2. **Preserve Unrelated Comments/Docstrings** in untouched code.
3. **Consult Subsystem Guides First:** Read the target module's `AGENTS.md` (and `docs/*.md`) before editing.
4. **Docs-as-Code Placement Rule (Keep `AGENTS.md` Concise):** Sync docs in the same work unit:
   - **Root & Module `AGENTS.md` (lint-enforced `< 16 KB` to stay well under the 18 KB truncation cutoff):** Root = cross-cutting rules, launch/env, subsystem links; Module = directory layout, invariants, public APIs, backward-compat shims.
   - **`docs/*.md` & `README.md`:** Deep feature specs, UI walkthroughs, parser quirks, user docs.
5. **Mandatory Autonomous Local Commit (Never `git push`):** After verification and doc sync, immediately stage only the files in your work unit and commit locally with a conventional message (`feat(...)`, `fix(...)`, `refactor(...)`, `docs(...)`). **NEVER execute `git push`** — pushing is reserved for the human developer.
