# 3GPP Assistant Subsystem Guide

Architecture, authority boundaries, tool registry, and execution rules for the **3GPP Research Assistant** (`src/modules/assistant`). Detailed multi-phase design specs live in [`README.md`](README.md).

---

## 1. Subsystem Overview

Bounded, evidence-grounded agent connecting local Ollama (`127.0.0.1:11434`) to read-only knowledge services (specs catalogue, indexed spec text, structured protocols).

> [!IMPORTANT]
> **Strict Authority Boundary: `KNOWLEDGE_READ` Only** — No arbitrary SQL, filesystem writes, shell execution, or network/cloud LLM access. Python retains sole authority over DB access, input validation, evidence provenance, and state.

---

## 2. Directory Structure

```text
src/modules/assistant/
├── AGENTS.md                          # This developer and AI agent guide
├── README.md                          # Comprehensive multi-phase design & architectural specification
├── core/
│   ├── agent_controller.py            # Bounded agent loop, step limits, caching, and state
│   ├── agent_models.py                # Run-local evidence, citation provenance, run results
│   ├── model_types.py                 # Message and conversation context types
│   ├── ollama_agent_adapter.py        # Translates tool schemas & formats Ollama /api/chat requests
│   ├── protocol_registry.py           # Protocol capability descriptor & spec mappings
│   ├── tool_registry.py               # Application-owned tool validation, MCP schemas & execution boundary
│   ├── knowledge/                     # Read-only service adapters querying application DBs
│   │   ├── protocols.py               # ProtocolKnowledgeService (messages, IEs, hierarchies)
│   │   ├── spec_search.py             # SpecSearchKnowledgeService (ranked text search, clauses)
│   │   └── specifications.py          # SpecificationKnowledgeService (catalogue & versions)
│   └── tools/                         # Semantic tool handlers registered with ToolRegistry
│       ├── external_agent/            # Modular 12-tool domain handler package (see external_agent/README.md)
│       ├── external_agent_tools.py    # Backward-compatible re-export shim for external_agent package
│       ├── factory.py                 # Tool registration factory (in-app & external agent profiles)
│       ├── protocol_tools.py          # Protocol inspection tools
│       ├── spec_search_tools.py       # FTS5 search & clause retrieval tools
│       └── specification_tools.py     # Catalogue lookup tools
└── ui/
    ├── tab/                           # Modular AssistantTab package (see tab/README.md)
    ├── assistant_components.py        # Reusable UI components, prompt editor, and style presets
    ├── assistant_tab.py               # Re-export shim -> ui/tab/ package
    └── assistant_worker.py            # AgentResearchWorker (background QThread)
```

---

## 3. Semantic Tool Registry

The LLM is strictly exposed to semantic engineering tools (not database primitives):

| Category | Tool Name | Description |
| :--- | :--- | :--- |
| **Catalogue** | `find_specifications` | Search specifications by title, number, or series |
| | `get_specification` | Retrieve metadata and known releases for a specific spec |
| **Spec Text** | `search_spec_text` | Trigram FTS5 search across indexed text; returns opaque handles |
| | `get_spec_clause` | Fetch complete extracted clause text by handle |
| **Protocols** | `list_protocols` | List supported protocol parsers and imported structured coverage |
| | `list_protocol_messages` | List messages defined for a protocol |
| | `find_protocol_message` | Retrieve message structure and Information Elements |
| | `find_protocol_ie` | Search for IE definitions across protocol messages |

---

## 4. Key Architectural Constraints

1. **Layer Dependency Direction:**
   `UI (AssistantTab) -> Qt Worker (AgentResearchWorker) -> Controller (AgentController) -> Tools -> Knowledge Services -> Databases`
   - Knowledge services and tools must **never** depend on Qt.
   - `OllamaClient` is transport infrastructure only; research policy resides in `AgentController`.
2. **Cooperative Cancellation:**
   - Always respect `CancellationToken`. Check cancellation before and after model and tool invocations.
   - Never call `QThread.terminate()`.
3. **Evidence Provenance:**
   - The application—not the LLM—assigns and tracks evidence identifiers (`E1`, `E2`, etc.) mapped to verified source metadata.
4. **QueueManager Isolation:**
   - Interactive research executes via `AgentResearchWorker` (`QThread`), **not** through `QueueManager`.

---

## 5. UI Architecture & Cockpit Components

Responsive, card-framed cockpit in `src/modules/assistant/ui`:
- **Top Control Ribbon:** Active Ollama status pill (`🟢 Ollama: {model}`) with 1-click config launcher, session actions (`➕ New Chat`, `📋 Copy Chat`, `💾 Export Markdown`), and layout toggle (`◫ Side Drawer` <-> `⬒ Bottom Split`).
- **Chat Stream & Message Cards:** Distinct user (`render_user_message_card`) and grounded assistant (`render_assistant_message_card`) cards with latency/steps/sources KPIs, technical Markdown rendering (`render_markdown_to_html`), clickable evidence citations (`<a href="evidence:E#">E#</a>` auto-selecting the source in the inspector), per-message `📋 Copy Answer`, and `StarterChipsWidget`.
- **Prompt Editor (`MultilinePromptEdit`):** Auto-wrapping `QPlainTextEdit` (`Enter` submit, `Shift+Enter` newline), char counter, clear button, and cooperative `■ Stop`.
- **Cockpit Drawer Tabs:**
  - **Sources (`EvidenceInspectorWidget`):** Tree + excerpt viewer with badges (`[Spec Text]`, `[Protocol]`, `[Catalogue]`) and deep-links (`🔎 In Spec Search`, `🔬 In Protocols`, `📚 In Specs`, `📋 Copy Excerpt`).
  - **Activity (`ActivityTimelineWidget`):** Tool icons, live progress bar, step counter, evidence tags.
  - **Diagnostics (`DiagnosticsTraceWidget`):** KPI strip (latency, prompt/eval tokens, tok/s), filterable raw trace, copy button.
- **Docked Companion Mode (`Ctrl+Shift+A`):** Reparents the live `AssistantTab.body_container` into `RightSidePanel` (leaving `docked_placeholder_card` on tab 9), switches `main_splitter` to `Qt.Vertical` with `dock_compact_strip`, auto-expands `Sources` on `[E#]` clicks, and restores cleanly on tab switch / `Full View` / dock close.

---

## 6. External Agent Integration (MCP & Headless CLI)

The tool registry also powers a zero-dependency **MCP `stdio` server** (`src/mcp_server.py`, `src/Launch_MCP.cmd`) and **Headless CLI** (`src/cli_tools.py`) for external agents (`claude`, `codex`, `gemini`).

### 6.1. Token-Budget Discipline (12 Multi-Action Domain Tools)
All 39+ capabilities are consolidated into **12 domain tools** (`src/modules/assistant/core/tools/external_agent/`, shim `external_agent_tools.py`) with `ToolDefinition.mcp_schema()` omitting empty descriptions (<2,000 tokens for `all`, 500–850 tokens per `--profile`):

| Domain Tool | Actions | Permissions |
| :--- | :--- | :--- |
| `spec_catalogue_and_files` | `search`, `get`, `resolve_files`, `download`, `index_for_search`, `convert_to_md` | `KNOWLEDGE_READ`, `NETWORK_FETCH`, `DOCUMENT_WRITE` |
| `spec_text_and_clauses` | `search_text`, `get_clause`, `evolution_matrix`, `diff_versions`, `extract_clause` | `KNOWLEDGE_READ`, `DOCUMENT_WRITE` |
| `spec_history_and_attribution` | `query_history`, `ingest_history`, `attribute_clause`, `rank_origin_crs`, `tdoc_dossier` | `KNOWLEDGE_READ`, `NETWORK_FETCH` |
| `protocol_inspector` | `list_protocols`, `list_messages`, `find_message`, `find_ie`, `evolution_matrix` | `KNOWLEDGE_READ` |
| `onsite_meeting_cockpit` | `network_status`, `browse_lan_server`, `fetch_live_revisions`, `sync_chair_notes`, `where_should_we_go_now`, `batch_cache_tdocs` | `KNOWLEDGE_READ`, `NETWORK_FETCH`, `DOCUMENT_WRITE` |
| `meetings_and_tdocs` | `search_meetings`, `search_tdocs`, `read_tdoc_markdown`, `get_or_set_stance`, `meeting_stats` | `KNOWLEDGE_READ`, `NETWORK_FETCH`, `DOCUMENT_WRITE` |
| `chair_notes_and_session_plan` | `parse_chair_notes_file`, `ingest_session_plan`, `query_session_plan`, `diff_session_plans` | `KNOWLEDGE_READ`, `DOCUMENT_WRITE` |
| `word_markdown_studio` | `word_to_md`, `md_to_word`, `slice_clauses`, `extract_visio` | `DOCUMENT_WRITE` |
| `cr_pcr_generator` | `generate_pcr`, `generate_cr`, `preview_filename` | `DOCUMENT_WRITE` |
| `word_sanitizer_and_compliance` | `audit_hygiene`, `sanitize_document`, `compare_documents`, `inspect_label`, `apply_label` | `KNOWLEDGE_READ`, `DOCUMENT_WRITE` |
| `plantuml_office_studio` | `validate_or_render_svg`, `export_vsdx`, `export_pptx`, `bridge_pptx_vsdx`, `extract_puml_from_vsdx`, `list_presets` | `KNOWLEDGE_READ`, `DOCUMENT_WRITE` |
| `work_items_and_emails` | `search_work_items`, `fetch_wi_from_portal`, `toggle_wi_tracking`, `search_emails` | `KNOWLEDGE_READ`, `NETWORK_FETCH`, `DOCUMENT_WRITE` |

> [!NOTE]
> **Visio Reverse Engineering Constraint (`extract_puml_from_vsdx`):** Extracting embedded PlantUML source code from `.vsdx` archives only works on Visio files previously generated by **3GPP Tools** (which embed PlantUML metadata into the OpenXML package).

### 6.2. Profiles, Read-Only Mode & On-Site `10.10.10.10` Routing
- **Profiles (`--profile`):** `all` (12 tools), `knowledge` (5 tools), `onsite` (5 tools), `authoring` (5 tools).
- **Read-Only Enforcement (`--read-only`):** Blocks all mutating actions (`DOCUMENT_WRITE`) at both tool registration and action dispatch boundaries while preserving read/inspection actions inside mixed tools.
- **On-Site Meeting Venue LAN (`10.10.10.10` / `3GPPWIFI`):** `onsite_meeting_cockpit` and `URLRouter` automatically probe `10.10.10.10:80` and active Windows Wi-Fi profiles (`3GPPWIFI`), bypass corporate proxies (`no_proxy` includes `10.10.10.10`), skip polite FTP scraper delays on LAN transfers, and access live `/Inbox/`, `/Revisions/`, and `/Inbox/Chair_Notes/` directories unavailable on public 3GPP FTP.
- **Strict `stdout` Isolation & Logging:** In `mcp_server.py` and `cli_tools.py`, all log records are routed exclusively to `3gpp_tools.log` and `sys.stderr` so `sys.stdout` remains 100% pure for JSON-RPC 2.0 / CLI JSON payloads.

### 6.3. MCP Telemetry, Desktop Cockpit & CLI Launcher (`mcp_activity.py`, `core/ui/mcp_monitor/` + `mcp_monitor_dialog.py` shim, `cli_detector.py`, `cli_launcher.py`, `cli_context_dialog.py`)
- **Cross-Process Telemetry (`mcp_activity.py`):** Atomic append-only JSONL journal (`mcp_activity.jsonl` in AppData `logs/`, auto-trimming to latest 1,500 when >2,500 lines), per-PID `session_{pid}.json`, and Win32 kernel liveness checks (`OpenProcess` & `GetExitCodeProcess`).
- **Desktop Monitor Cockpit (`McpMonitorDialog`, non-modal floating window):** Accessible via `SystemStatusBar` (`🤖 MCP: ...`), `ConsolePanel` (`🤖 MCP`), and `Help ▾ → 🤖 MCP Server && AI Agent Cockpit...` with 3 tabs: (1) *Live Activity Stream* (JSON argument/response inspector); (2) *AI Agent Client Setup* (live **Installed LLM CLIs** strip with `[🚀 Launch]`, `[🧠 Models & Company Context...]`, `[🔄 Re-scan]` + config snippets for Claude Desktop, Codex, Claude CLI, Gemini CLI `~/.gemini/settings.json`, Cursor/Windsurf, `Launch_MCP.cmd`); (3) *Interactive Tool Playground* (executes all 12 tools / 58 actions via `ToolRegistry`).
- **Interactive CLI Discovery & Launch:** `cli_detector.py` discovers `codex`, `claude`, `gemini` in `<2 ms` (reading `~/.codex/models_cache.json` + `~/.codex/config.toml` + `src/config/cli_agents_config.json`); `cli_launcher.py` writes `_CLI_CONTEXT.md` aggregating two-layer notes (`src/config/cli_user_context.md` + `_USER_NOTES.md`) and launches inside canonical `{meeting_dir}/{base_tdoc}/` (`stage_copies=False`) for single TDocs or `Export/CLI_Batch_*` for multi-TDoc batches.
