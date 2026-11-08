# 3GPP Assistant Subsystem Guide

This document outlines the architecture, authority boundaries, tool registry, and execution rules for the **3GPP Research Assistant** (`src/modules/assistant`). For detailed design rationale and multi-phase specifications, see [`README.md`](file:///src/modules/assistant/README.md).

---

## 1. Subsystem Overview

The Research Assistant is a bounded, evidence-grounded agent architecture connecting a local Ollama LLM to application knowledge services (specifications catalogue, indexed specification text, and structured protocol hierarchies).

> [!IMPORTANT]
> **Strict Authority Boundary: `KNOWLEDGE_READ` Only**
> - The model has **no** arbitrary SQL access, filesystem write access, process/shell execution, or network access.
> - The assistant operates **100% offline and locally** using the local Ollama daemon (`127.0.0.1:11434`). Cloud fallbacks are strictly prohibited.
> - Python retains complete authority over database access, input validation, evidence provenance, and state.

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
│       ├── external_agent_tools.py    # 12 multi-action domain tools for Claude CLI / Codex CLI (MCP & CLI)
│       ├── factory.py                 # Tool registration factory (in-app & external agent profiles)
│       ├── protocol_tools.py          # Protocol inspection tools
│       ├── spec_search_tools.py       # FTS5 search & clause retrieval tools
│       └── specification_tools.py     # Catalogue lookup tools
└── ui/
    ├── assistant_components.py        # Reusable UI components, prompt editor, and style presets
    ├── assistant_tab.py               # Assistant chat interface, evidence drawer, status view
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

The Assistant presentation layer (`src/modules/assistant/ui`) is built on a responsive, card-framed cockpit architecture:

- **Top Control Ribbon:** Features an active Ollama model status pill (`🟢 Ollama: {model}`) with 1-click configuration launcher, session management (`➕ New Chat`, `📋 Copy Chat`, `💾 Export Markdown`), and a layout toggle (`◫ Side Drawer` <-> `⬒ Bottom Split`).
- **Chat Stream & Message Cards:**
  - Distinct styled message cards for user questions (`render_user_message_card`) and grounded assistant answers (`render_assistant_message_card`) with timestamps and KPI metadata (`⏱️ Latency · Steps · Sources`).
  - Formatted technical Markdown rendering (`render_markdown_to_html`) supporting bold text, lists, code blocks, and markdown tables.
  - Interactive clickable evidence citations (`<a href="evidence:E#">E#</a>`) that auto-switch to the Sources tab, select the cited item in the tree, and display its retrieved text in the inspector.
  - Per-message copy action (`📋 Copy Answer`).
  - Empty-state quick starter chips (`StarterChipsWidget`).
- **Multi-Line Prompt Editor (`MultilinePromptEdit`):**
  - Auto-wrapping `QPlainTextEdit` supporting `Enter` to submit and `Shift+Enter` for multiline input, character count label, clear button, and danger-styled cooperative `■ Stop` button.
- **Tabbed Research Cockpit Drawer:**
  - **Tab 1: 📚 Sources & Embedded Inspector (`EvidenceInspectorWidget`):** Upper tree list plus embedded excerpt viewer with source badges (`[Spec Text]`, `[Protocol]`, `[Catalogue]`) and 1-click cross-tab deep-linking:
    - `🔎 In Spec Search`: switches to Spec Search and focuses the cited spec and clause.
    - `🔬 In Protocols`: switches to Protocols and jumps to the message definition.
    - `📚 In Specs`: switches to Specifications and selects the specification.
    - `📋 Copy Excerpt`: copies evidence text to clipboard.
  - **Tab 2: ⚡ Research Activity Timeline (`ActivityTimelineWidget`):** Semantic tool icons, live progress bar, step counter, and citable evidence tags.
  - **Tab 3: 🔬 Diagnostics Trace (`DiagnosticsTraceWidget`):** Real-time KPI strip (latency, prompt tokens, eval tokens, speed in tok/s), filterable raw trace, and 1-click trace copy.

---

## 6. External Agent Integration (Claude CLI & Codex CLI via MCP & Headless CLI)

In addition to the in-app Ollama Research Assistant, the tool registry powers a zero-dependency **Model Context Protocol (MCP) `stdio` server** (`src/mcp_server.py`, `src/Launch_MCP.cmd`) and a **Headless CLI** (`src/cli_tools.py`) exposing 39+ application capabilities to external coding and standards agents such as **Claude CLI (`claude`)** and **Codex CLI (`codex`)**.

### 6.1. Token-Budget Discipline (12 Multi-Action Domain Tools)
To prevent system prompt bloat on every turn in Claude/Codex, all 39+ capabilities are consolidated into **12 compact multi-action domain tools** (`src/modules/assistant/core/tools/external_agent_tools.py`) with `ToolDefinition.mcp_schema()` omitting empty descriptions (<2,000 tokens total for `all`, and 500–850 tokens when scoped via `--profile`):

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
