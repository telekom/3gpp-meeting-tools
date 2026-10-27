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
│   ├── tool_registry.py               # Application-owned tool validation & execution boundary
│   ├── knowledge/                     # Read-only service adapters querying application DBs
│   │   ├── protocols.py               # ProtocolKnowledgeService (messages, IEs, hierarchies)
│   │   ├── spec_search.py             # SpecSearchKnowledgeService (ranked text search, clauses)
│   │   └── specifications.py          # SpecificationKnowledgeService (catalogue & versions)
│   └── tools/                         # Semantic tool handlers registered with ToolRegistry
│       ├── factory.py                 # Tool registration factory
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
