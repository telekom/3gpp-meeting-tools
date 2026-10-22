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
