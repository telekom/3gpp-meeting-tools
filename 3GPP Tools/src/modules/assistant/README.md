# 3GPP Assistant

The `modules.assistant` package provides the application-side foundation for a
local, evidence-grounded 3GPP research assistant.

The assistant is designed to use the knowledge already maintained by **3GPP
Tools** rather than behaving as a standalone chatbot.  A local LLM (initially
through Ollama) may decide which semantic research tools to use, while Python
retains authority over database access, validation, permissions, evidence,
resource limits, and persistent application state.

> **Implementation status**
>
> Stage 1 implements the agent-independent domain models, protocol capability
> registry, and read-oriented knowledge services.  Ollama tool calling, the
> bounded agent controller, Qt worker, and Assistant UI are subsequent stages.

## Goals

The Assistant is intended to:

- answer engineering questions using locally available 3GPP knowledge;
- search indexed specification text and retrieve exact clauses progressively;
- query structured protocol/message/IE knowledge;
- retain source/version/clause provenance;
- distinguish application knowledge coverage from model knowledge;
- make missing structured knowledge observable;
- support future, user-approved knowledge acquisition without allowing an LLM
  to write arbitrary database records.

Phase 1 is deliberately **read-only from the agent's authority perspective**.

## Non-goals for Phase 1

Phase 1 does not give the LLM:

- arbitrary SQL access;
- arbitrary filesystem access;
- arbitrary HTTP/network access;
- shell/process execution;
- Python execution;
- configuration write access;
- database write access;
- autonomous specification downloads or imports;
- permission to rebuild or delete knowledge databases;
- cloud-LLM fallback.

Existing application maintenance functionality remains separate from agent
research.

## Architecture

The target architecture is:

```text
┌──────────────────────────────────────────────────────────────┐
│                         Qt UI                                │
│                      AssistantTab                            │
└─────────────────────────────┬────────────────────────────────┘
                              │
                    AgentResearchWorker
                              │
                              ▼
                    ┌─────────────────┐
                    │ AgentController │
                    └───────┬─────────┘
                            │
              ┌─────────────┴─────────────┐
              ▼                           ▼
      OllamaAgentAdapter              ToolRegistry
              │                           │
        OllamaClient          ┌───────────┼───────────┐
              │               ▼           ▼           ▼
            Qwen          Catalogue   Spec Text    Protocol
                           Tools       Tools        Tools
                              │           │           │
                              ▼           ▼           ▼
                         Knowledge Services
                              │           │           │
                              ▼           ▼           ▼
                         Existing 3GPP Databases
```

Stage 1 implements the bottom part of this diagram:

```text
Protocol Registry
        +
Knowledge Services
        +
Application-owned result/evidence models
```

## Dependency rules

Dependencies should point downward:

```text
UI
 ↓
Qt worker
 ↓
Agent controller
 ↓
Tools
 ↓
Knowledge services
 ↓
Existing application core/database APIs
```

The following boundaries are intentional:

- knowledge services must not depend on Qt;
- knowledge services must not depend on Ollama;
- tools must not depend on Qt;
- the agent controller must not depend on Qt;
- `OllamaClient` remains transport/provider infrastructure and must not contain
  agent policy;
- UI widgets must not become the authoritative knowledge API;
- the Assistant must not reach into another tab's private DB/widget state.

## Existing knowledge sources

The Assistant treats the application's existing knowledge systems as separate
but related sources.

### Specification catalogue

Backed by the existing Specifications database.

It answers questions such as:

- What specification is this?
- What is its title/type?
- What versions are known?
- What is the latest catalogued version?
- What metadata is available for the specification?

It does **not** provide clause text.

### Specification text index

Backed by the existing Specification Search database.

It answers questions such as:

- Which indexed clauses contain this phrase?
- Which indexed version contains a matching clause?
- What is the exact extracted text of the selected clause?

Search and detailed reading are intentionally separate operations.  Search
returns compact matches; detailed clause retrieval happens only when needed.

The indexed content is an extracted textual representation of a 3GPP source
document.  It should therefore be described as **indexed specification text**,
not as a byte-for-byte rendering of the original document.

### Structured protocol knowledge

Backed by the existing protocol database (`NASDatabase`, despite its historical
name).

The same structured store is populated by multiple protocol parsers, including
NAS, PFCP, GTP-U, RRC, RAN3 ASN.1 protocols, and PDU Session User Plane data.

Structured protocol knowledge is **derived evidence**.  Parsers may normalize,
expand, recursively unwrap, or synthesize structures.  It is extremely useful
for message/IE lookup but must not be confused with verbatim normative text.

When normative wording matters, specification-text evidence should be retrieved
as well.

## Version model

The application has three independent notions of "latest":

```text
Specification catalogue
    → latest version known by the catalogue

Specification Search
    → latest locally indexed text version

Protocol knowledge
    → latest locally imported/parsed structured version
```

These versions may differ.

For example:

```text
TS 29.244

Latest catalogued:  18.8.0
Latest indexed:     18.7.0
Latest structured:  18.6.0
```

The Assistant must never silently treat these as the same version.

### Default version policy

For a service-specific lookup with no explicit version:

- catalogue lookup uses catalogue knowledge;
- specification-text lookup uses the latest indexed version;
- protocol lookup uses the latest imported structured version.

When an exact version is requested, the service uses that version only if it is
available in that knowledge source.

When a Release is requested, the current implementation interprets the major
version component as the Release and selects the newest locally available
version for that Release.

## Protocol registry

`core.protocol_registry` describes **application capability**, not current data
coverage.

A protocol descriptor records:

- stable protocol ID;
- display name;
- protocol aliases;
- source specification(s);
- reference point/interface metadata where useful;
- whether deterministic parser support exists.

Example:

```text
PFCP
  source specification: TS 29.244
  parser supported: yes
  reference point: N4
```

This does **not** imply that TS 29.244 has already been imported into the
structured protocol database.

That distinction is fundamental to future knowledge-gap detection:

```text
parser exists + no imported data
    → data/coverage gap

no parser exists
    → capability gap
```

Reference points such as N4 are intentionally kept separate from protocol
aliases.  N4 may help describe PFCP applicability, but the registry should not
blindly treat every interface name as a protocol synonym.

## Stage 1 knowledge services

### `SpecificationKnowledgeService`

Location:

```text
core/knowledge/specifications.py
```

Responsibilities:

- search the specification catalogue;
- aggregate version/file rows into specification-level results;
- retrieve specification details;
- expose known and latest catalogued versions.

It does not search specification text.

### `SpecSearchKnowledgeService`

Location:

```text
core/knowledge/spec_search.py
```

Responsibilities:

- inspect indexed versions;
- resolve an indexed version/release scope;
- search indexed specification text;
- deterministically rank and bound matches;
- issue opaque search-result handles;
- retrieve the exact clause represented by a handle;
- explicitly mark oversized clause payloads as partial.

Opaque result handles deliberately hide SQLite primary keys from higher layers.

Example flow:

```text
search_text("PFCP Session Establishment", specification="29.244")
        │
        ▼
S-a1b2c3... → TS 29.244 vX, Clause Y, compact snippet
        │
        ▼
get_clause("S-a1b2c3...")
        │
        ▼
exact clause row referenced by the original search result
```

### `ProtocolKnowledgeService`

Location:

```text
core/knowledge/protocols.py
```

Responsibilities:

- report parser capability and imported structured coverage separately;
- resolve a protocol to its source specification(s);
- select the newest appropriate imported version;
- find structured protocol messages;
- retrieve message field/IE hierarchies;
- find IEs/fields and their message usage/definitions where available.

The service does not import protocol data.

## Search ranking

The existing Specification Search UI is an evolution-analysis interface, not an
LLM relevance engine.  The Assistant therefore performs a small deterministic
ranking step in the knowledge service without changing the existing database
search behavior.

Initial ranking favors roughly:

1. exact phrase/title matches;
2. phrase/all-term title matches;
3. exact phrase content/snippet matches;
4. occurrence count;
5. specification document order as a tie-breaker.

This is intentionally simple.

Phase 1 does **not** introduce:

- embeddings;
- a vector database;
- an LLM reranker.

Those should only be considered if evaluation demonstrates a real retrieval
problem.

## Result/status model

Knowledge operations return `ToolResult` objects with a machine-readable
`ToolStatus`.

Current statuses are:

| Status | Meaning |
| --- | --- |
| `FOUND` | Requested information was successfully retrieved. |
| `NOT_FOUND` | The source was successfully queried but no matching object was found. |
| `SOURCE_UNAVAILABLE` | The required knowledge source/version/coverage is unavailable. |
| `SOURCE_INCOMPLETE` | Useful information was returned, but the payload/coverage is incomplete. |
| `AMBIGUOUS` | Multiple plausible objects match the request. |
| `INVALID_REQUEST` | Arguments are invalid for the requested operation. |
| `ERROR` | The source could not be queried reliably because of an operational/internal error. |

An empty search result is therefore not automatically a knowledge gap.

## Evidence model

`core.agent_models` defines application-owned evidence objects.

Evidence includes:

- stable run-local evidence ID;
- source type;
- source name;
- specification;
- version;
- release date where available;
- clause/title where available;
- content;
- whether the content is complete;
- additional provenance metadata.

Future agent stages will give Qwen evidence IDs such as:

```text
E1
E2
E3
```

The application—not Qwen—will own the mapping from an evidence ID to its real
source.

This prevents the model from becoming the authority for citation identity.

## Research-run model

A future user question creates one `ResearchRun`.

A run owns:

- run ID;
- user question;
- bounded conversation context;
- cancellation token;
- evidence store;
- structured diagnostic trace.

A run ultimately ends as one of:

- `COMPLETED`;
- `CANCELLED`;
- `FAILED`.

`COMPLETED` does not necessarily mean every question was answerable.  A valid
completed result may explicitly report insufficient or conflicting evidence.

## Cancellation

`CancellationToken` implements cooperative cancellation.

Future controller/worker code should check cancellation:

- before a model call;
- after a model call;
- before a tool call;
- after a tool call;
- before final answer generation.

An active synchronous Ollama HTTP request may not be instantly interruptible.
The minimum required behavior is therefore:

```text
Stop requested
    ↓
mark run cancelled
    ↓
start no new model/tool operations
    ↓
discard an active operation's result when it returns
    ↓
finish as CANCELLED
```

The Assistant should not use `QThread.terminate()` for normal cancellation.

## Planned Phase 1 agent tools

The public model-facing tool set is intentionally small:

### Specification catalogue

```text
find_specifications
get_specification
```

### Specification text

```text
search_spec_text
get_spec_clause
```

### Structured protocol knowledge

```text
list_protocols
find_protocol_message
find_protocol_ie
```

These are semantic engineering tools rather than database primitives.

The model is never given:

- SQLite IDs;
- SQL;
- database paths;
- filesystem paths;
- generic HTTP access.

## Progressive retrieval

Specification retrieval follows:

```text
broad question
     ↓
search_spec_text
     ↓
compact ranked matches
     ↓
model selects relevant result
     ↓
get_spec_clause
     ↓
detailed evidence
```

The same principle applies to ambiguous protocol searches.

Large clauses are not silently truncated.  A bounded result is explicitly
marked `SOURCE_INCOMPLETE` with `content_complete=False` and metadata indicating
that more content exists.

## Security and authority

Phase 1 agent authority is conceptually:

```text
KNOWLEDGE_READ
```

only.

Model-generated tool arguments are untrusted input and will be validated by
Python before execution.

Retrieved specification/protocol content is **data/evidence**, never
instructions controlling the agent.

Future network/write capabilities must be separately authorized by the
application.  An LLM statement such as "the user probably wants this imported"
does not constitute permission.

## Database read-only boundary

Stage 1 exposes only read/query methods through the Assistant knowledge-service
API.

However, the existing application database manager classes perform some
initialization/schema-maintenance behavior in their constructors.  Stage 1
therefore does **not** claim that the underlying SQLite connection is physically
opened with `mode=ro`.

The important Phase 1 authority guarantee is:

> The Assistant exposes no mutation/import/delete operation to the agent.

Introducing strict SQLite read-only repository classes would be a broader
database refactor and is deferred unless there is a concrete need.

## Qt/threading model

The planned Qt execution model is:

```text
AssistantTab
     ↓
AgentResearchWorker(QThread)
     ↓
AgentController
```

The project already uses `QThread` workers extensively, so the Assistant will
follow that convention.

The worker should remain thin.  Agent orchestration belongs in the
Qt-independent controller.

One research run executes sequentially in one background worker.  Only one
research run is active in an Assistant conversation initially.

## QueueManager boundary

Interactive Assistant research should **not** run through `QueueManager`.

Research requires:

- incremental activity updates;
- conversational ownership;
- evidence state;
- cooperative Stop semantics.

Future user-approved knowledge acquisition is a better fit for the existing
maintenance/task queue.

The intended long-term distinction is:

```text
Assistant research
    → AgentResearchWorker

Approved persistent knowledge acquisition
    → deterministic application maintenance job
    → QueueManager
```

## Ollama integration

The existing `OllamaClient` remains provider/transport infrastructure.

A future `OllamaAgentAdapter` will:

- translate application tool definitions into Ollama native tool schemas;
- submit complete `/api/chat` requests;
- interpret tool calls;
- preserve Ollama response/usage metadata;
- return model-domain objects to `AgentController`.

`OllamaClient` should not contain research policy, tool-selection policy, or
evidence logic.

Phase 1 is local-only.  Failure to reach the configured Ollama service must not
silently send application data to a cloud provider.

## Diagnostics

Every future research run will maintain a structured in-memory trace containing
information such as:

- run/question metadata;
- model identity;
- model/tool sequence;
- normalized tool arguments;
- result status/metadata;
- evidence IDs;
- timings;
- warnings/errors;
- final outcome.

Normal application logs should remain compact and should not routinely duplicate
complete specification clauses or protocol payloads.

Internal model chain-of-thought is neither required nor stored.

## Configuration

Provider configuration and agent behavior are separate concerns.

Existing Ollama configuration remains responsible for:

- Ollama host;
- selected model/provider settings;
- proxy behavior;
- provider timeout/keep-alive behavior.

Agent-specific settings may later include:

- maximum research iterations;
- bounded conversation history;
- retrieval limits;
- evidence payload limits.

Security boundaries are not user preferences and must not be disabled through
configuration.

## Evaluation

The Assistant should be evaluated using scenario-based engineering questions,
not only by visually inspecting fluent answers.

Scenarios should cover:

- structured message lookup;
- IE lookup;
- specification discovery;
- procedural specification questions;
- cross-source questions;
- ambiguity;
- missing structured knowledge with text fallback;
- insufficient evidence;
- conflicting evidence;
- follow-up questions;
- malformed tool interaction;
- cancellation and infrastructure failures.

Hard failures include:

- fabricated evidence/citations;
- unauthorized actions;
- unsupported substantive claims presented as established facts;
- UI freezes/thread leaks;
- continuing research after cancellation.

Tool efficiency and latency are useful metrics but are secondary to correctness
and traceability.

## Stage 1 manual verification

Before Ollama integration, verify the knowledge layer directly.

Recommended checks:

1. `protocol_registry.resolve_protocol("PFCP")` resolves to TS 29.244.
2. `ProtocolKnowledgeService.list_protocols()` distinguishes parser capability
   from imported PFCP data.
3. `find_message("PFCP Session Establishment Request", protocol="PFCP")`
   returns normalized structured fields and provenance when PFCP data exists.
4. The same lookup returns `SOURCE_UNAVAILABLE`, not `NOT_FOUND`, when PFCP is
   supported but no relevant structured version is imported.
5. `SpecSearchKnowledgeService.get_indexed_versions("29.244")` reports text
   coverage independently from protocol coverage.
6. `search_text("PFCP Session Establishment", specification="29.244")` returns
   a bounded, ranked result set with opaque result IDs.
7. `get_clause(result_id)` retrieves the exact clause selected from the search.
8. `SpecificationKnowledgeService.get_specification("29.244")` reports catalogue
   coverage independently from text/protocol coverage.

The first milestone is successful when these three independent states can be
observed correctly:

```text
TS 29.244 catalogue coverage
TS 29.244 indexed-text coverage
TS 29.244 structured-protocol coverage
```

## Future phases

### Agent/controller integration

Later Phase 1 stages add:

- semantic tool registry;
- native Ollama tool calling;
- bounded controller loop;
- exact duplicate-call caching;
- evidence-ID citation validation;
- context budgeting;
- Qt worker;
- research-oriented Assistant UI.

### Knowledge-gap detection

A later phase can aggregate objective observations such as:

```text
PFCP parser supported
structured PFCP data unavailable
TS 29.244 indexed text available
text fallback succeeded
```

This is much stronger than allowing the LLM to declare "PFCP is missing."

### Proposal-driven knowledge acquisition

Future acquisition may:

1. identify a confirmed structured-knowledge gap;
2. identify the authoritative source specification;
3. determine whether a deterministic parser/importer already exists;
4. prepare an inspectable acquisition proposal;
5. require explicit user approval;
6. execute existing deterministic application maintenance code;
7. validate the resulting database;
8. retry the original structured lookup.

The LLM must not directly synthesize and insert arbitrary protocol database
records.

## Extending the Assistant safely

When adding a new knowledge capability:

1. prefer an existing application source/API;
2. add or extend a Qt/Ollama-independent knowledge service;
3. return normalized `ToolResult` statuses;
4. preserve provenance;
5. expose a semantic tool rather than a primitive;
6. give the tool the minimum required capability;
7. bound result size;
8. add benchmark/regression scenarios;
9. document the new capability here.

Do not bypass the knowledge-service boundary merely because a SQLite query is
convenient from a tool handler.


## Stage 2: semantic tool layer

Stage 2 adds `core/tool_registry.py` and `core/tools/`. The registry is the
application-owned enforcement point between model requests and knowledge
services. It recognizes only registered tools, enforces `KNOWLEDGE_READ`,
validates argument types/bounds, converts unexpected handler exceptions to
structured `ERROR` results, exports native Ollama-compatible function schemas,
and supplies a canonical exact-call key for later duplicate-call caching.

The public Phase 1 tool set is exactly:

```text
find_specifications
get_specification
search_spec_text
get_spec_clause
list_protocols
find_protocol_message
find_protocol_ie
```

Build it with:

```python
from modules.assistant.core.tools import build_read_only_tool_registry
registry = build_read_only_tool_registry(
    specs_db_path=...,
    spec_search_db_path=...,
    protocol_db_path=...,
)
```

Before Ollama integration, verify unknown tools and malformed/extra arguments
return `INVALID_REQUEST`; bounds are enforced; protocol data absence is
distinguished from a search miss; specification searches return opaque handles;
and `get_spec_clause` only accepts handles created by its own service instance.
`registry.ollama_tools()` should contain seven native function schemas.


## Stage 3: native Ollama tool calling

Stage 3 adds `core/ollama_agent_adapter.py` and a backward-compatible
non-streaming `OllamaClient.chat(...)` transport method.

### Transport boundary

`OllamaClient.chat(...)`:

- calls `/api/chat` with `stream: false`;
- preserves the complete response dictionary;
- accepts native Ollama tool schemas;
- applies the configured request timeout and `keep_alive`;
- contains no Assistant research policy.

The existing `stream_chat(...)` API remains available for existing callers. It
now also honors the configured timeout and `keep_alive`.

### Agent adapter boundary

`OllamaAgentAdapter`:

- uses the currently selected Ollama model unless explicitly overridden;
- serializes application-owned `ToolDefinition` objects using their native
  Ollama schemas;
- parses `message.tool_calls`;
- preserves timing/token metrics when Ollama returns them;
- converts responses to provider-independent `ModelResponse` and
  `ModelToolCall` objects;
- builds assistant/tool history messages for the future controller.

It does **not** execute tools or decide research strategy.

### First tool-calling smoke test

Before exposing all seven tools to Qwen, use only:

```text
list_protocols
find_protocol_message
```

with a question such as:

```text
What IEs are in the PFCP Session Establishment Request?
```

The expected sequence is:

```text
Qwen
  ↓
find_protocol_message(...) or list_protocols(...)
  ↓
ToolRegistry validates/executes
  ↓
tool result is returned to Qwen
  ↓
Qwen produces a grounded response
```

A `manual_stage3_smoke.py` helper is included at the bundle root. It is a manual
verification harness, not the introduction of an automated test framework.

### Stage 3 acceptance gate

Verify with the actual configured Qwen model:

1. Ollama returns a native `message.tool_calls` structure.
2. The adapter parses the tool name and arguments correctly.
3. `ToolRegistry` accepts valid arguments and rejects malformed ones.
4. A tool result can be appended to the conversation and Qwen continues.
5. The selected model does not invent unavailable tool names repeatedly.
6. `prompt_eval_count`, `eval_count`, and timing metrics are captured when
   provided by Ollama.
7. Existing `stream_chat()` callers continue to work.
8. Changing Timeout or Keep-Alive in the existing Ollama settings is honored by
   new client instances/reconfiguration.

The bounded multi-step research loop is intentionally deferred to Stage 4.


## Stage 4: bounded agent controller

Stage 4 adds `core/model_types.py` and `core/agent_controller.py`.

Provider-neutral `ModelResponse` / `ModelToolCall` types were deliberately moved
out of the Ollama adapter. This keeps the controller importable/testable without
Qt or Ollama and reinforces the dependency rule that agent policy is provider
independent.

`AgentController` now owns one sequential research run and implements bounded
conversation context, iteration limits, cooperative cancellation, validated
tool execution, exact duplicate-call caching, evidence creation, evidence-ID
citation validation, structured activity/trace events, and evidence-limited
finalization when the tool budget is exhausted.

Detailed tools create citable evidence; discovery tools do not:

```text
get_specification                 → catalogue evidence
get_spec_clause                   → indexed specification-text evidence
find_protocol_message / ..._ie    → structured protocol evidence

find_specifications / search_spec_text / list_protocols
                                   → discovery only
```

A deterministic `manual_stage4_controller_smoke.py` is included and requires no
Ollama or application database. It verifies model → tool → E1 → cited final
answer behavior.

Stage 4 acceptance includes duplicate-cache behavior, malformed tool recovery,
invented evidence-ID correction, cancellation, budget exhaustion, and then a
real-Qwen run using all seven tools. Qt integration remains the next stage.


## Stage 5: Qt research worker

Stage 5 adds:

```text
ui/__init__.py
ui/assistant_worker.py
```

and extends `AgentResult` so the presentation layer receives the application-owned
run ID, evidence objects, and diagnostic trace together with the final answer.

### Worker lifecycle

`AgentResearchWorker(QThread)` owns exactly one research run:

```text
start()
  ↓
research_started
  ↓
AgentController.run(...)
  ├─ activity_updated(ResearchActivity)
  ├─ evidence_added(Evidence)
  └─ ...
  ↓
one terminal signal:
  result_ready(AgentResult)
  research_cancelled(AgentResult)
  research_failed(message, AgentResult|None)
  ↓
QThread.finished
```

The worker is intentionally thin. It contains no research policy, SQL, prompt
strategy, or tool-selection logic.

### Cooperative Stop behavior

The UI calls:

```python
worker.request_cancel()
```

The worker forwards cancellation to the active controller. The controller checks
the run token between model/tool operations.

The worker does **not** call `QThread.terminate()`.

An already-blocking synchronous Ollama request may continue until it returns or
times out. After it returns, the controller observes cancellation and discards
the result rather than continuing research.

The UI should therefore distinguish:

```text
Stop pressed → Cancelling...
worker terminal signal → Stopped
```

### Evidence signals and final result

Evidence is emitted as soon as a detailed tool result creates an application
evidence object. The final `AgentResult` also carries the complete evidence list
and diagnostic trace.

This allows the future UI to render Sources from application data rather than
parsing citations back out of model prose.

### Exception boundary

`AgentResearchWorker.run()` is the top-level Qt exception boundary. Unexpected
Ollama/tool/controller exceptions are logged with a traceback and converted to
`research_failed`; they must not escape into the Qt event loop.

### Stage 5 smoke test

`manual_stage5_worker_smoke.py` uses `QCoreApplication`, a fake model adapter,
and a fake validated tool. It requires the application's PyQt5 environment but
does not require Ollama or any 3GPP database.

Verify:

1. the UI/event loop remains responsive;
2. activity signals arrive in order;
3. evidence `E1` is emitted;
4. the final result includes `E1`, run ID, and trace;
5. `finished` occurs and the worker is no longer running;
6. repeated worker runs do not accumulate threads;
7. calling `request_cancel()` produces `CANCELLED`;
8. cancelling during a real Ollama request shows `Cancelling...` until the
   request returns/timeout, then terminates without another tool/model call.

The actual Assistant conversation UI remains Stage 6.


## Stage 6: research-oriented Assistant UI

Stage 6 adds `ui/assistant_tab.py` and integrates the tab into `main_window.py`.

The initial UI deliberately focuses on the core research workflow:

```text
Conversation
Research activity
Sources
Question input + Send / Stop
```

### Conversation

Only user messages and final Assistant answers are retained as long-lived
conversation context. Detailed tool traffic remains scoped to the research run.

When a new question starts, the current conversation (excluding the just-added
question) is passed to `AgentResearchWorker`. The controller applies its own
recent-message bound before sending context to the model.

### Research activity

The activity pane displays objective application actions emitted by the
controller, not hidden model reasoning.

Examples include:

```text
Searching specification text...
Reading a selected specification clause...
Looking up a protocol message...
```

The pane can be collapsed after research completes.

### Sources

The Sources tree is populated from application-owned `Evidence` objects, not by
parsing model-generated citation text.

A source row displays:

- evidence ID;
- source/specification;
- version and clause where available.

Double-clicking a source currently opens a simple content preview. Navigation
into the existing Specification Search and Protocol tabs is intentionally
deferred until the core Assistant workflow has been validated.

### Send / Stop state

Only one research worker is active at a time.

While running:

```text
Send disabled
Question input disabled
Stop enabled
```

After Stop:

```text
status = Cancelling...
```

until the cooperative worker actually terminates. The UI does not claim that a
blocking Ollama HTTP request was instantly aborted.

### Shutdown

`main_window.closeEvent()` asks `AssistantTab` to cooperatively cancel an active
run and waits briefly. It does not force-terminate the Assistant thread.

### Main-window integration

The Assistant receives database paths rather than references to existing tabs:

```python
AssistantTab(
    specs_db_path=db_path,
    spec_search_db_path=spec_search_db_path,
    protocol_db_path=nas_db_path,
)
```

This preserves the boundary that the Assistant uses core knowledge services
rather than reaching into another UI widget's private state.

### Stage 6 manual acceptance gate

Run the application in its normal PyQt/Ollama environment and verify:

1. the new `🤖 Assistant` tab opens without affecting existing tabs;
2. asking a simple protocol question keeps the GUI responsive;
3. research activity appears while Qwen uses tools;
4. detailed tool results appear in Sources as application evidence;
5. final `[E#]` references correspond to visible Sources;
6. a follow-up question receives recent conversation context;
7. Send is disabled during an active run;
8. Stop changes to `Cancelling...` and no further research begins;
9. another question can be asked after completion/cancellation;
10. repeated runs do not leave running threads;
11. closing the application during research does not call
    `QThread.terminate()` on the Assistant worker;
12. Ollama unavailable/model misconfiguration produces a useful failure message;
13. existing Specifications, Spec Search, Protocols, Ollama settings, and other
    application functionality continue to behave as before.

After this gate, the next work should be real-model evaluation/tuning rather
than adding more UI features.


## Protocol discovery correction

Real-model testing exposed a missing semantic operation: a question such as
"What messages are defined for PFCP?" cannot validly call
`find_protocol_message`, because that tool requires a message name that the
model does not yet know.

The public tool set therefore now includes:

```text
list_protocol_messages(protocol, version=None, release=None, limit=100)
```

This is a structured discovery tool and creates structured protocol evidence.
`find_protocol_message` remains the detailed lookup for a known message.

Research activity now includes the concrete `ToolResult.message` for failures,
so an `INVALID_REQUEST` displays why it was rejected. Repeated invalid tool
calls also trigger an explicit schema-recovery instruction rather than allowing
the model to churn through the full research budget.
