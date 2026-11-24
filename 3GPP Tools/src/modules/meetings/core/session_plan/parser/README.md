# SA2 Session Plan Parser (`session_plan/parser`)

Fault-tolerant OpenXML (`.docx`) and Markdown/Text parser producing `SessionPlanSnapshot`, `SessionSlot`, and `AgendaSlotItem` domain models.

## Module Inventory

| File | Responsibility |
|---|---|
| `constants.py` | Period timings (`PERIOD_TIMES` / `QUARTER_TIMES`), `DAY_NAMES`, `ROOM_OVERRIDE_RE`, `ADMIN_NOTE_RE`, XML namespace/ignore sets |
| `agenda_items.py` | `AgendaItemsParserMixin` — room overrides, parenthesis balancing, KI# mapping, numeric range expansion |
| `docx_table.py` | `DocxTableParserMixin` — Track Changes XML extraction (`<w:ins>`/`<w:moveTo>`), `_vMerge` dedup, colspan expansion, break/session slot creation |
| `markdown_parser.py` | `MarkdownParserMixin` — Markdown/text table parsing, SHA-256 schedule hashing, idempotent `.docx`/plan staging |
| `parser.py` | Composed `SessionPlanParser` class |

`session_plan_parser.py` re-exports all public symbols and constants for full backward compatibility.
