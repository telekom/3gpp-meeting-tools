# Protocol SQLite Database Package (`nas/core/db`)

Modular SQLite manager for `3gpp_protocol_data.db` across NAS, 5GS UE Policies, RRC, NGAP/XnAP/F1AP/S1AP, GTPv1-U, PDU Session UP, and PFCP protocols.

## Module Inventory

| File | Responsibility |
|---|---|
| `schema.py` | Regex constants, `parse_version_tuple()`, and `NasDbSchemaMixin` (connections, WAL setup, migrations, clear/wipe) |
| `ingestion.py` | `NasDbIngestionMixin` — atomic `insert_parsed_spec()` upserting versions, messages, IEs, and definitions |
| `queries_messages.py` | `NasDbMessageQueriesMixin` — tokenized CTE search, reverse IE lookup, evolution DataFrame, message context info |
| `queries_ies.py` | `NasDbIeQueriesMixin` — IE definition listing, evolution DataFrame, umbrella clause disambiguation, cross-spec lookups |
| `database.py` | Composed `NASDatabase` class |

`nas/core/nas_db.py` remains as a thin backward-compatible re-export shim.
