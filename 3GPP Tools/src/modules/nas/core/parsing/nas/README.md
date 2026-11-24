# NAS & UE Policy Parser Package (`modules.nas.core.parsing.nas`)

Acyclic modular package extracted from `parsing/nas_parser.py` (thin backward-compatible re-export shim):

- `catalogs.py`: WordprocessingML tags, clause/caption regexes, encapsulated container sub-IE tables, and synthetic UE Policy catalogs (`TS24526_POLICY_MESSAGES`, `TS24588_POLICY_MESSAGES`, `TS24555_POLICY_MESSAGES`, `TS24514_POLICY_MESSAGES`, `TS24578_POLICY_MESSAGES`).
- `container_unroller.py`: `finalize_ie_definition()` and `unroll_container_sub_ies()`.
- `parser.py`: `NASDocxParser` for Clause 8 / Annex D.5 message tables and Clause 9 / Annex D.6 / UE Policy IE coding sections.
