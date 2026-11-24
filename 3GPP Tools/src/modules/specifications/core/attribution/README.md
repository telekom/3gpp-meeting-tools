# Clause-to-TDoc Attribution Package (`specifications/core/attribution`)

Multi-stage attribution engine linking 3GPP specification clause modifications and search queries to originating CRs and WG p-CRs.

## Module Inventory

| File | Responsibility |
|---|---|
| `text_helpers.py` | `BOILERPLATE_STOPWORDS` + `AttributionTextMixin` (`clean_tokens`, `extract_key_issue_references`, `extract_salient_diff_keywords`, `resolve_ancestor_context`) |
| `clause_matcher.py` | `ClauseMatcherMixin` — Stage 1 SQLite metadata/semantic ranking (Mode A CRs & Mode B omnibus p-CRs) |
| `docx_verifier.py` | `DocxVerifierMixin` — Stage 2 OpenXML Track-Changes-aware sentence & 6-shingle fingerprinting |
| `origin_ranker.py` | `OriginRankerMixin` — Cross-release Origin CR scoring (`rank_origin_crs_for_query`) |
| `engine.py` | Composed `SpecClauseAttributionEngine` class |
| `workers.py` | Background `QThread` workers (`SpecClauseAttributionThread`, `SpecClauseDeepVerifyThread`, `SpecFeatureOriginRankingThread`) |

`spec_clause_attribution.py` remains as a thin backward-compatible re-export shim.
