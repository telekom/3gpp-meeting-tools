# Company Sanitization Package (`core/utils/company_sanitization`)

Modular 3GPP contributor normalization and confirmed vs. tentative (`?`, `(?)`, `[...]`) source segmentation engine.

## Module Inventory

| File | Responsibility |
|---|---|
| `defaults.py` | Built-in `BUILTIN_SYNONYMS_PATTERNS` baseline catalog and precompiled helper regexes |
| `matcher.py` | `CompanyMatcherMixin` — `company_synonyms*.json` loading, C-accelerated matcher compilation, LRU cache clearing |
| `tokenizer.py` | `CompanyTokenizerMixin` — source segmentation (`tokenize_source_segments`), confirmed vs. tentative parsing (`parse_contributors_detailed`), primary author resolution |
| `formatters.py` | `CompanyFormatterMixin` — greyed/italic HTML, tooltip breakdowns, Two-Line Markdown headers, Markdown table cells, LLM triage context |
| `sanitizer.py` | Composed `CompanySanitizer` holding shared class state and triggering initial `reload_rules()` |

`core/utils/company_sanitizer.py` remains as a thin backward-compatible re-export shim.
