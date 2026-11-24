# `modules.word_tools.core.sensitivity`

Acyclic package backing the legacy `modules.word_tools.core.sensitivity_label` shim.

## Architecture

- `policy_discovery.py` — Parses Windows Office CLP/MIP `policy*.gz` SyncFile XML caches (`discover_windows_sensitivity_policy`, `_parse_clp_syncfile`).
- `inspector.py` — Headless `docProps/custom.xml` inspection and Word COM sensitivity-label readers (`get_document_sensitivity_label`, `MsoAssignmentMethod`).
- `com_labeler.py` — Applies configured MIP labels onto open Word COM documents (`set_sensitivity_label`).
- `batch_labeler.py` — Batch sensitivity labeling across directories or file lists (`batch_set_sensitivity_label`).
- `__init__.py` — Public package exports.

## Dependency Flow

`policy_discovery` (leaf) and `inspector` (leaf) -> `com_labeler` -> `batch_labeler` -> `__init__`.
