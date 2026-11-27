# Specification Import Studio (`spec_import/`)

Two-pane faceted specification browser and multi-spec ingestion queue for `NASTab` (`NASVersionSelectDialog`).

| File | Responsibility |
|---|---|
| [`dialog.py`](dialog.py) | `NASVersionSelectDialog` two-pane studio, quick-bundle actions, multi-spec queue, and `preselect_spec()` deep-link entry point |
| [`spec_list.py`](spec_list.py) | `SpecFacetBar` (search + Tech/Domain/Plane/Layer chips), `SpecRowWidget`, and `SpecListPanel` |
| [`release_panel.py`](release_panel.py) | `ReleasePanel` header card, release filter chips (`All` / `Latest / Rel` / `Cached` / `Rel-XX`), and checkable release table |
| [`chips.py`](chips.py) | Theme-tokenized facet badges (`make_badge`), filter chips (`make_chip`), and primary button QSS |
| [`__init__.py`](__init__.py) | Package export for `NASVersionSelectDialog` |
