# TDocs Reusable UI Components (`meetings/ui/tdocs/components`)

Theme-aware multi-selection filter combobox, Word document comparison cart, and rich tooltip helpers.

## Module Inventory

| File | Responsibility |
|---|---|
| `combo_delegate.py` | `CheckableComboItemDelegate` — vector checkmark & high-DPI popup row renderer |
| `checkable_combo.py` | `CheckableComboBox` — summary-sized multi-select dropdown with `(Select All)`, availability dimming & `activeFilter` QSS property |
| `comparison_cart.py` | `ComparisonCartWidget`, `populate_slot_word_menu()`, and `make_rich_tooltip()` |

`meetings/ui/tdocs_components.py` remains as a thin backward-compatible re-export shim.
