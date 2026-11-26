# NAS / Protocol Inspector UI Components Package (`components/`)

Modular decomposition of the protocol UI delegates, version selectors, cards, and inspector previously housed in `nas/ui/nas_components.py`.

## Submodule Navigation Map

| Module | Class / Responsibility | Key Symbols |
|---|---|---|
| [`delegates.py`](delegates.py) | 2-line message/IE list delegate and Evolution Matrix tree delegate (pixel-scaled) | `NASMessageItemDelegate`, `NASMatrixTreeDelegate` |
| [`version_tree.py`](version_tree.py) | Legacy collapsible tree selector for imported specs and versions | `NASVersionTreeWidget` |
| [`spec_card.py`](spec_card.py) | Compact release pills and collapsible specification card widgets | `_extract_major_release`, `NASReleasePillWidget`, `NASSpecCardWidget` |
| [`version_cards.py`](version_cards.py) | Scrollable cards container with `NASVersionTreeWidget` compatibility shims | `_NASCardCompatChild`, `_NASCardCompatSpec`, `NASVersionCardsWidget` |
| [`inspector_render.py`](inspector_render.py) | Clause 9 / ASN.1 / IE definition HTML rendering, cross-ref regexes, usage menu mixin | `RE_NAS_CROSS_REF`, `RE_EXPLICIT_SPEC_REF`, `NASInspectorRenderMixin` |
| [`inspector.py`](inspector.py) | Top-level `QFrame` Clause 9 / IE definition & Full Packet Stack inspector card | `NASInspectorWidget` |
| [`packet_stack_widget.py`](packet_stack_widget.py) | Interactive 5G UP Full Packet Stack viewer (`TS 29.281` + `TS 38.415`), flag toggles & `packetdiag` export | `PacketStackWidget` |
| [`__init__.py`](__init__.py) | Package exports | All public and compatibility symbols above |

## Backward Compatibility

- [`../nas_components.py`](../nas_components.py) remains as a thin re-export shim exposing all 12 classes, helpers, and constants for [`ui/__init__.py`](../__init__.py), [`nas_tab/`](../nas_tab), and the test suite.
