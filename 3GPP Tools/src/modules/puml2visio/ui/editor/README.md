# PlantUML Editor Package (`modules.puml2visio.ui.editor`)

Acyclic modular package extracted from `ui/ui_components.py` (thin backward-compatible re-export shim):

- `editor_mixins.py`: `CodeEditorEditingMixin` (snippet insertion, comment toggling, soft-tab/auto-indent key handling, and preset context menu).
- `code_editor.py`: `LineNumberArea`, extension constants, and `CodeDropTextEdit` (gutter rendering, current/error line highlighting, file loader, and drag-and-drop).
- `find_replace_bar.py`: `FindReplaceBar` inline search/replace toolbar.
