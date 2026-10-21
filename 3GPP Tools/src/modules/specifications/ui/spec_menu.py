# --- File: src/modules/specifications/ui/spec_menu.py ---
"""
Context menu builder for the Specifications table.
Provides multi-select actions for toggling favorites, synchronizing metadata,
and deleting specifications.
"""
import re
from PyQt5.QtWidgets import QLabel, QMenu


def show_spec_context_menu(tab, position):
    """
    Constructs and executes the context menu for selected specifications rows in the table.

    :param tab: The SpecificationsTab parent widget instance.
    :param position: QPoint position of the context menu request from table viewport.
    """
    selected_rows = tab.table.selectionModel().selectedRows()
    if not selected_rows:
        return

    target_specs = []
    for index in selected_rows:
        widget = tab.table.cellWidget(index.row(), 0)
        if widget:
            num_label = widget.findChild(QLabel, "specNumberLabel")
            if num_label:
                # Never use decorated UI text (e.g. "33.801-01 ⭐") as an identifier.
                canonical = num_label.property("specNumber")
                if canonical:
                    target_specs.append(str(canonical).strip())
                else:
                    # Compatibility fallback for widgets created by older code.
                    display_text = num_label.text().strip()
                    match = re.match(r"^(\d{2}\.\d{2,3}(?:-[A-Za-z0-9]+)?)", display_text)
                    if match:
                        target_specs.append(match.group(1))

    if not target_specs:
        return

    menu = QMenu(tab)
    menu.setStyleSheet("""
        QMenu { background-color: #FAFAFA; border: 1px solid #CCC; }
        QMenu::item { padding: 5px 20px 5px 15px; color: #333333; }
        QMenu::item:selected { background-color: #E1F0FF; color: #0078D7; }
    """)

    if len(target_specs) == 1:
        spec_num = target_specs[0]
        is_fav = spec_num in tab.favorites
        fav_label = "⭐ Remove from Favorites" if is_fav else "⭐ Add to Favorites"
        menu.addAction(fav_label).triggered.connect(lambda _, s=spec_num: tab._toggle_favorite(s))
    else:
        menu.addAction(f"⭐ Add selected ({len(target_specs)}) to Favorites").triggered.connect(
            lambda: tab._batch_set_favorites(target_specs, True)
        )
        menu.addAction(f"🗑️ Remove selected ({len(target_specs)}) from Favorites").triggered.connect(
            lambda: tab._batch_set_favorites(target_specs, False)
        )

    menu.addSeparator()
    update_action = menu.addAction(f"🔄 Sync metadata for selected ({len(target_specs)})")
    update_action.triggered.connect(
        lambda: tab.update_specific_requested.emit(target_specs, tab.force_meta_action.isChecked())
    )

    menu.addSeparator()
    if len(target_specs) == 1:
        menu.addAction(f"🗑️ Delete TS {target_specs[0]}").triggered.connect(
            lambda _, s=target_specs[0]: tab._delete_single_spec(s)
        )
    else:
        menu.addAction(f"🗑️ Delete selected ({len(target_specs)}) specifications").triggered.connect(
            lambda: tab._delete_batch_specs(target_specs)
        )

    menu.exec_(tab.table.viewport().mapToGlobal(position))
