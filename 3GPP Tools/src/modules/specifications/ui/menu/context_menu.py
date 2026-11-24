"""Table context menu builder for SpecificationsTab rows."""

from __future__ import annotations

from pathlib import Path
import re

from PyQt5.QtWidgets import QComboBox, QLabel, QMenu

from core.ui.theme import C
from modules.specifications.ui.menu.copy_menus import (
    populate_multi_spec_copy_menu,
    populate_spec_copy_menu,
)


def show_spec_context_menu(tab, position):
    """
    Constructs and executes the context menu for selected specifications rows in the table.

    :param tab: The SpecificationsTab parent widget instance.
    :param position: QPoint position of the context menu request from table viewport.
    """
    selected_rows = tab.table.selectionModel().selectedRows()
    if not selected_rows:
        return

    target_items = []
    for index in selected_rows:
        row = index.row()
        widget = tab.table.cellWidget(row, 0)
        if widget:
            num_label = widget.findChild(QLabel, "specNumberLabel")
            if num_label:
                canonical = num_label.property("specNumber")
                spec_num = ""
                if canonical:
                    spec_num = str(canonical).strip()
                else:
                    display_text = num_label.text().strip()
                    match = re.match(r"^(\d{2}\.\d{2,3}(?:-[A-Za-z0-9]+)?)", display_text)
                    if match:
                        spec_num = match.group(1)

                if spec_num:
                    # Resolve currently selected version from Column 2 combo
                    cell_widget = tab.table.cellWidget(row, 2)
                    v_info = None
                    if cell_widget:
                        combo = cell_widget.findChild(QComboBox)
                        if combo:
                            v_info = combo.currentData()
                    target_items.append((spec_num, v_info))

    if not target_items:
        return

    menu = QMenu(tab)
    menu.setStyleSheet(f"""
        QMenu {{ background-color: {C.BG_ALT_LIGHT}; border: 1px solid {C.BORDER_STRONG}; }}
        QMenu::item {{ padding: 5px 20px 5px 15px; color: {C.TEXT_BODY}; }}
        QMenu::item:selected {{ background-color: {C.ACCENT_SOFT}; color: {C.ACCENT_PRIMARY}; }}
        QMenu::item:disabled {{ color: {C.TEXT_FAINT}; }}
    """)

    if len(target_items) == 1:
        spec_num, v_info = target_items[0]
        is_fav = spec_num in tab.favorites
        fav_label = "⭐ Remove from Favorites" if is_fav else "⭐ Add to Favorites"
        menu.addAction(fav_label).triggered.connect(lambda _, s=spec_num: tab._toggle_favorite(s))

        edit_ann_action = menu.addAction("🏷️ Edit Tags && Notes...")
        edit_ann_action.triggered.connect(lambda _, s=spec_num: tab._open_edit_annotations(s))

        menu.addSeparator()

        # Quick Copy Submenu
        ver_badge = f" (v{v_info['version']})" if v_info and v_info.get("version") else ""
        copy_sub = menu.addMenu(f"📋 Quick Copy{ver_badge}")
        copy_sub.setToolTipsVisible(True)
        copy_sub.setToolTip("Copy specification metadata in plain text and rich HTML form.")
        populate_spec_copy_menu(copy_sub, tab, spec_num, version_info=v_info, db=tab.db)

        menu.addSeparator()

        info_action = menu.addAction("ℹ️ View Details")
        info_action.triggered.connect(lambda _, s=spec_num: tab._show_spec_info(s))

        ver_str = v_info.get("version", "") if v_info else ""
        clean_ver = ver_str.lstrip("vV").strip() if ver_str else ""

        hist_action = menu.addAction(f"📜 View Change History && TDocs{ver_badge}")
        hist_action.setToolTip("Browse all formal CRs, p-CRs, and integrated TDocs.")
        hist_action.triggered.connect(lambda _, s=spec_num, v=clean_ver: tab._open_version_history_dialog(s, v))

        search_action = menu.addAction("🔎 Track Evolution in Spec Search")
        search_action.triggered.connect(
            lambda _, s=spec_num, v=clean_ver: tab.search_in_spec_search_requested.emit(s, v)
        )

        is_indexed = (
            tab.spec_search_db.is_version_indexed(spec_num, clean_ver)
            if hasattr(tab, "spec_search_db") and tab.spec_search_db and clean_ver
            else False
        )
        if is_indexed:
            index_action = menu.addAction(f"🔄 Re-index v{clean_ver} in Spec Search DB")
            index_action.setToolTip(f"TS {spec_num} v{clean_ver} is already indexed. Re-extract and refresh search index.")
        elif clean_ver:
            index_action = menu.addAction(f"📥 Index v{clean_ver} into Spec Search DB")
            index_action.setToolTip(f"Ingest and index TS {spec_num} v{clean_ver} into the local trigram Spec Search database.")
        else:
            index_action = menu.addAction("📥 Index into Spec Search DB")
            index_action.setToolTip("Ingest this specification into the local trigram Spec Search database.")

        index_action.triggered.connect(
            lambda _, s=spec_num, v=clean_ver: tab.index_in_spec_search_requested.emit(s, v)
        )

        smart_index_action = menu.addAction(f"⭐ Index Latest Releases of TS {spec_num}...")
        smart_index_action.setToolTip(f"Check and ingest latest versions of all releases for TS {spec_num} into Spec Search DB.")
        smart_index_action.triggered.connect(
            lambda _, s=spec_num: tab.smart_index_in_spec_search_requested.emit(s)
        )

        menu.addSeparator()

        spec_target_dir = Path(tab.download_dir) / spec_num
        folder_action = menu.addAction("📂 Open Local Folder")
        folder_action.triggered.connect(lambda _, s=spec_num: tab._open_spec_folder(s))
        if not (spec_target_dir.exists() and any(spec_target_dir.iterdir())):
            folder_action.setText("📁 Folder Not Created")
            folder_action.setEnabled(False)

        menu.addSeparator()

        update_action = menu.addAction(f"🔄 Sync metadata for TS {spec_num}")
        update_action.triggered.connect(
            lambda: tab.update_specific_requested.emit([spec_num], tab.force_meta_action.isChecked())
        )

        menu.addSeparator()

        del_action = menu.addAction(f"🗑️ Delete TS {spec_num}")
        del_action.triggered.connect(lambda _, s=spec_num: tab._delete_single_spec(s))
    else:
        target_specs = [item[0] for item in target_items]
        menu.addAction(f"⭐ Add selected ({len(target_specs)}) to Favorites").triggered.connect(
            lambda: tab._batch_set_favorites(target_specs, True)
        )
        menu.addAction(f"🗑️ Remove selected ({len(target_specs)}) from Favorites").triggered.connect(
            lambda: tab._batch_set_favorites(target_specs, False)
        )
        menu.addSeparator()

        batch_tag_action = menu.addAction(f"🏷️ Add Tag to selected ({len(target_specs)})...")
        batch_tag_action.triggered.connect(lambda: tab._batch_add_tag(target_specs))

        menu.addSeparator()

        copy_sub = menu.addMenu(f"📋 Quick Copy ({len(target_items)} selected)")
        copy_sub.setToolTipsVisible(True)
        copy_sub.setToolTip("Copy metadata and download links for all selected specifications.")
        populate_multi_spec_copy_menu(copy_sub, tab, target_items, db=tab.db)

        menu.addSeparator()

        batch_index_action = menu.addAction(f"📥 Index selected ({len(target_items)}) into Spec Search DB")
        batch_index_action.setToolTip("Ingest and index active versions for all selected specifications into Spec Search DB.")
        batch_index_action.triggered.connect(lambda: tab._batch_index_in_spec_search(target_items))

        menu.addSeparator()

        update_action = menu.addAction(f"🔄 Sync metadata for selected ({len(target_specs)})")
        update_action.triggered.connect(
            lambda: tab.update_specific_requested.emit(target_specs, tab.force_meta_action.isChecked())
        )

        menu.addSeparator()

        del_action = menu.addAction(
            f"🗑️ Delete selected ({len(target_specs)}) specifications"
        )
        del_action.triggered.connect(lambda: tab._delete_batch_specs(target_specs))

    menu.exec_(tab.table.viewport().mapToGlobal(position))
