# --- File: src/modules/nas/ui/nas_matrix_controller.py ---
import re
from typing import Any, List, Optional

from PyQt5.QtCore import QModelIndex, QPoint, Qt
from PyQt5.QtWidgets import QAction, QMenu

RE_SPEC_FROM_COL = re.compile(r"(?:24|25|29|36|37|38)\.[0-9]{3}")
RE_CLAUSE_FROM_REF = re.compile(r"((?:9|8|7|6|5|4|D\.6)(?:\.[0-9A-Za-z]+)+)")


def handle_table_cell_click(tab: Any, index: QModelIndex) -> None:
    """
    Handles cell selection in the protocol evolution matrix:
    - Extracts IE name, clean ASN.1 / NAS type ref, and clause.
    - Resolves definition records across selected protocol versions.
    - Updates NASInspectorWidget with syntax, definitions, and fields.
    - Triggers asynchronous reverse lookup for messages referencing this IE.
    """
    model = tab.matrix_table.model()
    if not model or not index.isValid():
        return

    row, col = index.row(), index.column()
    ie_name = str(model.data(model.index(row, 1), Qt.DisplayRole) or "").strip().lstrip("└─ ")
    type_ref = str(model.data(model.index(row, 2), Qt.DisplayRole) or "").strip()

    field_desc = ""
    if hasattr(model, "_pivot_df") and not model._pivot_df.empty and row < len(model._pivot_df):
        row_data = model._pivot_df.iloc[row]
        field_desc = str(row_data.get("field_description", "") or "")

    clean_type = re.sub(r"^SetupRelease\s*\{\s*([A-Za-z0-9\-]+)\s*\}", r"\1", type_ref)
    clean_type = re.sub(
        r"^SEQUENCE\s*(?:\(SIZE\s*\([^)]*\)\)\s*)?OF\s+([A-Za-z0-9\-]+)",
        r"\1",
        clean_type,
    )
    clean_type = re.sub(
        r"^OCTET STRING\s*\(\s*CONTAINING\s+([A-Za-z0-9\-]+)\s*\)",
        r"\1",
        clean_type,
    )
    clean_type = re.sub(r"[\(\{\[].*$", "", clean_type).strip().rstrip(",").strip()

    if clean_type.upper() in (
        "ENUMERATED",
        "INTEGER",
        "BOOLEAN",
        "BIT STRING",
        "OCTET STRING",
        "NULL",
    ):
        clean_type = ""

    spec_num = getattr(tab, "_current_message_spec", None)
    target_version_hint = None

    if col >= 3:
        col_name = str(model._get_visible_column_name(col))
        match_spec = RE_SPEC_FROM_COL.search(col_name)
        if match_spec:
            spec_num = match_spec.group(0)
        target_version_hint = col_name

    match_clause = RE_CLAUSE_FROM_REF.search(type_ref)
    clause = match_clause.group(1).strip() if match_clause else ie_name

    defs = []
    if clean_type:
        defs = tab.db.get_ie_definitions_by_clause(
            clause=clean_type,
            alt_name=ie_name,
            spec_number=spec_num,
            version_ids=tab.selected_version_ids,
        )

    if not defs and clause and not any(
        clause.upper().startswith(k)
        for k in ("ENUMERATED", "INTEGER", "BOOLEAN", "BIT STRING")
    ):
        defs = tab.db.get_ie_definitions_by_clause(
            clause=clause,
            alt_name=ie_name,
            spec_number=spec_num,
            version_ids=tab.selected_version_ids,
        )

    tab.inspector.display_definitions(
        clause=clause,
        ie_name=clean_type or ie_name,
        spec_number=spec_num,
        defs=defs,
        containing_msgs=[],
        target_version_hint=target_version_hint,
        fallback_type_ref=type_ref,
        field_description=field_desc,
    )

    if defs:
        resolved_name = defs[0]["ie_name"]
        search_target_name = clean_type or resolved_name
        tab._start_reverse_lookup(
            clause=clause,
            ie_name=search_target_name,
            spec_num=spec_num,
        )


def show_matrix_context_menu(tab: Any, pos: QPoint) -> None:
    """
    Renders context menu for rows in the evolution matrix:
    - Filter left list for IE name
    - Inspect Definition
    - Open containing messages using this IE
    """
    model = tab.matrix_table.model()
    if not model:
        return

    index = tab.matrix_table.indexAt(pos)
    if not index.isValid():
        return

    row = index.row()
    ie_name = str(model.data(model.index(row, 1), Qt.DisplayRole) or "").strip().lstrip("└─ ")
    type_ref = str(model.data(model.index(row, 2), Qt.DisplayRole) or "").strip()

    match = RE_CLAUSE_FROM_REF.search(type_ref)
    clause = match.group(1).strip() if match else ie_name
    spec_num = getattr(tab, "_current_message_spec", None)

    col = index.column()
    target_version = None
    if col >= 3 and hasattr(model, "_get_visible_column_name"):
        col_name = str(model._get_visible_column_name(col))
        match_spec = RE_SPEC_FROM_COL.search(col_name)
        if match_spec:
            spec_num = match_spec.group(0)
        m_ver = re.search(r"v?([0-9]+\.[0-9]+\.[0-9]+)", col_name)
        if m_ver:
            target_version = m_ver.group(1)

    menu = QMenu(tab)
    menu.setStyleSheet("""
        QMenu {
            font-size: 11px;
            background-color: #FFFFFF;
            border: 1px solid #CBD5E1;
            padding: 4px;
        }
        QMenu::item {
            padding: 5px 18px 5px 10px;
            border-radius: 3px;
        }
        QMenu::item:selected {
            background-color: #EBF3FC;
            color: #1E5C99;
        }
    """)

    filter_term = ie_name if ie_name else clause
    act_filter = QAction(f"🔍 Filter message list for '{filter_term}'", tab)
    act_filter.triggered.connect(lambda: tab._on_filter_by_ie_requested(filter_term))
    menu.addAction(act_filter)

    act_inspect = QAction(f"📖 Inspect Definition ({filter_term})", tab)
    act_inspect.triggered.connect(lambda: tab._on_table_cell_clicked(index))
    menu.addAction(act_inspect)

    if clause or ie_name:
        containing_msgs = tab.db.get_messages_using_ie(
            clause=clause,
            ie_name=ie_name,
            spec_number=spec_num,
            version_ids=tab.selected_version_ids,
        )
        if containing_msgs:
            sub_menu = menu.addMenu(f"📂 Open message using this IE ({len(containing_msgs)} found)...")
            for m in containing_msgs:
                mn = m["message_name"]
                m_spec = m.get("spec_number", "")
                spec_tag = f" [TS {m_spec}]" if m_spec and not spec_num else ""
                sub_act = QAction(f"{mn} ({m['clause']}){spec_tag}", tab)
                sub_act.triggered.connect(lambda checked, target=mn: tab._jump_to_message(target))
                sub_menu.addAction(sub_act)

    if spec_num and hasattr(tab, "open_specification"):
        menu.addSeparator()
        ver_tag = f" v{target_version}" if target_version else ""
        act_open_word = QAction(f"📝 Open TS {spec_num}{ver_tag} in Word (Clause {clause})", tab)
        act_open_word.triggered.connect(
            lambda: tab.open_specification(
                spec_number=spec_num,
                version=target_version,
                clause_number=clause,
                doc_format="word",
            )
        )
        menu.addAction(act_open_word)

        act_open_pdf = QAction(f"📕 Open TS {spec_num}{ver_tag} in PDF", tab)
        act_open_pdf.triggered.connect(
            lambda: tab.open_specification(
                spec_number=spec_num,
                version=target_version,
                clause_number=clause,
                doc_format="pdf",
            )
        )
        menu.addAction(act_open_pdf)

        act_open_folder = QAction("📂 Open Containing Folder in Explorer", tab)
        act_open_folder.triggered.connect(
            lambda: tab.open_specification(
                spec_number=spec_num,
                version=target_version,
                doc_format="folder",
            )
        )
        menu.addAction(act_open_folder)

        if hasattr(tab, "navigate_to_specs_tab_requested"):
            menu.addSeparator()
            act_specs = QAction(f"📚 Show TS {spec_num}{ver_tag} in Specifications Tab", tab)
            act_specs.triggered.connect(
                lambda: tab.navigate_to_specs_tab_requested.emit(spec_num, target_version or "")
            )
            menu.addAction(act_specs)

            act_search = QAction(f"🔎 Search TS {spec_num} in Spec Search Tab", tab)
            act_search.triggered.connect(
                lambda: tab.search_in_spec_search_requested.emit(spec_num, target_version or "")
            )
            menu.addAction(act_search)

    menu.exec_(tab.matrix_table.viewport().mapToGlobal(pos))
