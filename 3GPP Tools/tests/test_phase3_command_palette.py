import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from PyQt5.QtCore import QEvent, Qt
from PyQt5.QtGui import QKeyEvent
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget, QWidget

from core.ui.command_palette import (
    _KIND_ITEM,
    _ROLE_KIND,
    CommandItem,
    CommandPaletteDialog,
    DynamicPaletteSearcher,
    build_static_command_items,
)
from core.ui.nav_rail import NavRail
from core.ui.right_side_panel import MODE_ASSISTANT, RightSidePanel
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.ui.ui_tabs import MeetingsTab
from modules.specifications.core.database import SpecsDatabase
from modules.specifications.ui.ui_tabs import SpecificationsTab
from modules.word_tools.ui.word_tabs import WordExtractorTab
from modules.work_items.core.wi_database import WorkItemsDatabase
from modules.work_items.ui.ui_tabs import WorkItemsTab


def _seed_temp_dbs(tmp_dir: Path):
    data_db = tmp_dir / "3gpp_data.db"
    proto_db = tmp_dir / "3gpp_protocol_data.db"

    # 1. Specifications
    sdb = SpecsDatabase(data_db)
    with sdb._get_connection() as conn:
        conn.execute("INSERT OR IGNORE INTO series (id, name, url) VALUES (1, '23', '')")
        conn.execute("INSERT OR IGNORE INTO series (id, name, url) VALUES (2, '24', '')")
        conn.execute("INSERT OR IGNORE INTO working_groups (id, name) VALUES (1, 'SA2')")
        conn.execute("INSERT OR IGNORE INTO working_groups (id, name) VALUES (2, 'CT1')")
        conn.execute(
            """
            INSERT OR REPLACE INTO specifications
            (id, series_id, number, title, type, radio_technology, primary_group_id)
            VALUES (1, 1, '23.501', 'System architecture for the 5G System (5GS)', 'TS', '5G', 1)
            """
        )
        conn.execute(
            """
            INSERT OR REPLACE INTO specifications
            (id, series_id, number, title, type, radio_technology, primary_group_id)
            VALUES (2, 2, '24.501', 'Non-Access-Stratum (NAS) protocol for 5G System (5GS)', 'TS', '5G', 2)
            """
        )
        conn.execute(
            """
            INSERT OR REPLACE INTO files (spec_id, filename, version, url, upload_date)
            VALUES (2, '24501-i50.zip', '18.5.0', '', '2024-12-10')
            """
        )
        conn.commit()

    # 2. Meetings
    mdb = MeetingsDatabase(data_db)
    with mdb._get_connection() as conn:
        conn.execute("INSERT OR IGNORE INTO working_groups (id, name) VALUES (1, 'SA2')")
        conn.execute(
            """
            INSERT INTO meetings
            (wg_id, meeting_number, location, start_date, end_date,
             first_tdoc_prefix, first_tdoc_num, last_tdoc_prefix, last_tdoc_num)
            VALUES (1, 'SA2#168', 'Budapest, Hungary', '2025-04-07', '2025-04-11',
                    'S2-25', 1000, 'S2-25', 2000)
            """
        )
        conn.commit()

    # 3. Work Items
    wdb = WorkItemsDatabase(data_db)
    with wdb._get_connection() as conn:
        conn.execute(
            """
            INSERT INTO work_items
            (code, acronym, name, latest_wid, release, start_date, end_date)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            ("980042", "FS_5G_ProSe_Test", "Study on 5G ProSe Relay", "SP-250123", "Rel-19", "2024-01-01", "2024-12-01"),
        )
        conn.execute(
            "INSERT OR IGNORE INTO wi_group_map (wi_code, group_id) VALUES ('980042', 1)"
        )
        conn.commit()

    # 4. Protocol messages
    conn = sqlite3.connect(str(proto_db))
    conn.execute(
        """
        CREATE TABLE spec_versions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            spec_number TEXT NOT NULL,
            version TEXT NOT NULL,
            spec_type TEXT DEFAULT 'NAS'
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE nas_messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            version_id INTEGER NOT NULL,
            clause TEXT NOT NULL,
            message_name TEXT NOT NULL,
            table_caption TEXT,
            description TEXT DEFAULT '',
            direction TEXT DEFAULT ''
        )
        """
    )
    conn.execute(
        "INSERT INTO spec_versions (id, spec_number, version, spec_type) VALUES (1, '24.501', '18.5.0', 'NAS')"
    )
    conn.execute(
        "INSERT INTO nas_messages (version_id, clause, message_name) VALUES (1, '8.2.6', 'REGISTRATION REQUEST')"
    )
    conn.commit()
    conn.close()

    return data_db, proto_db


class TestPhase3CommandPalette(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_nav_rail_cmdk_button_and_focus_invariant(self):
        rail = NavRail()
        self.assertEqual(len(rail.ordered_focus_widgets()), 10)
        self.assertTrue(hasattr(rail, "cmdk_btn"))
        self.assertNotIn(rail.cmdk_btn, rail.ordered_focus_widgets())
        self.assertIn("Ctrl+K", rail.cmdk_btn.text())

        clicked = []
        rail.command_palette_requested.connect(lambda: clicked.append(True))
        rail.cmdk_btn.click()
        self.assertEqual(clicked, [True])

        rail.set_collapsed(True)
        self.assertEqual(rail.cmdk_btn.text(), "")
        rail.set_collapsed(False)
        self.assertIn("Ctrl+K", rail.cmdk_btn.text())

    def test_dialog_keyboard_navigation_filtering_and_execution(self):
        executed = []
        items = [
            CommandItem("t1", "Tools", "Alpha Tool", "First", "Ctrl+1", callback=lambda: executed.append("t1")),
            CommandItem("t2", "Tools", "Beta Tool", "Second", "Ctrl+2", callback=lambda: executed.append("t2")),
            CommandItem("a1", "Side Panels & Actions", "Dock Assistant", "Panel", "Ctrl+Shift+A", callback=lambda: executed.append("a1")),
        ]
        dlg = CommandPaletteDialog(items)
        dlg.show()
        self.app.processEvents()

        # First item row is automatically selected (skipping header)
        cur = dlg.results_list.currentItem()
        self.assertIsNotNone(cur)
        self.assertEqual(cur.data(_ROLE_KIND), _KIND_ITEM)

        # Down arrow moves to next item (skipping any headers)
        down_ev = QKeyEvent(QEvent.KeyPress, Qt.Key_Down, Qt.NoModifier)
        dlg.eventFilter(dlg.search_input, down_ev)
        self.assertEqual(dlg.results_list.currentItem().text(), "Beta Tool")

        # Filter to Dock Assistant and hit Return
        dlg.search_input.setText("assistant")
        self.app.processEvents()
        self.assertEqual(dlg.results_list.currentItem().text(), "Dock Assistant")

        ret_ev = QKeyEvent(QEvent.KeyPress, Qt.Key_Return, Qt.NoModifier)
        dlg.eventFilter(dlg.search_input, ret_ev)
        self.app.processEvents()

        self.assertEqual(executed, ["a1"])
        self.assertEqual(dlg.result(), CommandPaletteDialog.Accepted)
        dlg.close()

    def test_dynamic_sqlite_providers_and_filter_safe_deep_links(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            data_db, proto_db = _seed_temp_dbs(tmp_dir)

            specs_opened = []
            mtgs_opened = []
            wis_opened = []
            proto_opened = []
            searcher = DynamicPaletteSearcher(
                db_path=data_db,
                protocol_db_path=proto_db,
                on_open_spec=specs_opened.append,
                on_open_meeting=mtgs_opened.append,
                on_open_work_item=wis_opened.append,
                on_open_protocol=lambda p, m: proto_opened.append((p, m)),
            )

            spec_hits = [x for x in searcher.search("24.501") if x.category == "Specifications"]
            self.assertEqual(len(spec_hits), 1)
            spec_hits[0].callback()
            self.assertEqual(specs_opened, ["24.501"])

            tdoc_hits = [x for x in searcher.search("S2-251234") if x.category == "Meetings & TDocs"]
            self.assertTrue(any("TDoc S2-251234" in x.title for x in tdoc_hits))
            tdoc_hits[0].callback()
            self.assertEqual(mtgs_opened, ["S2-251234"])

            wi_hits = [x for x in searcher.search("ProSe") if x.category == "Work Items"]
            self.assertEqual(len(wi_hits), 1)
            wi_hits[0].callback()
            self.assertEqual(wis_opened, ["FS_5G_ProSe_Test"])

            proto_hits = [x for x in searcher.search("REGISTRATION") if x.category == "Protocols"]
            self.assertEqual(len(proto_hits), 1)
            proto_hits[0].callback()
            self.assertEqual(proto_opened, [("NAS", "REGISTRATION REQUEST")])

            # --- Filter-safe UI helpers on real tab widgets ---
            specs_tab = SpecificationsTab(data_db)
            specs_tab.favorites = set()
            specs_tab.favorites_only = True
            specs_tab.fav_chip.setChecked(True)  # Hide non-favorites
            specs_tab.select_and_show_spec("24.501")
            self.app.processEvents()
            self.assertFalse(specs_tab.favorites_only)
            self.assertGreaterEqual(specs_tab.table.rowCount(), 1)
            specs_tab.close()

            mtgs_tab = MeetingsTab(data_db)
            mtgs_tab.search_input.setText("NONEXISTENT_MEETING")
            mtgs_tab.refresh_table()
            self.assertEqual(mtgs_tab.table_model.rowCount(), 0)
            mtgs_tab.focus_meeting("SA2#168")
            self.app.processEvents()
            self.assertGreaterEqual(mtgs_tab.table_model.rowCount(), 1)
            mtgs_tab.close()

            wi_tab = WorkItemsTab(data_db)
            wi_tab.following_only = True
            wi_tab.refresh_table(reset_pagination=True)
            self.assertEqual(wi_tab.table_model.rowCount(), 0)
            wi_tab.focus_work_item("FS_5G_ProSe_Test")
            self.app.processEvents()
            self.assertGreaterEqual(wi_tab.table_model.rowCount(), 1)
            wi_tab.close()

            word_tab = WordExtractorTab()
            word_tab.open_workbench("cr_pcr")
            self.assertEqual(word_tab.stack.currentIndex(), 7)
            word_tab.open_workbench("sanitizer")
            self.assertEqual(word_tab.stack.currentIndex(), 8)
            word_tab.close()

    def test_static_builder_wires_tools_and_dock_actions(self):
        host = QMainWindow()
        host.tabs = QTabWidget(host)
        for i in range(9):
            host.tabs.addTab(QWidget(host), f"Tab {i}")
        host.right_panel = RightSidePanel(host)
        host.open_right_panel = host.right_panel.open_mode
        host.word_tab = WordExtractorTab()
        host.nav_rail = NavRail(host)

        static_items = build_static_command_items(host)
        by_id = {c.id: c for c in static_items}

        by_id["action:dock_assistant"].callback()
        self.assertTrue(host.right_panel.is_open)
        self.assertEqual(host.right_panel.mode, MODE_ASSISTANT)

        by_id["wb:cr_pcr"].callback()
        self.assertEqual(host.tabs.currentIndex(), 2)
        self.assertEqual(host.word_tab.stack.currentIndex(), 7)
        host.word_tab.close()
        host.close()


if __name__ == "__main__":
    unittest.main()
