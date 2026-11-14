import sys
import tempfile
import unittest
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtWidgets import QApplication, QMainWindow, QSplitter, QTabWidget, QVBoxLayout, QWidget

from core.ui.right_side_panel import MODE_ASSISTANT, MODE_SPEC_SEARCH, RightSidePanel
from main_window import DragDropUI
from modules.assistant.core.agent_models import (
    AgentResult,
    Evidence,
    EvidenceSourceType,
    ResearchOutcome,
)
from modules.assistant.ui.assistant_tab import AssistantTab


class _MockAssistantDockHost(QMainWindow):
    def __init__(self, tmp_dir: Path):
        super().__init__()
        self._assistant_docked = False

        central = QWidget(self)
        layout = QVBoxLayout(central)
        self.workspace_splitter = QSplitter(Qt.Horizontal, self)
        self.workspace_splitter.setChildrenCollapsible(False)

        self.tabs = QTabWidget(self)
        self.other_tab = QWidget(self)
        self.assistant_tab = AssistantTab(
            specs_db_path=tmp_dir / "specs.db",
            spec_search_db_path=tmp_dir / "spec_search.db",
            protocol_db_path=tmp_dir / "proto.db",
        )
        self.tabs.addTab(self.other_tab, "Other")
        self.tabs.addTab(self.assistant_tab, "Assistant")

        self.right_panel = RightSidePanel(self)
        self.workspace_splitter.addWidget(self.tabs)
        self.workspace_splitter.addWidget(self.right_panel)
        self.workspace_splitter.setStretchFactor(0, 1)
        self.workspace_splitter.setStretchFactor(1, 0)
        layout.addWidget(self.workspace_splitter)
        self.setCentralWidget(central)

        self.tabs.currentChanged.connect(self._on_tab_changed)
        self.right_panel.toggled.connect(self._on_panel_state)
        self.right_panel.mode_changed.connect(lambda _m: self._on_panel_state(self.right_panel.is_open))
        self.right_panel.promote_requested.connect(self._on_promote)
        self.assistant_tab.undock_requested.connect(lambda: self._on_promote(MODE_ASSISTANT))

    def _set_assistant_docked(self, docked: bool):
        DragDropUI._set_assistant_docked(self, docked)

    def _sync_assistant_dock_slot(self):
        DragDropUI._sync_assistant_dock_slot(self)

    def _on_panel_state(self, _opened: bool):
        self._sync_assistant_dock_slot()
        self.right_panel.apply_splitter_sizes(self.workspace_splitter)

    def _on_promote(self, mode: str):
        DragDropUI._on_right_panel_promote(self, mode)

    def _on_tab_changed(self, index: int):
        DragDropUI._on_tab_changed(self, index)


class TestPhase3DockedAssistant(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_single_instance_reparent_compact_mode_and_citations(self):
        with tempfile.TemporaryDirectory() as tmp:
            host = _MockAssistantDockHost(Path(tmp))
            host.resize(1366, 768)
            host.show()
            self.app.processEvents()

            tab = host.assistant_tab
            self.assertIs(host.tabs.widget(1), tab)
            self.assertFalse(tab.is_docked)
            self.assertTrue(tab.ribbon_card.isVisibleTo(tab.body_container))
            self.assertFalse(tab.dock_compact_strip.isVisibleTo(tab.body_container))
            self.assertFalse(tab.docked_placeholder_card.isVisibleTo(tab))

            # Populate conversation + citation before docking
            tab._append_conversation("user", "What is PEI?")
            ev = Evidence(
                id="E1",
                source_type=EvidenceSourceType.SPECIFICATION_TEXT,
                source_name="23.501",
                specification="23.501",
                version="18.4.0",
                clause="5.3.2",
                title="Permanent Equipment Identifier",
                content="PEI definition text [E1].",
            )
            res = AgentResult(
                outcome=ResearchOutcome.COMPLETED,
                answer="PEI is Permanent Equipment Identifier [E1].",
                evidence=[ev],
            )
            tab._on_result(res)
            self.app.processEvents()

            # Dock Assistant into RightSidePanel
            host.right_panel.open_mode(MODE_ASSISTANT)
            self.app.processEvents()

            # Single-instance checks
            self.assertIs(host.tabs.widget(1), tab)
            self.assertTrue(tab.is_docked)
            self.assertIs(host.right_panel.slot_widget(MODE_ASSISTANT), tab.body_container)
            self.assertIs(host.right_panel.stack.currentWidget(), tab.body_container)
            self.assertFalse(tab.ribbon_card.isVisibleTo(tab.body_container))
            self.assertTrue(tab.dock_compact_strip.isVisible())
            self.assertFalse(tab.docked_placeholder_card.isHidden())
            self.assertEqual(tab.main_splitter.orientation(), Qt.Vertical)
            self.assertFalse(tab.cockpit_tabs.isVisible())
            self.assertEqual(tab.dock_sources_toggle_btn.text(), "📚 Sources (1)")
            self.assertEqual(len(tab._conversation), 2)

            # Clicking [E1] citation auto-expands Sources cockpit
            tab._on_anchor_clicked(QUrl("evidence:E1"))
            self.app.processEvents()
            self.assertTrue(tab.cockpit_tabs.isVisible())
            self.assertTrue(tab.dock_sources_toggle_btn.isChecked())
            self.assertEqual(tab.cockpit_tabs.currentIndex(), 0)

            # Toggling sources button collapses cockpit again
            tab.dock_sources_toggle_btn.click()
            self.app.processEvents()
            self.assertFalse(tab.cockpit_tabs.isVisible())

            # Switching right panel to Spec Search restores Assistant back to its tab
            host.right_panel.open_mode(MODE_SPEC_SEARCH)
            self.app.processEvents()
            self.assertFalse(tab.is_docked)
            self.assertIs(tab.body_container.parentWidget(), tab)
            self.assertTrue(tab.ribbon_card.isVisibleTo(tab.body_container))
            self.assertTrue(tab.docked_placeholder_card.isHidden())

            # Re-dock then click tab 1 -> auto-undocks and closes right dock
            host.right_panel.open_mode(MODE_ASSISTANT)
            self.app.processEvents()
            self.assertTrue(tab.is_docked)
            self.assertIs(host.right_panel.stack.currentWidget(), tab.body_container)
            host.tabs.setCurrentIndex(1)
            self.app.processEvents()
            self.assertFalse(tab.is_docked)
            self.assertFalse(host.right_panel.is_open)
            self.assertTrue(tab.ribbon_card.isVisible())
            self.assertEqual(len(tab._conversation), 2)

            host.close()


if __name__ == "__main__":
    unittest.main()
