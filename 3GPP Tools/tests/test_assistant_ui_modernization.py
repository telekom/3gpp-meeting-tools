"""Unit tests for the modernized 3GPP Assistant UI components and integration."""

import unittest
from pathlib import Path
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QKeyEvent
from PyQt5.QtWidgets import QApplication, QPushButton

# Ensure QApplication exists
app = QApplication.instance()
if app is None:
    app = QApplication([])

from modules.assistant.core.agent_models import (
    AgentResult,
    Evidence,
    EvidenceSourceType,
    ResearchActivity,
    ResearchOutcome,
)
from modules.assistant.ui.assistant_components import (
    ActivityTimelineWidget,
    DiagnosticsTraceWidget,
    EvidenceInspectorWidget,
    MultilinePromptEdit,
    StarterChipsWidget,
    render_assistant_message_card,
    render_markdown_to_html,
    render_system_message_card,
    render_user_message_card,
)
from modules.assistant.ui.assistant_tab import AssistantTab


class TestAssistantUIComponents(unittest.TestCase):
    """Tests individual assistant presentation components."""

    def test_multiline_prompt_edit_enter_and_shift_enter(self):
        edit = MultilinePromptEdit()
        submitted = []
        edit.submit_requested.connect(lambda: submitted.append(True))

        # Plain Enter: triggers submit_requested
        enter_event = QKeyEvent(QKeyEvent.KeyPress, Qt.Key_Return, Qt.NoModifier)
        edit.keyPressEvent(enter_event)
        self.assertEqual(len(submitted), 1)

        # Shift + Enter: does NOT trigger submit_requested
        shift_enter_event = QKeyEvent(QKeyEvent.KeyPress, Qt.Key_Return, Qt.ShiftModifier)
        edit.keyPressEvent(shift_enter_event)
        self.assertEqual(len(submitted), 1)

    def test_starter_chips_widget(self):
        chips = StarterChipsWidget()
        selected = []
        chips.prompt_selected.connect(selected.append)

        chip_btns = chips.findChildren(QPushButton)
        self.assertGreaterEqual(len(chip_btns), 4)
        chip_btns[0].click()
        self.assertEqual(len(selected), 1)
        self.assertTrue("Registration" in selected[0])

    def test_markdown_and_citations_rendering(self):
        md = """### Protocol Analysis
According to **TS 29.244** Clause 7.5.2 [E1], the message contains:
- **Node ID**: Mandatory [E1]
- **F-SEID**: Mandatory [E2]

```text
PFCP Session Establishment Request
```
"""
        html_out = render_markdown_to_html(md, valid_evidence_ids={"E1", "E2"})
        self.assertIn("Protocol Analysis", html_out)
        self.assertIn("<b>TS 29.244</b>", html_out)
        self.assertIn('href="evidence:E1"', html_out)
        self.assertIn('href="evidence:E2"', html_out)
        self.assertIn("PFCP Session Establishment Request", html_out)

    def test_message_cards_rendering(self):
        user_card = render_user_message_card("What is Registration?", timestamp="12:00")
        self.assertIn("👤 You", user_card)
        self.assertIn("What is Registration?", user_card)
        self.assertIn("12:00", user_card)

        ai_card = render_assistant_message_card(
            "It is a 5GC procedure [E1].",
            timestamp="12:01",
            valid_evidence_ids={"E1"},
            msg_index=0,
            kpi_badge="⏱️ 1.2s · 2 steps",
        )
        self.assertIn("🤖 3GPP Assistant", ai_card)
        self.assertIn('href="evidence:E1"', ai_card)
        self.assertIn('href="copy_msg:0"', ai_card)
        self.assertIn("1.2s", ai_card)

    def test_evidence_inspector_widget(self):
        inspector = EvidenceInspectorWidget()
        self.assertFalse(inspector.copy_btn.isEnabled())

        ev = Evidence(
            id="E1",
            source_type=EvidenceSourceType.SPECIFICATION_TEXT,
            source_name="TS 23.501",
            specification="23.501",
            version="18.4.0",
            clause="5.2.2",
            title="Registration Management",
            content="Registration management procedure...",
        )
        inspector.set_evidence(ev)
        self.assertTrue(inspector.copy_btn.isEnabled())
        self.assertFalse(inspector.spec_search_btn.isHidden())
        self.assertFalse(inspector.specs_tab_btn.isHidden())
        self.assertIn("Registration management procedure...", inspector.content_browser.toPlainText())

    def test_activity_timeline_widget(self):
        timeline = ActivityTimelineWidget()
        act = ResearchActivity(
            kind="tool_completed",
            message="get_spec_clause: found.",
            metadata={"tool": "get_spec_clause", "evidence_ids": ["E1"]},
        )
        timeline.add_activity(act)
        timeline.update_count(1)
        self.assertIn("1 step executed", timeline.step_label.text())
        self.assertIn("get_spec_clause", timeline.browser.toHtml())
        self.assertIn("E1", timeline.browser.toHtml())

    def test_diagnostics_trace_widget(self):
        diag = DiagnosticsTraceWidget()
        diag.add_exchange_event({
            "direction": "request",
            "method": "POST",
            "endpoint": "/api/chat",
            "payload": {"model": "qwen2.5:7b"},
            "timestamp": "12:00:00.123",
        })
        diag.update_metrics(total_seconds=2.5, prompt_tokens=400, eval_tokens=150, eval_seconds=3.0)
        self.assertIn("2.5s", diag.kpi_latency.text())
        self.assertIn("50.0 tok/s", diag.kpi_speed.text())
        self.assertIn("/api/chat", diag.browser.toPlainText())


class TestAssistantTabIntegration(unittest.TestCase):
    """Tests AssistantTab orchestration and signals."""

    def setUp(self):
        self.tab = AssistantTab(
            specs_db_path=Path("src/db/3gpp_data.db"),
            spec_search_db_path=Path("src/db/3gpp_spec_search.db"),
            protocol_db_path=Path("src/db/3gpp_protocol_data.db"),
        )

    def test_layout_toggle(self):
        # Default is Side-by-Side (Qt.Horizontal)
        self.assertEqual(self.tab.main_splitter.orientation(), Qt.Horizontal)
        self.assertEqual(self.tab.layout_toggle_btn.text(), "⬒ Bottom Split")

        # Toggle to Bottom Split
        self.tab.layout_toggle_btn.setChecked(False)
        self.tab._toggle_layout(False)
        self.assertEqual(self.tab.main_splitter.orientation(), Qt.Vertical)
        self.assertEqual(self.tab.layout_toggle_btn.text(), "◫ Side Drawer")

        # Toggle back to Side-by-Side
        self.tab.layout_toggle_btn.setChecked(True)
        self.tab._toggle_layout(True)
        self.assertEqual(self.tab.main_splitter.orientation(), Qt.Horizontal)

    def test_anchor_click_and_evidence_selection(self):
        ev = Evidence(
            id="E1",
            source_type=EvidenceSourceType.SPECIFICATION_TEXT,
            source_name="TS 23.501",
            specification="23.501",
            version="18.4.0",
            clause="5.2.2",
            title="Registration Management",
            content="Registration content",
        )
        self.tab._on_evidence(ev)
        self.assertEqual(self.tab.sources_tree.topLevelItemCount(), 1)

        # Simulate clicking on an evidence citation link
        self.tab._on_anchor_clicked(QUrl("evidence:E1"))
        self.assertEqual(self.tab.cockpit_tabs.currentIndex(), 0)
        self.assertEqual(self.tab.evidence_inspector._current_evidence.id, "E1")

    def test_session_new_chat_reset(self):
        self.tab._append_conversation("user", "Hello 3GPP")
        self.assertEqual(len(self.tab._conversation), 1)

        # New chat resets session
        self.tab._on_new_chat()
        self.assertEqual(len(self.tab._conversation), 0)
        self.assertEqual(self.tab.sources_tree.topLevelItemCount(), 0)
        self.assertFalse(self.tab.starter_chips.isHidden())


if __name__ == "__main__":
    unittest.main()
