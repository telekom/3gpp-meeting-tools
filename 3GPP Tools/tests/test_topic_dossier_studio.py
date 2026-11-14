# --- File: tests/test_topic_dossier_studio.py ---
"""
Unit and Integration Tests for Cross-WG Topic Dossier & Briefing Studio.

Tests:
1. MetaTopicDefinition & MetaTopicRules matching logic (WI tokens, AI patterns, keywords, pins, exclusions).
2. MetaTopicManager CRUD and JSON persistence.
3. MeetingCohortResolver date parsing and cross-WG quarterly clustering.
4. TopicDecision unified modeling (RAN text agreements + SA/CT pCRs/CRs).
5. CrossWgLiaisonTracker dialogue thread building and Markdown rendering.
6. Incremental export and LLM delta feed generation.
"""

import json
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

# Ensure src is in sys.path
_src_dir = Path(__file__).resolve().parent.parent / "src"
if str(_src_dir) not in sys.path:
    sys.path.insert(0, str(_src_dir))

from modules.meetings.core.cohort_resolver import (
    MeetingCohort,
    MeetingCohortResolver,
    get_cohort_key,
    parse_cohort_key,
    compare_cohort_keys,
)
from modules.meetings.core.cross_wg_ls_tracker import (
    CrossWgLiaisonTracker,
    LiaisonThreadNode,
)
from modules.meetings.core.topic_decision import (
    TopicDecision,
    classify_tdoc_decision_type,
    normalize_chair_notes_decision_type,
)
from modules.meetings.core.topic_manager import (
    MetaTopicDefinition,
    MetaTopicManager,
    MetaTopicRules,
    _extract_base_tokens,
)
from modules.meetings.core.exporters.export_manifest import ExportManifest


class TestMetaTopicMatching(unittest.TestCase):
    def setUp(self):
        self.rules = MetaTopicRules(
            wi_tokens=["ambientiot", "ambientiot_ph2", "fs_ambientiot"],
            ai_patterns={
                "SA2": ["5.14*", "5.12*"],
                "RAN2": ["8.7*"],
                "CT1": ["6.18*"],
            },
            keywords=["ambient iot", "zero power", "passive iot"],
            pinned_tdocs=["S2-2599999"],
            excluded_tdocs=["S2-2500001"],
        )
        self.topic = MetaTopicDefinition(
            id="test_ambient",
            name="Ambient IoT",
            target_wgs=["SA2", "RAN2", "CT1"],
            rules=self.rules,
        )

    def test_pinned_and_excluded_tdocs(self):
        # Pinned matches regardless of rules
        self.assertTrue(self.topic.matches_tdoc("S2-2599999", "SA2", "99.9", "Unrelated Title"))
        # Excluded never matches even if WI matches
        self.assertFalse(self.topic.matches_tdoc("S2-2500001", "SA2", "5.14", "Ambient IoT", "AmbientIoT_Ph2"))

    def test_wi_token_matching(self):
        # Exact match
        self.assertTrue(self.topic.matches_tdoc("S2-2501000", "SA2", "9.1", "Some doc", "AmbientIoT_Ph2-CT"))
        # Study item prefix
        self.assertTrue(self.topic.matches_tdoc("R2-2501000", "RAN2", "3.1", "Radio flow", "FS_AmbientIoT"))
        # Unrelated WI
        self.assertFalse(self.topic.matches_tdoc("S2-2501002", "SA2", "9.1", "Unrelated", "NR_AIML_Ph2"))

    def test_ai_pattern_matching(self):
        # SA2 AI match
        self.assertTrue(self.topic.matches_tdoc("S2-2501003", "SA2", "5.14.2", "Discovery flow"))
        self.assertTrue(self.topic.matches_tdoc("S2-2501004", "SA2", "5.12", "Architecture"))
        # Non-matching SA2 AI
        self.assertFalse(self.topic.matches_tdoc("S2-2501005", "SA2", "6.2", "Unrelated"))
        # RAN2 AI match
        self.assertTrue(self.topic.matches_tdoc("R2-2501006", "RAN2", "8.7.1", "Radio baseline"))
        # CT1 AI match
        self.assertTrue(self.topic.matches_tdoc("C1-2501007", "CT1", "6.18", "URSP rules"))

    def test_keyword_matching(self):
        self.assertTrue(self.topic.matches_tdoc("S2-2501008", "SA2", "9.9", "Evaluation of zero power devices"))
        self.assertTrue(self.topic.matches_tdoc("R2-2501009", "RAN2", "9.9", "Study on Passive IoT protocols"))
        self.assertFalse(self.topic.matches_tdoc("S2-2501010", "SA2", "9.9", "Completely unrelated core topic"))

    def test_ai_description_matching(self):
        # AI description matches keywords even if AI number and title don't explicitly mention it
        self.assertTrue(self.topic.matches_tdoc("S2-2501011", "SA2", "9.9", "General contribution", ai_description="Discussion on ambient iot framework"))
        self.assertFalse(self.topic.matches_tdoc("S2-2501012", "SA2", "9.9", "General contribution", ai_description="Voice and emergency services"))


class TestMetaTopicManager(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.config_path = Path(self.temp_dir) / "test_meta_topics.json"
        self.manager = MetaTopicManager(config_path=self.config_path)

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_default_presets_created(self):
        self.assertTrue(self.config_path.exists())
        topics = self.manager.list_topics()
        self.assertGreaterEqual(len(topics), 2)
        ambient = self.manager.get_topic("rel19_ambient_iot")
        self.assertIsNotNone(ambient)
        self.assertEqual(ambient.name, "Ambient IoT (Rel-19)")

    def test_save_and_reload(self):
        new_topic = MetaTopicDefinition(
            id="custom_topic",
            name="Custom Slicing Topic",
            description="End-to-end slicing SLA",
            target_wgs=["SA2", "SA5"],
        )
        self.manager.save_topic(new_topic)

        # Reload manager from disk
        mgr2 = MetaTopicManager(config_path=self.config_path)
        loaded = mgr2.get_topic("custom_topic")
        self.assertIsNotNone(loaded)
        self.assertEqual(loaded.name, "Custom Slicing Topic")
        self.assertEqual(loaded.target_wgs, ["SA2", "SA5"])

    def test_clone_and_delete(self):
        cloned = self.manager.clone_topic("rel19_ambient_iot")
        self.assertIsNotNone(cloned)
        self.assertIn("Copy", cloned.name)

        del_ok = self.manager.delete_topic(cloned.id)
        self.assertTrue(del_ok)
        self.assertIsNone(self.manager.get_topic(cloned.id))

    def test_deletion_persistence_never_reappears(self):
        """Verify that deleted topics do NOT reappear when manager is reloaded."""
        # Delete built-in ambient IoT
        self.assertIsNotNone(self.manager.get_topic("rel19_ambient_iot"))
        ok = self.manager.delete_topic("rel19_ambient_iot")
        self.assertTrue(ok)
        self.assertIsNone(self.manager.get_topic("rel19_ambient_iot"))

        # Re-instantiate manager pointing to same config
        mgr2 = MetaTopicManager(config_path=self.config_path)
        self.assertIsNone(mgr2.get_topic("rel19_ambient_iot"), "Deleted topic must NOT reappear on reload")

    def test_empty_topics_persistence(self):
        """Verify that deleting all topics stays empty and does not recreate defaults."""
        for t in list(self.manager.list_topics()):
            self.manager.delete_topic(t.id)

        self.assertEqual(len(self.manager.list_topics()), 0)

        # Reload manager
        mgr2 = MetaTopicManager(config_path=self.config_path)
        self.assertEqual(len(mgr2.list_topics()), 0, "Empty topic list must persist without re-seeding defaults")

    def test_restore_default_topics(self):
        """Verify non-destructive restore of built-in factory topics."""
        # Add custom topic
        custom = MetaTopicDefinition(id="my_custom", name="My Custom Topic")
        self.manager.save_topic(custom)

        # Delete all other topics
        for t in list(self.manager.list_topics()):
            if t.id != "my_custom":
                self.manager.delete_topic(t.id)

        self.assertEqual(len(self.manager.list_topics()), 1)
        self.assertIsNotNone(self.manager.get_topic("my_custom"))

        # Restore defaults
        restored_count = self.manager.restore_default_topics(overwrite=False)
        self.assertGreater(restored_count, 0)

        # Custom topic must still be intact
        self.assertIsNotNone(self.manager.get_topic("my_custom"))
        # Factory presets should now be back
        self.assertIsNotNone(self.manager.get_topic("rel19_ambient_iot"))
        self.assertIsNotNone(self.manager.get_topic("rel20_data_framework"))


class TestMeetingCohortResolver(unittest.TestCase):
    def test_cohort_key_parsing(self):
        self.assertEqual(get_cohort_key("2025-10-24"), "2025-Q4")
        self.assertEqual(get_cohort_key("2025-02-15"), "2025-Q1")
        self.assertEqual(get_cohort_key("2025-05-18"), "2025-Q2")
        self.assertEqual(get_cohort_key("2025-08-20"), "2025-Q3")
        self.assertEqual(parse_cohort_key("2025-Q4"), (2025, 4))
        self.assertEqual(compare_cohort_keys("2025-Q3", "2025-Q4"), -1)
        self.assertEqual(compare_cohort_keys("2026-Q1", "2025-Q4"), 1)

    def test_cluster_staggered_meetings(self):
        # SA2 meets late October, RAN2 and CT1 meet mid-November (different weeks, same Q4 cohort)
        sample_meetings = [
            {"wg_name": "SA2", "meeting_number": "160", "start_date": "2025-10-20", "end_date": "2025-10-24"},
            {"wg_name": "RAN2", "meeting_number": "129", "start_date": "2025-11-10", "end_date": "2025-11-14"},
            {"wg_name": "CT1", "meeting_number": "146", "start_date": "2025-11-17", "end_date": "2025-11-21"},
            # Next cycle: Q1 2026
            {"wg_name": "SA2", "meeting_number": "161", "start_date": "2026-02-16", "end_date": "2026-02-20"},
            {"wg_name": "RAN2", "meeting_number": "130", "start_date": "2026-02-23", "end_date": "2026-02-27"},
        ]
        cohorts = MeetingCohortResolver.cluster_meetings_into_cohorts(sample_meetings)
        self.assertEqual(len(cohorts), 2)
        c_q4 = cohorts[0]
        self.assertEqual(c_q4.key, "2025-Q4")
        self.assertEqual(c_q4.wgs_present, {"SA2", "RAN2", "CT1"})
        self.assertEqual(len(c_q4.meetings), 3)

        c_q1 = cohorts[1]
        self.assertEqual(c_q1.key, "2026-Q1")
        self.assertEqual(c_q1.wgs_present, {"SA2", "RAN2"})


class TestTopicDecisionModel(unittest.TestCase):
    def test_unified_rendering(self):
        # RAN Text Agreement
        ran_dec = TopicDecision(
            cohort_key="2025-Q4",
            wg="RAN2",
            meeting_number="RAN2 #129",
            date="2025-11-12",
            decision_type="Text Agreement",
            title_or_summary="Agreed: Preamble sequence A selected for direct discovery",
            source="Chairman's Notes §8.7",
            target_spec="TR 38.8xx",
        )
        row1 = ran_dec.to_markdown_row()
        self.assertIn("| **RAN2** |", row1)
        self.assertIn("🟢 **Text Agreement**", row1)
        self.assertIn("TR 38.8xx", row1)

        # SA2 Approved pCR
        sa2_dec = TopicDecision(
            cohort_key="2025-Q4",
            wg="SA2",
            meeting_number="SA2 #160",
            date="2025-10-24",
            decision_type="Approved pCR",
            title_or_summary="Solution 4: Direct discovery architecture flows",
            tdoc_id="S2-2501234",
            source="Qualcomm, Nokia",
            target_spec="TR 23.700-99",
            dossier_rel_path="2025-Q4/SA2_160/pCR/S2-2501234.md",
        )
        row2 = sa2_dec.to_markdown_row()
        self.assertIn("| **SA2** |", row2)
        self.assertIn("🟢 **Approved pCR**", row2)
        self.assertIn("[S2-2501234](2025-Q4/SA2_160/pCR/S2-2501234.md)", row2)

    def test_classification(self):
        self.assertEqual(classify_tdoc_decision_type("pCR", "Draft pCR on TR", "Agreed"), "Approved pCR")
        self.assertEqual(classify_tdoc_decision_type("CR", "CR on TS 38.331", "Approved"), "Approved CR")
        self.assertEqual(classify_tdoc_decision_type("Discussion", "Joint WF on Ambient IoT", "Agreed"), "Endorsed WF")
        self.assertEqual(classify_tdoc_decision_type("Discussion", "Regular discussion paper", "Noted"), None)
        self.assertEqual(normalize_chair_notes_decision_type("Working Assumption"), "Working Assumption")
        self.assertEqual(normalize_chair_notes_decision_type("FFS"), "Open Issue / FFS")


class TestCrossWgLiaisonTracker(unittest.TestCase):
    def test_threads_building(self):
        tracker = CrossWgLiaisonTracker()

        # SA2 sends LS OUT to RAN2 and CT1
        tracker.ingest_ls_tdoc(
            tdoc_id="S2-2501999",
            wg="SA2",
            meeting_number="SA2 #160",
            cycle="2025-Q4",
            tdoc_type="LS out",
            title="LS on Ambient IoT direct discovery to RAN2, CT1",
            source="SA2",
            status="Approved",
            dossier_rel_path="2025-Q4/SA2_160/LS_OUT/S2-2501999.md",
        )

        # RAN2 receives it as LS IN
        tracker.ingest_ls_tdoc(
            tdoc_id="R2-2502100",
            wg="RAN2",
            meeting_number="RAN2 #129",
            cycle="2025-Q4",
            tdoc_type="LS in",
            title="LS IN from SA2 on Ambient IoT",
            source="SA2",
            status="Handled",
            reply_to="S2-2501999",
            dossier_rel_path="2025-Q4/RAN2_129/LS_IN/R2-2502100.md",
        )

        # RAN2 approves reply LS OUT back to SA2
        tracker.ingest_ls_tdoc(
            tdoc_id="R2-2503100",
            wg="RAN2",
            meeting_number="RAN2 #129",
            cycle="2025-Q4",
            tdoc_type="LS out",
            title="Reply LS to SA2 on direct discovery (reply to S2-2501999)",
            source="RAN2",
            status="Approved",
            reply_to="S2-2501999",
            dossier_rel_path="2025-Q4/RAN2_129/LS_OUT/R2-2503100.md",
        )

        md = tracker.generate_markdown_report()
        self.assertIn("Cross-WG Liaison Statement (LS) Dialogue Tracker", md)
        self.assertIn("S2-2501999", md)
        self.assertIn("R2-2503100", md)
        self.assertIn("Dialogue Threads", md)


class TestTopicDossierExporterIncremental(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.target_dir = Path(self.temp_dir) / "TestTopicDossier"

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_incremental_export_and_deltas(self):
        from unittest.mock import MagicMock, patch
        from PyQt5.QtWidgets import QApplication
        from modules.meetings.core.exporters.topic_dossier_exporter import TopicDossierExporterThread

        topic = MetaTopicDefinition(
            id="test_topic",
            name="Test Ambient IoT",
            target_wgs=["SA2", "RAN2"],
            rules=MetaTopicRules(wi_tokens=["ambientiot"]),
        )
        cohort = MeetingCohort(
            key="2025-Q4",
            year=2025,
            quarter=4,
            label="2025 Q4",
            meetings=[
                {"wg_name": "SA2", "meeting_number": "160", "start_date": "2025-10-20", "folder_name": "SP-160"},
                {"wg_name": "RAN2", "meeting_number": "129", "start_date": "2025-11-10", "folder_name": "R2-129"},
            ],
            wgs_present={"SA2", "RAN2"},
        )
        mock_db = MagicMock()

        # Session 1: Initial Export
        exporter1 = TopicDossierExporterThread(
            meetings_db=mock_db,
            topic=topic,
            selected_cohorts=[cohort],
            output_base_dir=Path(self.temp_dir),
            target_folder_name="TestTopicDossier",
            generate_consolidated_md=True,
        )

        with patch.object(exporter1, "_load_meeting_tdocs") as mock_tdocs, \
             patch.object(exporter1, "_load_meeting_agreements") as mock_agreements, \
             patch.object(exporter1, "_convert_docs_pipeline"):

            mock_tdocs.side_effect = lambda mtg: [
                {"TDoc": "S2-2501001", "Title": "Solution 1", "AI": "5.14", "Related WIs": "AmbientIoT_Ph2", "Type": "pCR", "Status": "Agreed"},
                {"TDoc": "S2-2501999", "Title": "LS to RAN2", "AI": "5.14", "Related WIs": "AmbientIoT", "Type": "LS out", "Status": "Approved"},
            ] if mtg.get("wg_name") == "SA2" else [
                {"TDoc": "R2-2503456", "Title": "WF on Ambient IoT", "AI": "8.7", "Related WIs": "FS_AmbientIoT", "Type": "Discussion", "Status": "Agreed"},
            ]
            mock_agreements.side_effect = lambda mtg: [
                {"text": "Agreed: Option 1 is selected", "ai": "8.7", "decision_type": "Agreement", "wi": "FS_AmbientIoT"}
            ] if mtg.get("wg_name") == "RAN2" else []

            exporter1._run_export()

        self.assertTrue((self.target_dir / "index.md").is_file())
        self.assertTrue((self.target_dir / "agreements_timeline.md").is_file())
        self.assertTrue((self.target_dir / "ls_dialogue.md").is_file())
        self.assertTrue((self.target_dir / ".export_manifest.json").is_file())
        self.assertTrue((self.target_dir / "deltas" / "delta_session_1.md").is_file())
        self.assertTrue((self.target_dir / "consolidated_briefing.md").is_file())

        # Session 2: Incremental Export
        exporter2 = TopicDossierExporterThread(
            meetings_db=mock_db,
            topic=topic,
            selected_cohorts=[cohort],
            output_base_dir=Path(self.temp_dir),
            target_folder_name="TestTopicDossier",
            generate_consolidated_md=True,
        )

        with patch.object(exporter2, "_load_meeting_tdocs") as mock_tdocs2, \
             patch.object(exporter2, "_load_meeting_agreements") as mock_agreements2, \
             patch.object(exporter2, "_convert_docs_pipeline"):

            mock_tdocs2.side_effect = lambda mtg: [
                {"TDoc": "S2-2501001", "Title": "Solution 1", "AI": "5.14", "Related WIs": "AmbientIoT_Ph2", "Type": "pCR", "Status": "Agreed"},
                {"TDoc": "S2-2501999", "Title": "LS to RAN2", "AI": "5.14", "Related WIs": "AmbientIoT", "Type": "LS out", "Status": "Approved"},
                {"TDoc": "S2-2502000", "Title": "New Solution 2", "AI": "5.14", "Related WIs": "AmbientIoT_Ph2", "Type": "pCR", "Status": "Agreed"},
            ] if mtg.get("wg_name") == "SA2" else []
            mock_agreements2.side_effect = lambda mtg: []

            exporter2._run_export()

        self.assertTrue((self.target_dir / "deltas" / "delta_session_2.md").is_file())

    def test_default_topic_export_base(self):
        from modules.meetings.core.exporters.topic_dossier_exporter import get_default_topic_export_base
        from core.utils.paths import get_project_root
        base = get_default_topic_export_base()
        self.assertIsInstance(base, Path)
        self.assertTrue(str(base).endswith("Export"))
        # Verify it points to the user cache root and not into the src codebase
        self.assertNotEqual(base, get_project_root() / "Export")


class TestAgendaItemSuggester(unittest.TestCase):
    def setUp(self):
        from unittest.mock import MagicMock
        from modules.meetings.core.ai_suggester import AgendaItemSuggester, DiscoveredAi

        self.mock_db = MagicMock()
        self.suggester = AgendaItemSuggester(self.mock_db)

    def test_suggested_pattern_wildcard(self):
        from modules.meetings.core.ai_suggester import DiscoveredAi
        item1 = DiscoveredAi(
            wg="SA2", ai_number="20.6.21", ai_description="Data Framework",
            confidence="High", matching_reason="keyword match", tdoc_count=10, matching_tdocs=8
        )
        self.assertEqual(item1.suggested_pattern, "20.6.21*")

        item2 = DiscoveredAi(
            wg="RAN2", ai_number="9.3.3.1*", ai_description="Data Transfer",
            confidence="High", matching_reason="keyword match", tdoc_count=10, matching_tdocs=8
        )
        self.assertEqual(item2.suggested_pattern, "9.3.3.1*")

    def test_ai_suggestion_scoring_and_ranking(self):
        from unittest.mock import patch
        from modules.meetings.core.ai_suggester import AgendaItemSuggester

        mock_db = unittest.mock.MagicMock()
        suggester = AgendaItemSuggester(mock_db)

        # Mock meeting resolution and TDocs loading
        fake_meetings = [{"folder_name": "TSGS2_170", "meeting_number": "170", "wg_name": "SA2"}]
        fake_tdocs = [
            # Matching AI desc + WI + Title
            {"Agenda Item": "20.6.21", "AI Description": "6G Data Framework in SA WG2", "Title": "Architecture for 6G Data Framework", "Related WIs": "FS_6G_ARC"},
            {"Agenda Item": "20.6.21", "AI Description": "6G Data Framework in SA WG2", "Title": "Data storage mechanisms", "Related WIs": "FS_6G_ARC"},
            # Unrelated AI under same broad WI (no keyword in desc, no keyword in title)
            {"Agenda Item": "20.6.12", "AI Description": "Voice Services for 6G", "Title": "Emergency calling", "Related WIs": "FS_6G_ARC"},
            # AI with keyword only in TDoc title
            {"Agenda Item": "10.2.3", "AI Description": "Coordination with SA2", "Title": "LS on 6G data framework coordination", "Related WIs": "FS_6G_ARC"},
        ]

        with patch.object(suggester, "_resolve_recent_meetings", return_value=fake_meetings), \
             patch.object(suggester, "_load_meeting_tdocs", return_value=fake_tdocs):

            results = suggester.suggest_ais_for_topic(
                target_wgs=["SA2"],
                keywords=["data framework"],
                wi_tokens=["fs_6g_arc"],
            )

            sa2_ais = results.get("SA2", [])
            # 20.6.21 should be top and High confidence
            self.assertGreaterEqual(len(sa2_ais), 2)
            top = sa2_ais[0]
            self.assertEqual(top.ai_number, "20.6.21")
            self.assertEqual(top.confidence, "High")
            self.assertTrue("AI description matches" in top.matching_reason)

            # 10.2.3 should also be discovered due to title match
            ai_nums = [a.ai_number for a in sa2_ais]
            self.assertIn("10.2.3", ai_nums)

            # 20.6.12 (Voice Services) has 0 keyword matches and should NOT be suggested
            self.assertNotIn("20.6.12", ai_nums)

    def test_infer_topic_from_known_ais(self):
        from unittest.mock import patch
        from modules.meetings.core.ai_suggester import AgendaItemSuggester

        mock_db = unittest.mock.MagicMock()
        suggester = AgendaItemSuggester(mock_db)

        fake_meetings = [{"folder_name": "TSGS2_178", "meeting_number": "178", "wg_name": "SA2"}]
        fake_tdocs = [
            {
                "Agenda Item": "20.6.21",
                "Agenda item description": "Study on Architecture for 6G System (FS_6G_ARC) - 6G data framework in SA WG2",
                "Title": "Discussion on 6G data transfer mechanisms",
                "Related WIs": "FS_6G_ARC",
            },
            {
                "Agenda Item": "20.6.21",
                "Agenda item description": "Study on Architecture for 6G System (FS_6G_ARC) - 6G data framework in SA WG2",
                "Title": "Update on data framework architecture",
                "Related WIs": "FS_6G_ARC",
            },
        ]

        with patch.object(suggester, "_resolve_recent_meetings", return_value=fake_meetings), \
             patch.object(suggester, "_load_meeting_tdocs", return_value=fake_tdocs):

            inferred = suggester.infer_topic_from_known_ais({"SA2": ["20.6.21"]})
            kws = [k.lower() for k in inferred.get("keywords", [])]
            wis = inferred.get("wi_tokens", [])

            self.assertIn("6g data framework", kws)
            self.assertIn("data framework", kws)
            self.assertIn("data transfer", kws)
            self.assertIn("fs_6g_arc", wis)
            self.assertTrue("SA2 20.6.21" in inferred.get("source_summary", ""))

    def test_generic_ai_downgrade(self):
        from unittest.mock import patch
        from modules.meetings.core.ai_suggester import AgendaItemSuggester

        mock_db = unittest.mock.MagicMock()
        suggester = AgendaItemSuggester(mock_db)

        fake_meetings = [{"folder_name": "TSGR2_135", "meeting_number": "135", "wg_name": "RAN2"}]
        fake_tdocs = [
            # Generic AI (Corrections) with 4 matching titles
            {"Agenda Item": "7.2.2", "Agenda item description": "Corrections", "Title": "Data transfer correction 1", "Related WIs": ""},
            {"Agenda Item": "7.2.2", "Agenda item description": "Corrections", "Title": "Data transfer correction 2", "Related WIs": ""},
            {"Agenda Item": "7.2.2", "Agenda item description": "Corrections", "Title": "Data transfer correction 3", "Related WIs": ""},
            {"Agenda Item": "7.2.2", "Agenda item description": "Corrections", "Title": "Data transfer correction 4", "Related WIs": ""},
            # Real topic AI
            {"Agenda Item": "9.3.3.1", "Agenda item description": "Data transfer", "Title": "Discussion on Data transfer", "Related WIs": ""},
        ]

        with patch.object(suggester, "_resolve_recent_meetings", return_value=fake_meetings), \
             patch.object(suggester, "_load_meeting_tdocs", return_value=fake_tdocs):

            results = suggester.suggest_ais_for_topic(
                target_wgs=["RAN2"],
                keywords=["data transfer"],
            )

            r2_ais = {a.ai_number: a for a in results.get("RAN2", [])}
            # 9.3.3.1 is High confidence
            self.assertEqual(r2_ais["9.3.3.1"].confidence, "High")
            # 7.2.2 (Corrections) must NOT be High confidence despite 4 title matches
            self.assertNotEqual(r2_ais["7.2.2"].confidence, "High")
            self.assertIn("Caution: Generic/administrative", r2_ais["7.2.2"].matching_reason)



class TestConsideredInputsAndMeetingRounds(unittest.TestCase):
    def test_format_round_label(self):
        from modules.meetings.core.cohort_resolver import format_round_label, format_cohort_label
        self.assertEqual(format_round_label("2026-Q1"), "Feb 2026 Meeting Round")
        self.assertEqual(format_round_label("2026-Q2"), "May 2026 Meeting Round")
        self.assertEqual(format_round_label("2026-Q3"), "Aug 2026 Meeting Round")
        self.assertEqual(format_round_label("2026-Q4"), "Nov 2026 Meeting Round")
        self.assertEqual(format_cohort_label("2026-Q1"), "Feb 2026 Meeting Round")

    def test_date_range_cohort_resolution(self):
        from unittest.mock import MagicMock
        mock_db = MagicMock()
        mock_meetings = [
            {"wg_name": "SA2", "meeting_number": "160", "start_date": "2024-10-20", "end_date": "2024-10-24"},
            {"wg_name": "SA2", "meeting_number": "161", "start_date": "2025-02-17", "end_date": "2025-02-21"},
            {"wg_name": "SA2", "meeting_number": "162", "start_date": "2025-05-19", "end_date": "2025-05-23"},
            {"wg_name": "SA2", "meeting_number": "163", "start_date": "2025-08-18", "end_date": "2025-08-22"},
            {"wg_name": "SA2", "meeting_number": "164", "start_date": "2025-10-20", "end_date": "2025-10-24"},
            {"wg_name": "SA2", "meeting_number": "165", "start_date": "2026-02-16", "end_date": "2026-02-20"},
        ]
        mock_db.search_meetings.return_value = mock_meetings

        topic = MetaTopicDefinition(
            id="test_range",
            name="Test Range Topic",
            target_wgs=["SA2"],
            cycle_mode="date_range",
            date_from="2025-01-01",
            date_to="2025-09-30",
        )
        resolved = MeetingCohortResolver.resolve_cohorts_for_topic(mock_db, topic)
        keys = [c.key for c in resolved]
        self.assertEqual(keys, ["2025-Q1", "2025-Q2", "2025-Q3"])

    def test_topic_inputs_dialog_initialization(self):
        from PyQt5.QtWidgets import QApplication
        from unittest.mock import MagicMock
        from modules.meetings.ui.dialogs.topic_inputs_dialog import TopicInputsViewerDialog

        app = QApplication.instance() or QApplication([])

        mock_db = MagicMock()
        topic = MetaTopicDefinition(
            id="test_inputs",
            name="Data Framework & Data Transfer",
            target_wgs=["SA2", "RAN2"],
        )
        round_1 = MeetingCohort(
            key="2026-Q1",
            year=2026,
            quarter=1,
            label="Feb 2026 Meeting Round",
            meetings=[{"wg_name": "SA2", "meeting_number": "165", "start_date": "2026-02-16"}],
            wgs_present={"SA2"},
        )
        dlg = TopicInputsViewerDialog(
            meetings_db=mock_db,
            topic=topic,
            selected_rounds=[round_1],
            scope_mode="full_dossier",
            timeframe_label="Feb 2026 Meeting Round",
        )
        self.assertIsNotNone(dlg)
        self.assertEqual(dlg.topic.name, "Data Framework & Data Transfer")
        self.assertEqual(len(dlg.rounds), 1)
        self.assertEqual(dlg.tabs.count(), 4)
        dlg.close()


class TestTopicPillWidgets(unittest.TestCase):
    def test_wg_toggle_pill(self):
        from PyQt5.QtWidgets import QApplication
        from modules.meetings.ui.topic_pills import WgTogglePill

        app = QApplication.instance() or QApplication([])
        pill = WgTogglePill("SA2", checked=False)
        self.assertFalse(pill.isChecked())
        self.assertIn("+ SA2", pill.text())

        pill.setChecked(True)
        self.assertTrue(pill.isChecked())
        self.assertIn("✓ SA2", pill.text())

    def test_tag_cloud_widget(self):
        from PyQt5.QtWidgets import QApplication
        from modules.meetings.ui.topic_pills import TagCloudWidget

        app = QApplication.instance() or QApplication([])
        cloud = TagCloudWidget()
        cloud.setText("ambientiot, ambientiot_ph2, fs_ambientiot")
        self.assertEqual(cloud.tags(), ["ambientiot", "ambientiot_ph2", "fs_ambientiot"])
        self.assertEqual(cloud.text(), "ambientiot, ambientiot_ph2, fs_ambientiot")

        cloud.add_tag("new_tag")
        self.assertIn("new_tag", cloud.tags())

        cloud.remove_tag("ambientiot_ph2")
        self.assertNotIn("ambientiot_ph2", cloud.tags())

    def test_agenda_item_pills_widget(self):
        from PyQt5.QtWidgets import QApplication
        from modules.meetings.ui.topic_pills import AgendaItemPillsWidget

        app = QApplication.instance() or QApplication([])
        ai_widget = AgendaItemPillsWidget()
        ai_widget.setText("SA2: 20.6.21*, 5.14*; RAN2: 9.3.3.1*")
        pats = ai_widget.patterns_map()
        self.assertEqual(pats["SA2"], ["20.6.21*", "5.14*"])
        self.assertEqual(pats["RAN2"], ["9.3.3.1*"])

        ai_widget.add_pattern("CT4", "21.5*")
        self.assertIn("21.5*", ai_widget.patterns_map()["CT4"])

        ai_widget.remove_pattern("SA2", "5.14*")
        self.assertEqual(ai_widget.patterns_map()["SA2"], ["20.6.21*"])

    def test_topic_editor_dialog_pill_integration(self):
        from PyQt5.QtWidgets import QApplication
        from modules.meetings.ui.dialogs.topic_dossier_dialog import TopicEditorDialog

        app = QApplication.instance() or QApplication([])
        topic = MetaTopicDefinition(
            id="test_ambient",
            name="Ambient IoT",
            target_wgs=["SA2", "RAN2"],
            rules=MetaTopicRules(
                wi_tokens=["ambientiot"],
                ai_patterns={"SA2": ["5.14*"], "RAN2": ["8.7*"]},
                keywords=["ambient iot"],
            ),
        )
        dlg = TopicEditorDialog(topic=topic)
        self.assertTrue(dlg.wg_checks["SA2"].isChecked())
        self.assertTrue(dlg.wg_checks["RAN2"].isChecked())
        self.assertFalse(dlg.wg_checks["CT4"].isChecked())

        self.assertEqual(dlg.edit_wi.text(), "ambientiot")
        self.assertIn("5.14*", dlg.edit_ai.text())
        self.assertIn("8.7*", dlg.edit_ai.text())
        self.assertEqual(dlg.edit_kw.text(), "ambient iot")
        dlg.close()

    def test_filter_pill_button(self):
        from PyQt5.QtWidgets import QApplication
        from modules.meetings.ui.topic_pills import FilterPillButton

        app = QApplication.instance() or QApplication([])
        pill = FilterPillButton("RAN2", "RAN2", count=15)
        self.assertIn("RAN2 (15)", pill.text())
        pill.set_count(25)
        self.assertIn("RAN2 (25)", pill.text())
        pill.setChecked(True)
        self.assertTrue(pill.isChecked())

    def test_topic_inputs_viewer_pills_filtering(self):
        from PyQt5.QtWidgets import QApplication
        from unittest.mock import MagicMock
        from modules.meetings.ui.dialogs.topic_inputs_dialog import TopicInputsViewerDialog

        app = QApplication.instance() or QApplication([])
        mock_db = MagicMock()
        topic = MetaTopicDefinition(
            id="test_filter_pills",
            name="Data Framework & Data Transfer",
            target_wgs=["SA2", "RAN2"],
        )
        round_1 = MeetingCohort(
            key="2026-Q1",
            year=2026,
            quarter=1,
            label="Feb 2026 Meeting Round",
            meetings=[{"wg_name": "SA2", "meeting_number": "165", "start_date": "2026-02-16"}],
            wgs_present={"SA2"},
        )
        dlg = TopicInputsViewerDialog(
            meetings_db=mock_db,
            topic=topic,
            selected_rounds=[round_1],
            scope_mode="full_dossier",
        )
        if dlg._collector_worker:
            dlg._collector_worker.cancel()
            dlg._collector_worker.wait(1000)

        # Mock sample data
        dlg.all_contributions = [
            {"tdoc_id": "S2-2601001", "wg": "SA2", "type": "Discussion", "round_label": "Feb 2026 Meeting Round", "title": "Doc 1"},
            {"tdoc_id": "R2-2601002", "wg": "RAN2", "type": "Discussion", "round_label": "Feb 2026 Meeting Round", "title": "Doc 2"},
        ]
        dlg.all_decisions = [
            TopicDecision(cohort_key="2026-Q1", wg="SA2", meeting_number="165", date="2026-02-16", decision_type="Approved pCR", title_or_summary="pCR 1", tdoc_id="S2-2601001")
        ]
        dlg._update_kpi_strip()
        self.assertEqual(dlg.wg_pills["SA2"].text(), "SA2 (1)")
        self.assertEqual(dlg.wg_pills["RAN2"].text(), "RAN2 (1)")

        # Filter by RAN2
        dlg.wg_pills["RAN2"].setChecked(True)
        self.assertEqual(dlg.table_contributions.rowCount(), 1)
        self.assertEqual(dlg.table_contributions.item(0, 0).text(), "R2-2601002")

        # Reset
        dlg._reset_filters()
        self.assertEqual(dlg.table_contributions.rowCount(), 2)
        dlg.close()

    def test_standalone_agreement_markdown_export(self):
        import tempfile
        from modules.meetings.core.exporters.topic_dossier_exporter import TopicDossierExporterThread
        from unittest.mock import MagicMock

        with tempfile.TemporaryDirectory() as temp_dir:
            mock_db = MagicMock()
            topic = MetaTopicDefinition(
                id="test_standalone_agreements",
                name="Data Framework & Data Transfer",
                target_wgs=["RAN2"],
            )
            exporter = TopicDossierExporterThread(
                meetings_db=mock_db,
                topic=topic,
                selected_cohorts=[],
                output_base_dir=Path(temp_dir),
                target_folder_name="AgreementExportTest",
            )

            target_dir = Path(temp_dir) / "AgreementExportTest"
            target_dir.mkdir(parents=True, exist_ok=True)

            candidate_docs = [
                {
                    "tdoc_id": "R2-2600942",
                    "rel_path": "RAN2/RAN2_133/Discussion/R2-2600942.md",
                    "abs_path": target_dir / "RAN2/RAN2_133/Discussion/R2-2600942.md",
                }
            ]

            dec = TopicDecision(
                cohort_key="2026-Q1",
                wg="RAN2",
                meeting_number="133",
                date="2026-02-09",
                decision_type="Text Agreement",
                title_or_summary="Agreements\n1 For data transfer, study both long-cycle and bursty traffic.",
                tdoc_id="R2-2600942",
                ai="10.3.2.1",
                meeting_folder_id="RAN2_133",
            )

            rel_path = exporter._write_agreement_markdown(target_dir, dec, candidate_docs, seq=1)
            expected_file = target_dir / "RAN2" / "RAN2_133" / "Agreements" / "Agreement_R2-2600942.md"
            self.assertTrue(expected_file.is_file())
            self.assertEqual(rel_path, "RAN2/RAN2_133/Agreements/Agreement_R2-2600942.md")

            content = expected_file.read_text(encoding="utf-8")
            self.assertIn("# 📜 3GPP Text Agreement: RAN2 #133 (AI 10.3.2.1)", content)
            self.assertIn("1 For data transfer, study both long-cycle and bursty traffic.", content)
            self.assertIn("[R2-2600942]", content)
            self.assertIn("Discussion/R2-2600942.md", content)

    def test_decision_provenance_badges_and_tooltips(self):
        from PyQt5.QtWidgets import QApplication
        from unittest.mock import MagicMock
        from modules.meetings.ui.dialogs.topic_inputs_dialog import TopicInputsViewerDialog

        app = QApplication.instance() or QApplication([])
        mock_db = MagicMock()
        topic = MetaTopicDefinition(id="t_prov", name="Topic Provenance Test", target_wgs=["RAN2", "SA2"])
        round_1 = MeetingCohort(
            key="2026-Q1", year=2026, quarter=1, label="Feb 2026 Meeting Round",
            meetings=[{"wg_name": "RAN2", "meeting_number": "133", "start_date": "2026-02-09"}],
            wgs_present={"RAN2"},
        )
        dlg = TopicInputsViewerDialog(
            meetings_db=mock_db,
            topic=topic,
            selected_rounds=[round_1],
        )
        if dlg._collector_worker:
            dlg._collector_worker.cancel()
            dlg._collector_worker.wait(1000)

        dec_text = TopicDecision(
            cohort_key="2026-Q1", wg="RAN2", meeting_number="133", date="2026-02-09",
            decision_type="Text Agreement", title_or_summary="Agreement on sensing",
            tdoc_id="R2-2602804", ai="9.3.3.1",
        )
        dec_pcr = TopicDecision(
            cohort_key="2026-Q1", wg="SA2", meeting_number="173", date="2026-02-09",
            decision_type="Approved pCR", title_or_summary="Conclusions for KI#2",
            tdoc_id="S2-2601558", ai="20.5.1", target_spec="23.700-30",
        )

        dlg._populate_decisions_table([dec_text, dec_pcr])
        self.assertEqual(dlg.table_decisions.rowCount(), 2)

        # Row 0: Text Agreement
        item_text = dlg.table_decisions.item(0, 0)
        self.assertIn("Text Agreement (Notes)", item_text.text())
        self.assertIn("Binding Consensus Text", item_text.toolTip())
        self.assertIn("Chairman's Notes", item_text.toolTip())

        # Row 1: Approved pCR
        item_pcr = dlg.table_decisions.item(1, 0)
        self.assertIn("Approved pCR (Doc)", item_pcr.text())
        self.assertIn("Approved Specification Change", item_pcr.toolTip())
        self.assertIn("23.700-30", item_pcr.toolTip())

        dlg.close()

    def test_agreement_detail_reader_drawer(self):
        from PyQt5.QtWidgets import QApplication
        from unittest.mock import MagicMock
        from modules.meetings.ui.dialogs.topic_inputs_dialog import TopicInputsViewerDialog

        app = QApplication.instance() or QApplication([])
        mock_db = MagicMock()
        topic = MetaTopicDefinition(id="t_drawer", name="Drawer Test", target_wgs=["RAN2"])
        round_1 = MeetingCohort(
            key="2026-Q1", year=2026, quarter=1, label="Feb 2026 Meeting Round",
            meetings=[{"wg_name": "RAN2", "meeting_number": "133", "start_date": "2026-02-09"}],
            wgs_present={"RAN2"},
        )
        dlg = TopicInputsViewerDialog(
            meetings_db=mock_db,
            topic=topic,
            selected_rounds=[round_1],
        )
        if dlg._collector_worker:
            dlg._collector_worker.cancel()
            dlg._collector_worker.wait(1000)

        dec = TopicDecision(
            cohort_key="2026-Q1", wg="RAN2", meeting_number="133", date="2026-02-09",
            decision_type="Text Agreement",
            title_or_summary="Agreements\n1 Study long-cycle and bursty traffic.",
            tdoc_id="R2-2600942", ai="10.3.2.1",
        )

        dlg._populate_decisions_table([dec])

        # Verify drawer populated
        self.assertIn("Text Agreement (Chairman's Notes)", dlg.lbl_drawer_provenance.text())
        html = dlg.detail_text_edit.toHtml()
        self.assertIn("Study long-cycle and bursty traffic.", html)
        self.assertIn("RAN2", html)
        self.assertIn("R2-2600942", html)

        # Verify toggle collapse/expand
        self.assertFalse(dlg.detail_text_edit.isHidden())
        self.assertEqual(dlg.btn_toggle_drawer.text(), "▲ Collapse")
        dlg._toggle_drawer()
        self.assertTrue(dlg.detail_text_edit.isHidden())
        self.assertEqual(dlg.btn_toggle_drawer.text(), "▼ Expand")
        dlg._toggle_drawer()
        self.assertFalse(dlg.detail_text_edit.isHidden())
        self.assertEqual(dlg.btn_toggle_drawer.text(), "▲ Collapse")

        # Verify copy text
        dlg._on_copy_detail_text()
        self.assertEqual(app.clipboard().text(), dec.title_or_summary)

        # Verify contribution selection
        contrib = {
            "tdoc_id": "R2-2601999", "wg": "RAN2", "meeting_number": "133",
            "type": "Discussion", "status": "Noted", "title": "Discussion Paper on 6G",
            "source": "T-Mobile", "ai": "10.3.2.1",
            "raw_row": {"Abstract": "Proposal for data transfer."},
        }
        dlg._populate_contributions_table([contrib])
        dlg.tabs.setCurrentIndex(1)
        dlg.table_contributions.selectRow(0)
        c_html = dlg.detail_text_edit.toHtml()
        self.assertIn("Discussion Paper on 6G", c_html)
        self.assertIn("Proposal for data transfer.", c_html)

        dlg.close()

    def test_category_pill_auto_tab_switching_and_fast_filter(self):
        from PyQt5.QtWidgets import QApplication
        from unittest.mock import MagicMock
        from modules.meetings.ui.dialogs.topic_inputs_dialog import TopicInputsViewerDialog

        app = QApplication.instance() or QApplication([])
        mock_db = MagicMock()
        topic = MetaTopicDefinition(id="t_tab_switch", name="Tab Switch Test", target_wgs=["RAN2"])
        round_1 = MeetingCohort(
            key="2026-Q1", year=2026, quarter=1, label="Feb 2026 Meeting Round",
            meetings=[{"wg_name": "RAN2", "meeting_number": "133", "start_date": "2026-02-09"}],
            wgs_present={"RAN2"},
        )
        dlg = TopicInputsViewerDialog(
            meetings_db=mock_db,
            topic=topic,
            selected_rounds=[round_1],
        )
        if dlg._collector_worker:
            dlg._collector_worker.cancel()
            dlg._collector_worker.wait(1000)

        dec = TopicDecision(
            cohort_key="2026-Q1", wg="RAN2", meeting_number="133", date="2026-02-09",
            decision_type="Text Agreement",
            title_or_summary="Agreements\n1 Study long-cycle traffic.",
            tdoc_id="R2-2600942", ai="10.3.2.1",
        )
        dlg.all_decisions = [dec]
        dlg.all_contributions = [
            {"tdoc_id": "R2-2600942", "wg": "RAN2", "meeting_number": "133", "type": "Discussion", "status": "Noted", "title": "Doc", "ai": "10.3.2.1", "is_decision": True, "decision_type": "Text Agreement"},
            {"tdoc_id": "R2-2600943", "wg": "RAN2", "meeting_number": "133", "type": "LS out", "status": "Approved", "title": "LS", "ai": "10.3.2.1", "is_decision": False},
        ]
        dlg._update_kpi_strip()

        # Switch to tab 1
        dlg.tabs.setCurrentIndex(1)
        self.assertEqual(dlg.tabs.currentIndex(), 1)

        # Toggle Decisions pill -> auto-switches to Decisions tab 0
        dlg.cat_pills["🟢 Decisions Only (Agreements, pCRs, CRs)"].setChecked(True)
        self.assertEqual(dlg.tabs.currentIndex(), 0)
        self.assertEqual(dlg.table_decisions.rowCount(), 1)

        # Toggle Liaisons pill -> auto-switches to Liaisons tab 2
        dlg.cat_pills["📬 Liaison Statements Only"].setChecked(True)
        self.assertEqual(dlg.tabs.currentIndex(), 2)

        # Toggle Discussions pill -> auto-switches to Contributions tab 1
        dlg.cat_pills["📄 Discussion Papers Only"].setChecked(True)
        self.assertEqual(dlg.tabs.currentIndex(), 1)

        dlg.close()

    def test_always_download_missing_documents_from_ftp_defaults(self):
        """
        Verifies that:
        1. TopicDossierExporterThread has allow_network_download=True by default.
        2. TopicDossierDialog has chk_download checked (True) by default.
        3. InputsCollectorThread creates exporter with allow_network_download=True.
        """
        from unittest.mock import MagicMock, patch
        from PyQt5.QtWidgets import QApplication
        from modules.meetings.core.exporters.topic_dossier_exporter import TopicDossierExporterThread
        from modules.meetings.ui.dialogs.topic_dossier_dialog import TopicDossierDialog
        from modules.meetings.ui.dialogs.topic_inputs_dialog import InputsCollectorThread
        from modules.meetings.core.topic_manager import MetaTopicDefinition
        from modules.meetings.core.cohort_resolver import MeetingCohort

        mock_db = MagicMock()
        topic = MetaTopicDefinition(id="test_dl", name="Test DL", target_wgs=["RAN2"])

        # 1. ExporterThread default
        exp = TopicDossierExporterThread(
            meetings_db=mock_db,
            topic=topic,
            selected_cohorts=[],
        )
        self.assertTrue(exp.allow_network_download, "Exporter should default allow_network_download to True")

        # 2. TopicDossierDialog default
        app = QApplication.instance() or QApplication([])
        dlg = TopicDossierDialog(meetings_db=mock_db, parent=None)
        self.assertTrue(dlg.chk_download.isChecked(), "TopicDossierDialog checkbox should be checked by default")
        dlg.close()

        # 3. InputsCollectorThread helper thread
        round_1 = MeetingCohort(
            year=2026, quarter=1,
            key="2026-Q1", label="Feb 2026",
            meetings=[{"wg_name": "RAN2", "meeting_number": "133", "folder_name": "TSGR2_133"}],
            wgs_present={"RAN2"},
        )
        col = InputsCollectorThread(
            meetings_db=mock_db,
            topic=topic,
            rounds=[round_1],
        )
        # Verify run() sets allow_network_download=True when instantiating helper_thread
        with patch("modules.meetings.core.exporters.topic_dossier_exporter.TopicDossierExporterThread") as mock_exp_cls:
            col.run()
            mock_exp_cls.assert_called_once()
            _, kwargs = mock_exp_cls.call_args
            self.assertTrue(kwargs.get("allow_network_download", False))


class TestTopicDossierRefactorAndDigests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp())
        self.cache_dir = self.temp_dir / "cache"
        self.export_dir = self.temp_dir / "export"
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.export_dir.mkdir(parents=True, exist_ok=True)

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_resolve_meeting_folder_id_disambiguates_bis_and_ahe(self):
        from modules.meetings.core.exporters.topic_digest_builder import resolve_meeting_folder_id
        self.assertEqual(
            resolve_meeting_folder_id({"wg_name": "RAN3", "meeting_number": "133", "folder_name": "TSGR3_133"}),
            "RAN3_133",
        )
        self.assertEqual(
            resolve_meeting_folder_id({"wg_name": "RAN3", "meeting_number": "133", "folder_name": "TSGR3_133-bis"}),
            "RAN3_133-bis",
        )
        self.assertEqual(
            resolve_meeting_folder_id({"wg_name": "RAN2", "meeting_number": "133BIS", "folder_name": "TSGR2_133bis"}),
            "RAN2_133BIS",
        )
        self.assertEqual(
            resolve_meeting_folder_id({"wg_name": "SA2", "meeting_number": "175", "folder_name": "TSGS2_175_AHE"}),
            "SA2_175-AHE",
        )

    def test_canonical_cache_reuse_skips_download_and_rewrites_media_links(self):
        from unittest.mock import MagicMock, patch
        from modules.meetings.core.exporters.topic_dossier_exporter import TopicDossierExporterThread
        from modules.meetings.core.exporters.export_manifest import is_export_complete
        from modules.meetings.core.tdocs.tdoc_markdown_converter import CONVERTER_VERSION

        # Seed canonical cache for TSGR2_133/R2-2600942/R2-2600942.md + media/fig1.png
        mtg_cache = self.cache_dir / "TSGR2_133"
        tdoc_cache = mtg_cache / "R2-2600942"
        (tdoc_cache / "media").mkdir(parents=True, exist_ok=True)
        (tdoc_cache / "media" / "fig1.png").write_bytes(b"\x89PNG\r\n\x1a\nFAKEPNG")
        (tdoc_cache / "R2-2600942.md").write_text(
            f"<!--- tdoc_converter: version={CONVERTER_VERSION}; mode=markup --->\n"
            "# R2-2600942: Framework Discussion\n"
            "- **Source:** Nokia\n\n---\n\n"
            "## 1. Introduction\nWe propose E2E framework.\n\n"
            "![Arch](media/fig1.png)\n",
            encoding="utf-8",
        )

        topic = MetaTopicDefinition(
            id="t_cache",
            name="Data Framework",
            target_wgs=["RAN2"],
            rules=MetaTopicRules(wi_tokens=["fs_6g_df"]),
        )
        cohort = MeetingCohort(
            key="2026-Q1", year=2026, quarter=1, label="Feb 2026 Meeting Round",
            meetings=[{"wg_name": "RAN2", "meeting_number": "133", "start_date": "2026-02-09", "folder_name": "TSGR2_133"}],
            wgs_present={"RAN2"},
        )
        exporter = TopicDossierExporterThread(
            meetings_db=MagicMock(),
            topic=topic,
            selected_cohorts=[cohort],
            output_base_dir=self.export_dir,
            target_folder_name="TopicCacheTest",
            allow_network_download=True,
        )

        with patch.object(exporter, "_get_meeting_dir", return_value=mtg_cache), \
             patch.object(exporter, "_load_meeting_tdocs", return_value=[
                 {"TDoc": "R2-2600942", "Title": "Framework Discussion", "AI": "10.3.2", "Related WIs": "FS_6G_DF", "Type": "Discussion", "Status": "Noted", "Source": "Nokia"}
             ]), \
             patch.object(exporter, "_load_meeting_agreements", return_value=[]), \
             patch.object(exporter, "_find_or_download_tdoc_file") as mock_dl:

            exporter._run_export()
            mock_dl.assert_not_called()

        dest_md = self.export_dir / "TopicCacheTest" / "RAN2" / "RAN2_133" / "Discussion" / "R2-2600942.md"
        dest_img = self.export_dir / "TopicCacheTest" / "RAN2" / "RAN2_133" / "Discussion" / "R2-2600942_media" / "fig1.png"
        self.assertTrue(dest_md.is_file())
        self.assertTrue(dest_img.is_file())
        content = dest_md.read_text(encoding="utf-8")
        self.assertIn("![Arch](R2-2600942_media/fig1.png)", content)
        self.assertNotIn("](media/fig1.png)", content)
        self.assertTrue(is_export_complete(dest_md))

    def test_conversion_writes_back_to_canonical_cache(self):
        from unittest.mock import MagicMock, patch
        from types import SimpleNamespace
        from modules.meetings.core.exporters.topic_dossier_exporter import TopicDossierExporterThread
        from modules.meetings.core.tdocs.tdoc_markdown_converter import CONVERTER_VERSION

        mtg_cache = self.cache_dir / "TSGR2_133"
        src_folder = mtg_cache / "R2-2600999"
        src_folder.mkdir(parents=True, exist_ok=True)
        fake_docx = src_folder / "R2-2600999r01.docx"
        fake_docx.write_bytes(b"PK\x03\x04FAKEDOCX")

        topic = MetaTopicDefinition(
            id="t_wb", name="Writeback Test", target_wgs=["RAN2"], rules=MetaTopicRules(wi_tokens=["fs_6g_df"])
        )
        cohort = MeetingCohort(
            key="2026-Q1", year=2026, quarter=1, label="Feb 2026 Meeting Round",
            meetings=[{
                "wg_name": "RAN2", "meeting_number": "133", "start_date": "2026-02-09",
                "folder_name": "TSGR2_133", "docs_folder_url": "https://www.3gpp.org/ftp/tsg_ran/WG2_RL2/TSGR2_133/Docs",
            }],
            wgs_present={"RAN2"},
        )
        exporter = TopicDossierExporterThread(
            meetings_db=MagicMock(),
            topic=topic,
            selected_cohorts=[cohort],
            output_base_dir=self.export_dir,
            target_folder_name="WritebackTest",
        )

        def fake_convert(**kwargs):
            out_dir = Path(kwargs["out_dir"])
            m_dir = Path(kwargs["media_dir"])
            m_dir.mkdir(parents=True, exist_ok=True)
            (m_dir / "chart.png").write_bytes(b"\x89PNG\r\n\x1a\n")
            md_p = out_dir / f"{kwargs['tdoc_id']}.md"
            md_p.write_text(
                f"<!--- tdoc_converter: version={CONVERTER_VERSION}; mode=markup --->\n"
                f"# {kwargs['tdoc_id']}: Proposal\n- **3GPP Server URL:** [{kwargs['tdoc_url']}]({kwargs['tdoc_url']})\n\n---\n\n"
                f"## Proposal 1\nUse option B.\n![C]({kwargs['media_rel']}/chart.png)\n",
                encoding="utf-8",
            )
            return SimpleNamespace(md_path=md_p)

        with patch.object(exporter, "_get_meeting_dir", return_value=mtg_cache), \
             patch.object(exporter, "_load_meeting_tdocs", return_value=[
                 {"TDoc": "R2-2600999r01", "Title": "Proposal", "AI": "10.3.2", "Related WIs": "FS_6G_DF", "Type": "Discussion", "Status": "Noted", "Source": "Ericsson"}
             ]), \
             patch.object(exporter, "_load_meeting_agreements", return_value=[]), \
             patch("modules.meetings.core.exporters.topic_dossier_exporter.convert_document_to_markdown", side_effect=fake_convert):

            exporter._run_export()

        canon_md = mtg_cache / "R2-2600999" / "R2-2600999r01.md"
        topic_md = self.export_dir / "WritebackTest" / "RAN2" / "RAN2_133" / "Discussion" / "R2-2600999r01.md"
        self.assertTrue(canon_md.is_file(), "Must write back converted .md to canonical base_tdoc cache folder")
        self.assertIn("media/chart.png", canon_md.read_text(encoding="utf-8"))
        self.assertIn("/INBOX/Revisions/R2-2600999r01.zip", canon_md.read_text(encoding="utf-8"))
        self.assertTrue(topic_md.is_file())
        self.assertIn("R2-2600999r01_media/chart.png", topic_md.read_text(encoding="utf-8"))

    def test_sa2_chair_notes_excluded_and_revision_dedup_and_digests_and_checkpoints(self):
        from unittest.mock import MagicMock, patch
        from modules.meetings.core.exporters.topic_dossier_exporter import TopicDossierExporterThread
        from modules.meetings.core.exporters.export_manifest import load_manifest

        # Create a real SA2 chair_notes_cache.json containing chair triage instructions (must be ignored!)
        sa2_cache = self.cache_dir / "TSGS2_173"
        (sa2_cache / "Agenda").mkdir(parents=True, exist_ok=True)
        (sa2_cache / "Agenda" / "chair_notes_cache.json").write_text(
            json.dumps({
                "version": 2,
                "agreements": [
                    {"text": "Merge into S2-2609911", "ai": "20.6.21", "decision_type": "Agreement", "wi": "FS_6G_DF"}
                ],
            }),
            encoding="utf-8",
        )

        topic = MetaTopicDefinition(
            id="t_full", name="Data Framework", target_wgs=["SA2", "RAN3"], rules=MetaTopicRules(wi_tokens=["fs_6g_df"])
        )
        cohort = MeetingCohort(
            key="2026-Q3", year=2026, quarter=3, label="Aug 2026 Meeting Round",
            meetings=[
                {"wg_name": "SA2", "meeting_number": "173", "start_date": "2026-08-17", "folder_name": "TSGS2_173"},
                {"wg_name": "RAN3", "meeting_number": "133", "start_date": "2026-08-20", "folder_name": "TSGR3_133"},
                {"wg_name": "RAN3", "meeting_number": "133", "start_date": "2026-10-12", "folder_name": "TSGR3_133-bis"},
            ],
            wgs_present={"SA2", "RAN3"},
        )

        # Exporter with dedup=True and include_discussions=False
        exporter = TopicDossierExporterThread(
            meetings_db=MagicMock(),
            topic=topic,
            selected_cohorts=[cohort],
            output_base_dir=self.export_dir,
            target_folder_name="FullPipelineTest",
            include_discussions=False,
            deduplicate_revisions=True,
            allow_network_download=False,
        )

        # Verify _load_meeting_agreements skips SA2 chair notes
        with patch.object(exporter, "_get_meeting_dir", return_value=sa2_cache):
            self.assertEqual(exporter._load_meeting_agreements({"wg_name": "SA2", "folder_name": "TSGS2_173"}), [])

        def fake_tdocs(mtg):
            fn = mtg.get("folder_name")
            if fn == "TSGS2_173":
                return [
                    {"TDoc": "S2-2601000", "Title": "pCR v0", "AI": "20.6.21", "Related WIs": "FS_6G_DF", "Type": "pCR", "Status": "Revised", "Revised to": "S2-2601000r01", "Source": "Nokia, Ericsson"},
                    {"TDoc": "S2-2601000r01", "Title": "pCR v1", "AI": "20.6.21", "Related WIs": "FS_6G_DF", "Type": "pCR", "Status": "Agreed", "Is revision of": "S2-2601000", "Source": "Nokia, [Ericsson?]"},
                    {"TDoc": "S2-2601002", "Title": "Random discussion", "AI": "20.6.21", "Related WIs": "FS_6G_DF", "Type": "Discussion", "Status": "Noted", "Source": "Huawei"},
                ]
            if fn == "TSGR3_133":
                return [
                    {"TDoc": "R3-2602001", "Title": "LS to SA2", "AI": "10.2.3", "Related WIs": "FS_6G_DF", "Type": "LS out", "Status": "Approved", "Source": "RAN3"},
                ]
            return [
                {"TDoc": "R3-2603001", "Title": "Bis Agreed WF", "AI": "10.2.3", "Related WIs": "FS_6G_DF", "Type": "WF", "Status": "Agreed", "Source": "Qualcomm"},
            ]

        with patch.object(exporter, "_get_meeting_dir", return_value=None), \
             patch.object(exporter, "_load_meeting_tdocs", side_effect=fake_tdocs), \
             patch.object(exporter, "_load_meeting_agreements", side_effect=lambda m: [{"text": "RAN3 Agreement 1", "ai": "10.2.3", "wi": "FS_6G_DF"}] if m["wg_name"] == "RAN3" and m["folder_name"] == "TSGR3_133" else []):

            exporter._run_export()

        root = self.export_dir / "FullPipelineTest"
        # Disambiguated -bis directories exist without colliding
        self.assertTrue((root / "RAN3" / "RAN3_133" / "_meeting_brief.md").is_file())
        self.assertTrue((root / "RAN3" / "RAN3_133-bis" / "_meeting_brief.md").is_file())
        self.assertTrue((root / "SA2" / "SA2_173" / "_company_positions_input.md").is_file())
        self.assertTrue((root / "company_positions_e2e_prompt.md").is_file())

        # Revision dedup removed S2-2601000, discussion filter removed S2-2601002, kept S2-2601000r01
        self.assertFalse((root / "SA2" / "SA2_173" / "pCR" / "S2-2601000.md").exists())
        self.assertFalse((root / "SA2" / "SA2_173" / "Discussion" / "S2-2601002.md").exists())
        self.assertTrue((root / "SA2" / "SA2_173" / "pCR" / "S2-2601000r01.md").is_file())

        # Check CompanySanitizer carry-over attribution warning in digest
        sa2_digest = (root / "SA2" / "SA2_173" / "_company_positions_input.md").read_text(encoding="utf-8")
        self.assertIn("S2-2601000r01", sa2_digest)
        self.assertIn("Supersedes S2-2601000", sa2_digest)
        self.assertIn("carry-over co-sponsors", sa2_digest)

        # Check manifest v2 checkpoints
        m = load_manifest(root)
        self.assertIsNotNone(m)
        self.assertIn("SA2/SA2_173", m.exported_meetings)
        self.assertIn("RAN3/RAN3_133", m.exported_meetings)
        self.assertIn("RAN3/RAN3_133-bis", m.exported_meetings)
        self.assertGreaterEqual(len(m.exported_decisions), 3)

        # Session 2 with identical data should report 0 new decisions in delta_session_2.md
        exporter2 = TopicDossierExporterThread(
            meetings_db=MagicMock(),
            topic=topic,
            selected_cohorts=[cohort],
            output_base_dir=self.export_dir,
            target_folder_name="FullPipelineTest",
            include_discussions=False,
            deduplicate_revisions=True,
            allow_network_download=False,
        )
        with patch.object(exporter2, "_get_meeting_dir", return_value=None), \
             patch.object(exporter2, "_load_meeting_tdocs", side_effect=fake_tdocs), \
             patch.object(exporter2, "_load_meeting_agreements", side_effect=lambda m: [{"text": "RAN3 Agreement 1", "ai": "10.2.3", "wi": "FS_6G_DF"}] if m["wg_name"] == "RAN3" and m["folder_name"] == "TSGR3_133" else []):
            exporter2._run_export()

        delta2 = (root / "deltas" / "delta_session_2.md").read_text(encoding="utf-8")
        self.assertIn("No new decisions recorded in this session", delta2)

    def test_open_output_folder_button_state_and_action(self):
        from unittest.mock import MagicMock, patch
        from PyQt5.QtWidgets import QApplication
        from modules.meetings.ui.dialogs.topic_dossier_dialog import TopicDossierDialog

        app = QApplication.instance() or QApplication([])
        dlg = TopicDossierDialog(meetings_db=MagicMock(), parent=None)

        # Nonexistent folder -> disabled
        missing_dir = self.temp_dir / "nonexistent_topic_dir"
        dlg.edit_dest_folder.setText(str(missing_dir))
        self.assertFalse(dlg.btn_open_folder.isEnabled())

        # Existing folder (with or without prior export in current session) -> enabled & opens destination
        existing_dir = self.temp_dir / "existing_topic_dir"
        existing_dir.mkdir(parents=True, exist_ok=True)
        dlg.last_target_dir = str(self.temp_dir / "other_stale_dir")  # stale previous topic path
        dlg.edit_dest_folder.setText(str(existing_dir))
        self.assertTrue(dlg.btn_open_folder.isEnabled())

        with patch("modules.meetings.ui.dialogs.topic_dossier_dialog._open_folder") as mock_open:
            dlg._on_open_output()
            mock_open.assert_called_once()
            self.assertEqual(mock_open.call_args[0][0].resolve(), existing_dir.resolve())

        dlg.close()


if __name__ == "__main__":
    unittest.main()
