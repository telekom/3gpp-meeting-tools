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


if __name__ == "__main__":
    unittest.main()
