# --- File: tests/test_email_ai_summary_core.py ---
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

project_root = Path(__file__).resolve().parent.parent
src_dir = project_root / "src"
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from core.ai.llm_runner import ModelSpec, RunnerResult
from modules.assistant.core.tools.external_agent_tools import build_external_agent_tool_registry
from modules.emails.core.ai_summary import (
    EmailSummaryOrchestrator,
    EmailSummaryRecord,
    EmailSummaryStore,
    SummaryScope,
    parse_summary_output,
)
from modules.emails.core.email_summary_worker import clean_email_body_for_summary
from modules.emails.core.general_email_db import GeneralEmailDatabase
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.specifications.core.database import SpecsDatabase


class TestEmailAiSummaryCore(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp_path = Path(self._tmp.name)
        self.agenda_dir = self.tmp_path / "TSGS2_174_Malta_2026-04" / "Agenda"
        self.agenda_dir.mkdir(parents=True, exist_ok=True)
        self.store = EmailSummaryStore(self.agenda_dir / "email_summaries.db")

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_clean_email_body_preserves_start_end_blocks(self):
        raw = "<<START>>\n[S2-2602001r01] Ericsson supports Option A.\n<<END>>\n\nBest regards,\nAlice\n\nFrom: Bob\nSent: Monday\n> old quote"
        cleaned, _ = clean_email_body_for_summary(raw)
        self.assertIn("Ericsson supports Option A", cleaned)
        self.assertNotIn("old quote", cleaned)

    def test_parse_summary_output_is_fault_tolerant(self):
        raw_ok = "### 1. Topic\nURSP\n\n```summary-meta\n{\"consensus_state\": \"converging\"}\n```"
        md, _delta, meta = parse_summary_output(raw_ok)
        self.assertEqual(md, "### 1. Topic\nURSP")
        self.assertEqual(meta.get("consensus_state"), "converging")

        raw_bad = "### 🆕 What Changed Since Last Cutoff\n- Nokia objected\n\n### 1. Topic\nURSP\n\n```summary-meta\n{bad json\n```"
        md2, delta2, meta2 = parse_summary_output(raw_bad)
        self.assertIn("### 1. Topic", md2)
        self.assertIn("Nokia objected", delta2)
        self.assertEqual(meta2, {})

    def test_store_covered_ids_catches_backfilled_older_email_and_undo(self):
        m1 = {"id": "E1", "sender_name": "Alice", "company": "Ericsson", "date_sent": "2026-04-14 10:00:00", "subject": "S2-2602001", "body_text": "Proposal A"}
        m2 = {"id": "E2", "sender_name": "Bob", "company": "Nokia", "date_sent": "2026-04-14 12:00:00", "subject": "Re: S2-2602001", "body_text": "We support"}

        orch = EmailSummaryOrchestrator(self.store, self.agenda_dir.parent, wg="SA2", meeting_id="S2-174")
        spec = ModelSpec("codex", "gpt-6.2-astra", "low")

        with patch("modules.emails.core.ai_summary.orchestrator.run_llm_text") as mock_run:
            mock_run.return_value = RunnerResult(
                ok=True,
                text="### 1. Topic\nInitial summary v1",
                provider="codex",
                model="gpt-6.2-astra",
                reasoning_effort="low",
            )
            rec1, err1 = orch.summarize_leaf(SummaryScope("tdoc", "S2-2602001", "URSP"), [m1, m2], model_spec=spec)
            self.assertEqual(err1, "")
            self.assertEqual(rec1.cutoff_utc, "2026-04-14 12:00:00")
            self.assertEqual(set(rec1.covered_message_ids), {"E1", "E2"})

        # Backfill an OLDER email (08:00 < 12:00 cutoff!) via split-thread merge
        m0_backfill = {"id": "E0", "sender_name": "Carol", "company": "Huawei", "date_sent": "2026-04-14 08:00:00", "subject": "S2-2602001", "body_text": "Early objection"}
        fresh = self.store.compute_freshness("tdoc", "S2-2602001", [m0_backfill, m1, m2])
        self.assertEqual(fresh.new_count, 1)
        self.assertEqual(fresh.new_message_ids, ["E0"])

        # Incremental update runs and increments hop count
        with patch("modules.emails.core.ai_summary.orchestrator.run_llm_text") as mock_run:
            mock_run.return_value = RunnerResult(
                ok=True,
                text="### 🆕 What Changed Since Last Cutoff\n- Huawei raised early objection\n\n### 1. Topic\nMerged v2",
                provider="codex",
                model="gpt-6.2-astra",
                reasoning_effort="low",
            )
            rec2, err2 = orch.summarize_leaf(SummaryScope("tdoc", "S2-2602001", "URSP"), [m0_backfill, m1, m2], model_spec=spec)
            self.assertEqual(err2, "")
            self.assertEqual(rec2.incremental_hops, 1)
            self.assertIn("E0", rec2.covered_message_ids)
            self.assertIn("[Msg #1]", mock_run.call_args.args[1])  # backfilled message gets exact chronological index #1

        # Undo restores v1 cleanly
        restored = self.store.revert_previous_version("tdoc", "S2-2602001")
        self.assertIsNotNone(restored)
        self.assertIn("Initial summary v1", restored.summary_md)

    def test_per_ai_map_reduce_reuses_fresh_leaves_and_updates_stale_only(self):
        # Pre-seed an up-to-date leaf for TDoc A
        self.store.upsert_summary(
            EmailSummaryRecord(
                scope_type="tdoc",
                scope_key="S2-2602001",
                title="URSP A",
                summary_md="### 1. Topic\nAgreed on A",
                covered_message_ids=["A1"],
                message_count=1,
                cutoff_utc="2026-04-14 10:00:00",
                provider="ollama",
                model="qwen2.5:7b",
            )
        )
        buckets = {
            "TDOC::S2-2602001": {
                "scope_type": "tdoc",
                "scope_key": "S2-2602001",
                "title": "URSP A",
                "messages": [{"id": "A1", "company": "Ericsson", "date_sent": "2026-04-14 10:00:00", "body_text": "Agree"}],
            },
            "TDOC::S2-2602002": {
                "scope_type": "tdoc",
                "scope_key": "S2-2602002",
                "title": "URSP B",
                "messages": [{"id": "B1", "company": "Nokia", "date_sent": "2026-04-14 11:00:00", "body_text": "Need revision"}],
            },
        }
        orch = EmailSummaryOrchestrator(self.store, self.agenda_dir.parent, wg="SA2", meeting_id="S2-174")
        calls = []

        def fake_run(sys_p, usr_p, spec, **_kw):
            calls.append((spec.provider, spec.model, usr_p[:120]))
            if "STAGE 2" in usr_p or "LEAF SUMMARIES" in usr_p:
                return RunnerResult(ok=True, text="### 1. Agenda Item Overview\nCross-TDoc synthesis", provider=spec.provider, model=spec.model)
            return RunnerResult(ok=True, text="### 1. Topic\nSummary for B", provider=spec.provider, model=spec.model)

        with patch("modules.emails.core.ai_summary.orchestrator.run_llm_text", side_effect=fake_run):
            rec, err = orch.summarize_ai(
                "6.1",
                "6.1 — 5GS Policy",
                buckets,
                leaf_model_spec=ModelSpec("ollama", "qwen2.5:7b"),
                synthesis_model_spec=ModelSpec("codex", "gpt-6.2-sol", "high"),
            )
        self.assertEqual(err, "")
        self.assertIsNotNone(rec)
        # Only 2 LLM calls: 1 for missing leaf S2-2602002 (NOT S2-2602001!), 1 for Stage-2 AI synthesis!
        self.assertEqual(len(calls), 2)
        self.assertEqual(calls[0][0], "ollama")
        self.assertEqual(calls[1][0], "codex")
        self.assertEqual(set(rec.covered_message_ids), {"A1", "B1"})

    def test_mcp_email_summary_actions_and_read_only_guard(self):
        specs_db_path = self.tmp_path / "specs.db"
        SpecsDatabase(db_path=specs_db_path)
        mdb = MeetingsDatabase(db_path=specs_db_path)
        mdb.upsert_single_meeting({
            "wg_name": "SA2",
            "folder_name": "TSGS2_174_Malta_2026-04",
            "meeting_number": "174",
            "name": "S2-174",
            "location": "Malta",
            "start_date": "2026-04-13",
            "end_date": "2026-04-17",
            "url_key": "tsg_sa/WG2_Arch/TSGS2_174_Malta_2026-04",
        })
        tdb = TrackedThreadsDatabase(self.agenda_dir / "tracked_threads.db")
        tid = tdb.create_or_get_thread("URSP debate", "ursp debate", meeting_id="S2-174", wg="SA2", linked_tdocs=["S2-2602001"])
        tdb.upsert_messages(tid, [
            {"id": "M1", "subject": "URSP debate", "sender_name": "Alice", "company": "Ericsson", "date_sent": "2026-04-14 10:00:00", "body_text": "Msg 1"},
            {"id": "M2", "subject": "Re: URSP debate", "sender_name": "Bob", "company": "Nokia", "date_sent": "2026-04-14 11:00:00", "body_text": "Msg 2"},
        ])

        with patch("modules.assistant.core.tools.external_agent.email_summary_tools._get_meetings_cache_roots", return_value=[self.tmp_path]):
            reg_rw = build_external_agent_tool_registry(specs_db_path=specs_db_path, read_only=False)
            save_res = reg_rw.execute("work_items_and_emails", {
                "action": "save_email_summary",
                "wg": "SA2",
                "meeting_number": "174",
                "scope_type": "thread",
                "scope_key": tid,
                "summary_md": "### 1. Topic\nSaved via MCP",
                "covered_ids": "M1",
                "cutoff_utc": "2026-04-14 10:00:00",
            })
            self.assertTrue(save_res.ok)
            self.assertTrue(save_res.data["ok"])

            get_res = reg_rw.execute("work_items_and_emails", {
                "action": "get_email_summary",
                "wg": "SA2",
                "meeting_number": "174",
                "scope_type": "thread",
                "scope_key": tid,
                "include_delta": True,
            })
            self.assertTrue(get_res.ok)
            self.assertTrue(get_res.data["found"])
            self.assertEqual(len(get_res.data["delta_messages"]), 1)
            self.assertEqual(get_res.data["delta_messages"][0]["id"], "M2")

            list_res = reg_rw.execute("work_items_and_emails", {
                "action": "list_email_summaries",
                "wg": "SA2",
                "meeting_number": "174",
            })
            self.assertEqual(list_res.data["count"], 1)

            reg_ro = build_external_agent_tool_registry(specs_db_path=specs_db_path, read_only=True)
            blocked = reg_ro.execute("work_items_and_emails", {
                "action": "save_email_summary",
                "wg": "SA2",
                "meeting_number": "174",
                "scope_type": "thread",
                "scope_key": tid,
                "summary_md": "Blocked",
            })
            self.assertFalse(blocked.data["ok"])


if __name__ == "__main__":
    unittest.main()
