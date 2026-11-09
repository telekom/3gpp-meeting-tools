# --- File: tests/test_mcp_and_cli_tools.py ---
from __future__ import annotations

import io
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

# Ensure 'src' is in sys.path for test execution
project_root = Path(__file__).resolve().parent.parent
src_dir = project_root / "src"
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from cli_tools import main as cli_main
from mcp_server import MCPServer, _read_next_rpc_message, _write_rpc_message, serve_stdio
from modules.assistant.core.tools.external_agent_tools import (
    PROFILE_TOOLS,
    _resolve_db_paths,
    build_external_agent_tool_registry,
)
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.tdocs.tdocs_db import TDocsDatabase
from modules.nas.core.nas_db import NASDatabase
from modules.spec_search.core.spec_search_db import SpecSearchDatabase
from modules.specifications.core.database import SpecsDatabase
from modules.specifications.core.spec_history_db import SpecHistoryDatabase
from modules.work_items.core.wi_database import WorkItemsDatabase


class TestMcpAndCliTools(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp_path = Path(self._tmp.name)

        self.specs_db_path = self.tmp_path / "specs.db"
        self.spec_search_db_path = self.tmp_path / "spec_search.db"
        self.protocol_db_path = self.tmp_path / "protocol.db"
        self.spec_history_db_path = self.tmp_path / "spec_history.db"
        self.user_tdocs_db_path = self.tmp_path / "user_tdocs.db"
        self.session_plan_db_path = self.tmp_path / "session_plans.db"

        # 1. Seed SpecsDatabase (`specifications` + `files` tables)
        specs_db = SpecsDatabase(db_path=self.specs_db_path)
        with specs_db._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO series (name, url) VALUES (?, ?)", ("23_series", "https://www.3gpp.org/ftp/Specs/archive/23_series"))
            s23_id = cursor.lastrowid
            cursor.execute("INSERT INTO working_groups (name) VALUES (?)", ("SA2",))
            wg_id = cursor.lastrowid
            cursor.execute(
                """
                INSERT INTO specifications (series_id, number, url, title, type, initial_release, primary_group_id)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    s23_id,
                    "23.501",
                    "https://www.3gpp.org/ftp/Specs/archive/23_series/23.501",
                    "System architecture for the 5G System (5GS); Stage 2",
                    "TS",
                    "Rel-15",
                    wg_id,
                ),
            )
            sp_id = cursor.lastrowid
            cursor.execute(
                "INSERT INTO files (spec_id, filename, version, url, upload_date) VALUES (?, ?, ?, ?, ?)",
                (
                    sp_id,
                    "23501-h00.zip",
                    "17.0.0",
                    "https://www.3gpp.org/ftp/Specs/archive/23_series/23.501/23501-h00.zip",
                    "2021-03-25",
                ),
            )
            cursor.execute(
                "INSERT INTO files (spec_id, filename, version, url, upload_date) VALUES (?, ?, ?, ?, ?)",
                (
                    sp_id,
                    "23501-i00.zip",
                    "18.0.0",
                    "https://www.3gpp.org/ftp/Specs/archive/23_series/23.501/23501-i00.zip",
                    "2023-03-30",
                ),
            )

        # 2. Seed SpecSearchDatabase
        spec_search_db = SpecSearchDatabase(db_path=self.spec_search_db_path)
        spec_search_db.insert_parsed_spec(
            "23.501",
            "17.0.0",
            [
                {
                    "clause_number": "5.5.1",
                    "clause_title": "Registration procedure",
                    "content": "The UE provides a 5G-GUTI or SUCI during initial registration.",
                }
            ],
            release_date="2021-03-25",
        )
        spec_search_db.insert_parsed_spec(
            "23.501",
            "18.0.0",
            [
                {
                    "clause_number": "5.5.1",
                    "clause_title": "Registration procedure",
                    "content": (
                        "The UE provides a 5G-GUTI or SUCI during initial registration. "
                        "In Rel-18 the AMF may also request PEI verification for SNPN onboarding."
                    ),
                }
            ],
            release_date="2023-03-30",
        )

        # 3. Seed NASDatabase
        nas_db = NASDatabase(db_path=self.protocol_db_path)
        nas_db.insert_parsed_spec(
            spec_number="24.501",
            version="18.0.0",
            messages=[
                {
                    "clause": "8.2.6",
                    "message_name": "REGISTRATION REQUEST",
                    "table_caption": "Table 8.2.6.1.1: REGISTRATION REQUEST message content",
                    "description": "Sent by the UE to the AMF to register.",
                    "direction": "UE to Network",
                    "ies": [
                        {
                            "iei": "",
                            "information_element": "5GS registration type",
                            "type_reference": "9.11.3.7",
                            "presence": "M",
                            "format": "V",
                            "length": "1/2",
                        },
                        {
                            "iei": "77",
                            "information_element": "5GS mobile identity",
                            "type_reference": "9.11.3.4",
                            "presence": "M",
                            "format": "LV-E",
                            "length": "6-n",
                        },
                    ],
                }
            ],
            ie_defs=[
                {
                    "clause": "9.11.3.7",
                    "ie_name": "5GS registration type",
                    "raw_description": "Indicates the type of registration requested by the UE.",
                    "structure_table": "",
                }
            ],
            spec_type="NAS",
        )

        # 4. Seed MeetingsDatabase, SpecHistoryDatabase, WorkItemsDatabase, and TDocsDatabase
        meetings_db = MeetingsDatabase(db_path=self.specs_db_path)
        for mtg_dict in [
            {
                "wg_name": "SA2",
                "folder_name": "TSGS2_174_Malta_2026-04",
                "meeting_number": "174",
                "name": "S2-174",
                "location": "Malta",
                "start_date": "2026-04-13",
                "end_date": "2026-04-17",
                "url_key": "tsg_sa/WG2_Arch/TSGS2_174_Malta_2026-04",
                "docs_folder_url": "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_174_Malta_2026-04/Docs/",
                "first_tdoc": "S2-2601720",
                "last_tdoc": "S2-2603600",
                "mtg_id": "60765",
            },
            {
                "wg_name": "SA2",
                "folder_name": "TSGS2_175_Dalian_2026-05",
                "meeting_number": "175",
                "name": "S2-175",
                "location": "Dalian, CN",
                "start_date": "2026-05-18",
                "end_date": "2026-05-22",
                "url_key": "tsg_sa/WG2_Arch/TSGS2_175_Dalian_2026-05",
                "docs_folder_url": "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_175_Dalian_2026-05/Docs/",
                "first_tdoc": "S2-2604000",
                "last_tdoc": "S2-2605999",
                "mtg_id": "60766",
            },
        ]:
            meetings_db.upsert_single_meeting(mtg_dict)

        # Seed cached meeting spreadsheet JSON in a disposable cache dir
        self.agenda_dir = self.tmp_path / "TSGS2_174_Malta_2026-04" / "Agenda"
        self.agenda_dir.mkdir(parents=True, exist_ok=True)
        tdoc_rows = [
            {
                "TDoc": "S2-2602001",
                "Title": "Clarification on URSP rule evaluation",
                "Source": "Ericsson",
                "Type": "pCR",
                "For": "Approval",
                "Agenda Item": "6.1",
                "Agenda item description": "5GS Policy",
                "Related WIs": "TEI18",
                "Spec": "23.503",
                "Release": "Rel-18",
                "TDoc Status": "agreed",
                "Secretary Remarks": "Agreed",
            },
            {
                "TDoc": "S2-2602002",
                "Title": "Carry-over pCR with question mark co-signer",
                "Source": "Nokia, Ericsson?",
                "Type": "pCR",
                "For": "Approval",
                "Agenda Item": "6.1",
                "Agenda item description": "5GS Policy",
                "Related WIs": "TEI18",
                "Spec": "23.503",
                "Release": "Rel-18",
                "TDoc Status": "revised",
                "Secretary Remarks": "Revised to r01",
            },
            {
                "TDoc": "S2-2602003",
                "Title": "Carry-over pCR with bracketed co-signer",
                "Source": "Qualcomm, [Ericsson]",
                "Type": "pCR",
                "For": "Approval",
                "Agenda Item": "6.2",
                "Agenda item description": "5GS Core",
                "Related WIs": "TEI18",
                "Spec": "23.501",
                "Release": "Rel-18",
                "TDoc Status": "noted",
                "Secretary Remarks": "Noted",
            },
            {
                "TDoc": "S2-2602004",
                "Title": "Rapporteur CR pack for SNPN",
                "Source": "Ericsson [Rapporteur]",
                "Type": "CR",
                "For": "Approval",
                "Agenda Item": "6.3",
                "Agenda item description": "SNPN",
                "Related WIs": "SNPN_Ph2",
                "Spec": "23.501",
                "Release": "Rel-18",
                "TDoc Status": "approved",
                "Secretary Remarks": "Approved",
            },
        ]
        (self.agenda_dir / "TDoc_List_Meeting_SA2#174.xlsx.json").write_text(
            json.dumps(tdoc_rows),
            encoding="utf-8",
        )

        # Seed SpecHistoryDatabase
        hist_db = SpecHistoryDatabase(db_path=self.spec_history_db_path)
        hist_db.upsert_history_records(
            [
                {
                    "spec_number": "23.501",
                    "version": "18.0.0",
                    "row_order": 1,
                    "date": "2023-03",
                    "meeting": "SP#99",
                    "source_tdoc": "SP-230101",
                    "cr_number": "4100",
                    "revision": "1",
                    "category": "B",
                    "subject": "PEI verification for SNPN onboarding",
                    "affected_clauses": "5.5.1",
                    "pcr_tdocs": ["S2-2301234"],
                },
                {
                    "spec_number": "23.501",
                    "version": "18.0.0",
                    "row_order": 2,
                    "date": "2023-03",
                    "meeting": "SP#99",
                    "source_tdoc": "SP-230102",
                    "cr_number": "4101",
                    "revision": "2",
                    "category": "F",
                    "subject": "Clarification on URSP rule evaluation",
                    "affected_clauses": "5.15.1",
                    "pcr_tdocs": ["S2-2602001"],
                },
            ]
        )

        # Seed WorkItemsDatabase
        wi_db = WorkItemsDatabase(db_path=self.specs_db_path)
        wi_db.upsert_work_items(
            "SA2",
            [
                {
                    "code": "1000001",
                    "acronym": "FS_6G_ARC",
                    "name": "Study on 6G System Architecture",
                    "release": "Rel-20",
                },
                {
                    "code": "1000002",
                    "acronym": "SNPN_Ph2",
                    "name": "Enhanced support of Non-Public Networks Phase 2",
                    "release": "Rel-18",
                },
            ],
        )

        # Seed TDocsDatabase
        TDocsDatabase(db_path=self.user_tdocs_db_path).upsert(
            "S2-2602001",
            "🟢 Support",
            "Aligned with our URSP position",
        )

        self.db_kwargs = {
            "specs_db_path": self.specs_db_path,
            "spec_search_db_path": self.spec_search_db_path,
            "protocol_db_path": self.protocol_db_path,
            "spec_history_db_path": self.spec_history_db_path,
            "user_tdocs_db_path": self.user_tdocs_db_path,
            "session_plan_db_path": self.session_plan_db_path,
        }

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_tool_registry_profiles_and_token_budget(self):
        all_reg = build_external_agent_tool_registry(**self.db_kwargs, profile="all")
        all_tools = all_reg.mcp_tools()
        self.assertEqual(len(all_tools), 12)

        names = [t["name"] for t in all_tools]
        self.assertEqual(set(names), set(PROFILE_TOOLS["all"]))

        # Verify token budget: compact JSON serialization of all 12 MCP tool schemas
        # should stay well under ~2,500 tokens (~10,000 chars) so Claude/Codex system prompts stay light.
        compact_schema_json = json.dumps(all_tools, separators=(",", ":"))
        estimated_tokens = len(compact_schema_json) // 4
        self.assertLess(
            estimated_tokens,
            2400,
            f"Tool schema token budget exceeded: ~{estimated_tokens} tokens",
        )

        # Profile filtering
        for profile_name in ("knowledge", "onsite", "authoring"):
            prof_reg = build_external_agent_tool_registry(**self.db_kwargs, profile=profile_name)
            prof_names = [t["name"] for t in prof_reg.mcp_tools()]
            self.assertEqual(set(prof_names), set(PROFILE_TOOLS[profile_name]))

    def test_read_only_mode_blocks_mutating_actions(self):
        ro_reg = build_external_agent_tool_registry(**self.db_kwargs, profile="all", read_only=True)

        # Read action should succeed
        res_read = ro_reg.execute(
            "spec_catalogue_and_files",
            {"action": "search", "query": "23.501"},
        )
        self.assertTrue(res_read.ok)
        self.assertGreaterEqual(res_read.data["count"], 1)

        # Mutating action on a read tool should be blocked in read-only mode
        res_write = ro_reg.execute(
            "spec_catalogue_and_files",
            {"action": "download", "spec_number": "23.501", "version": "18.0.0"},
        )
        self.assertTrue(res_write.ok)
        self.assertFalse(res_write.data["ok"])
        self.assertIn("Read-only mode", res_write.data["error"])

        # Pure document-write tools are omitted from read-only registry
        self.assertIsNone(ro_reg.get("word_markdown_studio"))

    def test_domain_tools_execution(self):
        reg = build_external_agent_tool_registry(**self.db_kwargs, profile="all")

        # Domain 1: spec_catalogue_and_files
        r1 = reg.execute(
            "spec_catalogue_and_files",
            {"action": "get", "spec_number": "23.501"},
        )
        self.assertTrue(r1.ok and r1.data["ok"])
        self.assertEqual(len(r1.data["versions"]), 2)

        # Domain 2: spec_text_and_clauses (search_text, diff_versions, evolution_matrix)
        r2_search = reg.execute(
            "spec_text_and_clauses",
            {"action": "search_text", "query": "SNPN onboarding", "spec_number": "23.501"},
        )
        self.assertTrue(r2_search.ok and r2_search.data["ok"])
        self.assertEqual(r2_search.data["count"], 1)

        r2_diff = reg.execute(
            "spec_text_and_clauses",
            {
                "action": "diff_versions",
                "spec_number": "23.501",
                "clause_id": "5.5.1",
                "old_version": "17.0.0",
                "new_version": "18.0.0",
            },
        )
        self.assertTrue(r2_diff.ok and r2_diff.data["ok"])
        self.assertIn("SNPN onboarding", r2_diff.data["unified_diff"])

        r2_evo = reg.execute(
            "spec_text_and_clauses",
            {"action": "evolution_matrix", "spec_number": "23.501", "clause_id": "5.5.1"},
        )
        self.assertTrue(r2_evo.ok and r2_evo.data["ok"])
        self.assertGreaterEqual(r2_evo.data["versions_checked"], 1)

        # Domain 4: protocol_inspector
        r4_list = reg.execute("protocol_inspector", {"action": "list_protocols"})
        self.assertTrue(r4_list.ok and r4_list.data["ok"])

        r4_msg = reg.execute(
            "protocol_inspector",
            {
                "action": "find_message",
                "version": "18.0.0",
                "query": "REGISTRATION REQUEST",
            },
        )
        self.assertTrue(r4_msg.ok and r4_msg.data["ok"])
        self.assertEqual(len(r4_msg.data["information_elements"]), 2)

        # Domain 5: onsite_meeting_cockpit (network_status + where_should_we_go_now)
        r5_net = reg.execute("onsite_meeting_cockpit", {"action": "network_status", "wg": "SA2"})
        self.assertTrue(r5_net.ok and r5_net.data["ok"])
        self.assertIn("is_local_reachable", r5_net.data)
        self.assertIn("local_server_base", r5_net.data)

        r5_now = reg.execute("onsite_meeting_cockpit", {"action": "where_should_we_go_now", "wg": "SA2"})
        self.assertTrue(r5_now.ok)

        # Domain 9: cr_pcr_generator (preview_filename)
        r9_prev = reg.execute(
            "cr_pcr_generator",
            {
                "action": "preview_filename",
                "spec_number": "23.501",
                "version": "18.0.0",
                "tdoc_id": "S2-2601234",
                "clauses": "5.5.1",
                "title": "PEI verification for SNPN",
            },
        )
        self.assertTrue(r9_prev.ok and r9_prev.data["ok"])
        self.assertIn("23501", r9_prev.data["proposed_filename"])

        # Domain 11: plantuml_office_studio (list_presets + extract_puml_from_vsdx constraint check)
        r11_presets = reg.execute("plantuml_office_studio", {"action": "list_presets"})
        self.assertTrue(r11_presets.ok and r11_presets.data["ok"])
        self.assertGreater(r11_presets.data["count"], 0)

        # Create a dummy .vsdx zip without 3GPP Tools PlantUML metadata and verify the explicit constraint message
        dummy_vsdx = self.tmp_path / "external_visio.vsdx"
        with zipfile.ZipFile(dummy_vsdx, "w") as zf:
            zf.writestr("visio/document.xml", "<VisioDocument></VisioDocument>")

        r11_vsdx = reg.execute(
            "plantuml_office_studio",
            {"action": "extract_puml_from_vsdx", "input_path": str(dummy_vsdx)},
        )
        self.assertTrue(r11_vsdx.ok)
        self.assertFalse(r11_vsdx.data["ok"])
        self.assertIn("generated by 3GPP Tools", r11_vsdx.data["error"])

    def test_mcp_server_jsonrpc_newline_and_content_length(self):
        reg = build_external_agent_tool_registry(**self.db_kwargs, profile="knowledge")
        server = MCPServer(reg, profile="knowledge")

        # 1. Test newline-delimited JSON-RPC stream
        in_buf = io.BytesIO()
        out_buf = io.BytesIO()

        requests = [
            {
                "jsonrpc": "2.0",
                "id": 1,
                "method": "initialize",
                "params": {
                    "protocolVersion": "2024-11-05",
                    "clientInfo": {"name": "claude-code", "version": "1.0"},
                },
            },
            {"jsonrpc": "2.0", "method": "notifications/initialized"},
            {"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}},
            {
                "jsonrpc": "2.0",
                "id": 3,
                "method": "tools/call",
                "params": {
                    "name": "spec_catalogue_and_files",
                    "arguments": {"action": "search", "query": "23.501"},
                },
            },
            {"jsonrpc": "2.0", "id": 4, "method": "ping"},
        ]
        for req in requests:
            _write_rpc_message(in_buf, req, framing="jsonl")
        in_buf.seek(0)

        serve_stdio(server, in_stream=in_buf, out_stream=out_buf)

        out_buf.seek(0)
        responses = []
        while True:
            msg, mode = _read_next_rpc_message(out_buf)
            if msg is None:
                break
            self.assertEqual(mode, "jsonl")
            responses.append(msg)

        # Notification should not emit a response -> 4 responses total
        self.assertEqual(len(responses), 4)
        self.assertEqual(responses[0]["id"], 1)
        self.assertEqual(responses[0]["result"]["serverInfo"]["name"], "3gpp-tools")
        self.assertEqual(responses[1]["id"], 2)
        self.assertEqual(len(responses[1]["result"]["tools"]), len(PROFILE_TOOLS["knowledge"]))
        self.assertEqual(responses[2]["id"], 3)
        self.assertFalse(responses[2]["result"]["isError"])
        call_result_data = json.loads(responses[2]["result"]["content"][0]["text"])
        self.assertEqual(call_result_data["count"], 1)
        self.assertEqual(responses[3]["id"], 4)

        # 2. Test Content-Length framed JSON-RPC
        cl_in = io.BytesIO()
        cl_out = io.BytesIO()
        _write_rpc_message(cl_in, {"jsonrpc": "2.0", "id": 99, "method": "ping"}, framing="content-length")
        cl_in.seek(0)

        serve_stdio(server, in_stream=cl_in, out_stream=cl_out)
        cl_out.seek(0)
        cl_msg, cl_mode = _read_next_rpc_message(cl_out)
        self.assertEqual(cl_mode, "content-length")
        self.assertEqual(cl_msg["id"], 99)
        self.assertEqual(cl_msg["result"], {})

    def test_cli_tools_commands(self):
        rc_list = cli_main(["--profile", "onsite", "list"])
        self.assertEqual(rc_list, 0)

        rc_schema = cli_main(["schema", "onsite_meeting_cockpit"])
        self.assertEqual(rc_schema, 0)

        rc_call = cli_main(
            [
                "call",
                "onsite_meeting_cockpit",
                "--action",
                "network_status",
                "--arg",
                "wg=SA2",
            ]
        )
        self.assertEqual(rc_call, 0)

        # Also verify --args JSON payload alias works on cli_tools call
        rc_args_alias = cli_main(
            [
                "--read-only",
                "call",
                "meetings_and_tdocs",
                "--args",
                '{"action": "search_meetings", "wg": "SA2", "limit": 2}',
            ]
        )
        self.assertEqual(rc_args_alias, 0)

    def test_app01_to_app05_regressions(self):
        from unittest.mock import patch

        # APP-01: Default database path resolution points to 3gpp_spec_search.db & 3gpp_protocol_data.db
        resolved = _resolve_db_paths()
        self.assertEqual(resolved["specs_db"].name, "3gpp_data.db")
        self.assertEqual(resolved["spec_search_db"].name, "3gpp_spec_search.db")
        self.assertEqual(resolved["protocol_db"].name, "3gpp_protocol_data.db")
        self.assertEqual(resolved["spec_history_db"].name, "3gpp_spec_history.db")

        # Explicit overrides remain respected
        custom_paths = _resolve_db_paths(
            spec_search_db_path=self.spec_search_db_path,
            protocol_db_path=self.protocol_db_path,
        )
        self.assertEqual(custom_paths["spec_search_db"], self.spec_search_db_path)
        self.assertEqual(custom_paths["protocol_db"], self.protocol_db_path)

        reg = build_external_agent_tool_registry(**self.db_kwargs, profile="all")

        # APP-02: search_meetings with WG filter, limit, year, and non-matching query
        m_res = reg.execute(
            "meetings_and_tdocs",
            {"action": "search_meetings", "wg": "SA2", "limit": 1},
        )
        self.assertTrue(m_res.ok and m_res.data["ok"])
        self.assertEqual(m_res.data["count"], 1)
        self.assertEqual(m_res.data["meetings"][0]["wg_name"], "SA2")

        m_empty = reg.execute(
            "meetings_and_tdocs",
            {"action": "search_meetings", "wg": "SA2", "query": "NONEXISTENT_MEETING_99999"},
        )
        self.assertTrue(m_empty.ok and m_empty.data["ok"])
        self.assertEqual(m_empty.data["count"], 0)

        # APP-03: TDoc search, confirmed vs tentative company filtering, stance/notes, and meeting_stats
        with patch(
            "modules.assistant.core.tools.external_agent_tools._get_meetings_cache_roots",
            return_value=[self.tmp_path],
        ):
            # Confirmed-only Ericsson should match "Ericsson" and "Ericsson [Rapporteur]" (2 rows),
            # while excluding "Nokia, Ericsson?" and "Qualcomm, [Ericsson]"
            t_conf = reg.execute(
                "meetings_and_tdocs",
                {
                    "action": "search_tdocs",
                    "wg": "SA2",
                    "meeting_number": "174",
                    "source_company": "Ericsson",
                    "confirmed_company_only": True,
                },
            )
            self.assertTrue(t_conf.ok and t_conf.data["ok"])
            conf_nums = [r["tdoc_number"] for r in t_conf.data["tdocs"]]
            self.assertEqual(sorted(conf_nums), ["S2-2602001", "S2-2602004"])
            # Sidecar stance/notes should be attached to S2-2602001
            row_2001 = next(r for r in t_conf.data["tdocs"] if r["tdoc_number"] == "S2-2602001")
            self.assertEqual(row_2001["my_status"], "🟢 Support")
            self.assertIn("URSP", row_2001["my_notes"])

            # Including tentative should return all 4 rows
            t_all = reg.execute(
                "meetings_and_tdocs",
                {
                    "action": "search_tdocs",
                    "wg": "SA2",
                    "meeting_number": "174",
                    "source_company": "Ericsson",
                    "confirmed_company_only": False,
                },
            )
            self.assertTrue(t_all.ok and t_all.data["ok"])
            self.assertEqual(t_all.data["count"], 4)

            # get_or_set_stance round-trip
            stance_get = reg.execute(
                "meetings_and_tdocs",
                {"action": "get_or_set_stance", "tdoc_id": "S2-2602001"},
            )
            self.assertTrue(stance_get.ok and stance_get.data["ok"])
            self.assertEqual(stance_get.data["my_status"], "🟢 Support")

            stance_set = reg.execute(
                "meetings_and_tdocs",
                {
                    "action": "get_or_set_stance",
                    "tdoc_id": "S2-2602004",
                    "stance": "🟡 Question",
                    "notes": "Verify SNPN clause numbering",
                },
            )
            self.assertTrue(stance_set.ok and stance_set.data["ok"])
            self.assertEqual(stance_set.data["my_status"], "🟡 Question")

            # meeting_stats
            stats_res = reg.execute(
                "meetings_and_tdocs",
                {"action": "meeting_stats", "wg": "SA2", "meeting_number": "174"},
            )
            self.assertTrue(stats_res.ok and stats_res.data["ok"])
            self.assertEqual(stats_res.data["total_tdocs"], 4)
            top_map = {c["company"]: c["tdoc_count"] for c in stats_res.data["top_confirmed_contributors"]}
            self.assertEqual(top_map.get("Ericsson"), 2)

        # APP-04: SpecHistoryDatabase.search_history, query_history, attribute_clause, rank_origin_crs, tdoc_dossier
        h_query = reg.execute(
            "spec_history_and_attribution",
            {"action": "query_history", "spec_number": "23.501", "query": "SNPN"},
        )
        self.assertTrue(h_query.ok and h_query.data["ok"])
        self.assertEqual(h_query.data["count"], 1)
        self.assertEqual(h_query.data["records"][0]["cr_number"], "4100")
        self.assertIn("S2-2301234", h_query.data["records"][0]["pcr_tdocs"])

        h_attr = reg.execute(
            "spec_history_and_attribution",
            {"action": "attribute_clause", "spec_number": "23.501", "clause_id": "5.5.1"},
        )
        self.assertTrue(h_attr.ok and h_attr.data["ok"])
        self.assertGreaterEqual(h_attr.data["all_candidates_count"], 1)

        h_rank = reg.execute(
            "spec_history_and_attribution",
            {"action": "rank_origin_crs", "spec_number": "23.501", "query": "SNPN"},
        )
        self.assertTrue(h_rank.ok and h_rank.data["ok"])
        self.assertGreaterEqual(len(h_rank.data["ranked_origin_crs"]), 1)

        with patch(
            "modules.assistant.core.tools.external_agent_tools._get_meetings_cache_roots",
            return_value=[self.tmp_path],
        ):
            h_dossier = reg.execute(
                "spec_history_and_attribution",
                {
                    "action": "tdoc_dossier",
                    "spec_number": "23.501",
                    "version": "18.0.0",
                    "output_dir": str(self.tmp_path),
                },
            )
            self.assertTrue(h_dossier.ok and h_dossier.data["ok"])
            self.assertTrue(Path(h_dossier.data["markdown_dossier_path"]).exists())

        # APP-05: search_work_items with query, release ("20" and "Rel-20"), wg, and tracked_only
        wi_q = reg.execute(
            "work_items_and_emails",
            {"action": "search_work_items", "query": "6G", "release": "20", "wg": "SA2"},
        )
        self.assertTrue(wi_q.ok and wi_q.data["ok"])
        self.assertEqual(wi_q.data["count"], 1)
        self.assertEqual(wi_q.data["work_items"][0]["acronym"], "FS_6G_ARC")

        with patch(
            "modules.work_items.core.wi_settings.WorkItemsSettings.get_tracking_sets",
            return_value=({"1000002"}, set()),
        ):
            wi_tracked = reg.execute(
                "work_items_and_emails",
                {"action": "search_work_items", "tracked_only": True, "limit": 1},
            )
            self.assertTrue(wi_tracked.ok and wi_tracked.data["ok"])
            self.assertEqual(wi_tracked.data["count"], 1)
            self.assertEqual(wi_tracked.data["work_items"][0]["code"], "1000002")
            self.assertTrue(wi_tracked.data["work_items"][0]["starred"])

    def test_mcp_activity_and_monitor_ui(self) -> None:
        """Verify cross-process activity telemetry, session tracking, and McpMonitorDialog UI."""
        import os
        from unittest.mock import patch
        from PyQt5.QtWidgets import QApplication
        from core.ai.mcp_activity import (
            clear_mcp_activity_log,
            generate_call_summary,
            get_active_mcp_sessions,
            get_mcp_activity_stats,
            get_mcp_client_configurations,
            get_recent_mcp_events,
            is_pid_alive,
            record_mcp_event,
            register_mcp_session,
            unregister_mcp_session,
        )
        from core.ui.mcp_monitor_dialog import McpMonitorDialog

        # 1. Test Process Liveness & Session Registration
        cur_pid = os.getpid()
        self.assertTrue(is_pid_alive(cur_pid))
        self.assertFalse(is_pid_alive(99999999))

        test_log_file = self.tmp_path / "test_mcp_activity.jsonl"
        test_sessions_dir = self.tmp_path / "sessions"

        with patch("core.ai.mcp_activity.get_mcp_activity_log_path", return_value=test_log_file), \
             patch("core.ai.mcp_activity.get_mcp_sessions_dir", return_value=test_sessions_dir):

            # Register session
            register_mcp_session(cur_pid, client="codex", version="1.2.0", profile="all")
            sessions = get_active_mcp_sessions()
            self.assertEqual(len(sessions), 1)
            self.assertEqual(sessions[0]["client"], "codex")
            self.assertEqual(sessions[0]["pid"], cur_pid)

            # Record telemetry events
            record_mcp_event(
                "tools/call",
                client="codex",
                client_ver="1.2.0",
                pid=cur_pid,
                tool="spec_text_and_clauses",
                action="search_text",
                arguments={"spec_number": "23.501", "query": "URSP", "limit": 2},
                elapsed_ms=45.2,
                ok=True,
                raw_data={"count": 2, "items": [{"result_id": "1"}, {"result_id": "2"}]},
            )

            record_mcp_event(
                "tools/call",
                client="claude",
                client_ver="2.0",
                pid=cur_pid,
                tool="protocol_inspector",
                action="find_message",
                arguments={"protocol": "5GS_NAS", "query": "REGISTRATION REQUEST"},
                elapsed_ms=18.5,
                ok=True,
                raw_data={"count": 1, "messages": [{"name": "REGISTRATION REQUEST"}]},
            )

            # Query events & stats
            events = get_recent_mcp_events(limit=10)
            self.assertEqual(len(events), 2)
            self.assertEqual(events[0]["tool"], "protocol_inspector")
            self.assertEqual(events[1]["tool"], "spec_text_and_clauses")

            stats = get_mcp_activity_stats()
            self.assertEqual(stats["total_calls"], 2)
            self.assertEqual(stats["success_rate"], 100.0)
            self.assertIn("spec_text_and_clauses", stats["by_tool"])
            self.assertIn("protocol_inspector", stats["by_tool"])

            # Test Call Summary Generator
            summary = generate_call_summary(
                "spec_text_and_clauses",
                "search_text",
                {"spec_number": "23.501", "query": "URSP"},
                {"count": 2},
            )
            self.assertIn("TS 23.501", summary)
            self.assertIn("2 clause hits", summary)

            # Test Client Configuration Generator
            cfgs = get_mcp_client_configurations()
            self.assertIn("claude_desktop", cfgs)
            self.assertIn("codex_cli", cfgs)
            self.assertIn("claude_cli", cfgs)
            self.assertIn("cursor", cfgs)
            self.assertIn("mcpServers", cfgs["claude_desktop"]["config_json"])

            # Test Clear Log
            clear_mcp_activity_log()
            self.assertEqual(len(get_recent_mcp_events()), 0)

            # Unregister session
            unregister_mcp_session(cur_pid)
            self.assertEqual(len(get_active_mcp_sessions()), 0)

            # 2. Test McpMonitorDialog GUI in Qt environment
            app = QApplication.instance() or QApplication([])
            dialog = McpMonitorDialog(parent=None)
            self.assertIsNotNone(dialog)
            self.assertEqual(dialog.tabs.count(), 3)
            self.assertEqual(dialog.table.columnCount(), 8)

            # Populate with mock event
            record_mcp_event(
                "tools/call",
                client="codex",
                pid=cur_pid,
                tool="spec_text_and_clauses",
                action="search_text",
                arguments={"spec_number": "23.501"},
                ok=True,
                raw_data={"count": 1},
            )
            dialog.refresh_all()
            self.assertEqual(dialog.table.rowCount(), 1)
            self.assertEqual(dialog.table.item(0, 1).text(), "codex")
            self.assertEqual(dialog.table.item(0, 3).text(), "spec_text_and_clauses")

            # Selection update in inspection pane
            dialog.table.selectRow(0)
            self.assertIn("23.501", dialog.args_view.toPlainText())

            # Cleanup dialog
            dialog.close()
            dialog.deleteLater()

    def test_app06_to_app09_regressions(self):
        """
        Verify APP-06 through APP-09 regressions:
        - APP-06: Container presence not contaminated by child annotations (M vs O Need S/R)
        - APP-07: Scalar field formats are INTEGER/ENUMERATED, not SEQUENCE
        - APP-08: SIB2 description lookup preserves owning context (no InterFreqCarrierFreqInfo attribution)
        - APP-09: Embedded JSON has no NaN/Infinity and strict serialization succeeds
        """
        import math
        from mcp_server import _safe_json_dumps, _sanitize_for_json
        from modules.nas.core.parsing.asn1_base_parser import BaseAsn1DocxParser
        from modules.assistant.core.agent_models import ToolStatus
        from modules.assistant.core.knowledge.protocols import ProtocolKnowledgeService

        # 1. Test APP-09 JSON sanitization
        dirty_payload = {
            "nan_val": float("nan"),
            "inf_pos": float("inf"),
            "inf_neg": float("-inf"),
            "valid_num": 42.5,
            "nested": {"inner_nan": float("nan"), "text": "ok"},
            "items": [1, float("nan"), 3],
        }
        clean = _sanitize_for_json(dirty_payload)
        self.assertIsNone(clean["nan_val"])
        self.assertIsNone(clean["inf_pos"])
        self.assertIsNone(clean["inf_neg"])
        self.assertEqual(clean["valid_num"], 42.5)
        self.assertIsNone(clean["nested"]["inner_nan"])
        self.assertEqual(clean["nested"]["text"], "ok")
        self.assertEqual(clean["items"], [1, None, 3])

        # Strict JSON dumping produces RFC 8259 valid output without ValueError
        serialized = _safe_json_dumps(dirty_payload)
        self.assertNotIn("NaN", serialized)
        self.assertNotIn("Infinity", serialized)
        parsed = json.loads(serialized)
        self.assertIsNone(parsed["nan_val"])

        # 2. Test APP-06 and APP-07 via BaseAsn1DocxParser on synthetic fixture
        parser = BaseAsn1DocxParser([], spec_number="38.331")
        synthetic_asn1 = """
            mandatoryContainer SEQUENCE {
                childWithNeed INTEGER (0..10) OPTIONAL, -- Need S
                childOptional ENUMERATED {true} OPTIONAL -- Need R
            },
            optionalContainer SEQUENCE {
                subChild BOOLEAN
            } OPTIONAL, -- Need R
            scalarInt INTEGER (2..16) OPTIONAL, -- Need S
            scalarEnum ENUMERATED {true} OPTIONAL -- Need R
        """
        fields = parser._parse_sequence_fields(synthetic_asn1, "SEQUENCE")
        field_by_name = {f["name"]: f for f in fields}

        # APP-06: Outer presence is M, not O (Need S) or O (Need R) from child annotations
        self.assertEqual(field_by_name["mandatoryContainer"]["presence"], "M")
        self.assertEqual(field_by_name["mandatoryContainer"]["format"], "SEQUENCE")
        self.assertEqual(field_by_name["optionalContainer"]["presence"], "O (Need R)")

        # APP-07: Scalar fields expose intrinsic format, not parent SEQUENCE
        self.assertEqual(field_by_name["scalarInt"]["format"], "INTEGER")
        self.assertEqual(field_by_name["scalarEnum"]["format"], "ENUMERATED")

        # 3. Test APP-06..08 against live protocol database if imported
        service = ProtocolKnowledgeService(self.protocol_db_path)
        res = service.find_message("SIB2", protocol="NR RRC", version="19.3.0")
        if res.status == ToolStatus.FOUND and res.data:
            sib2_fields = res.data.get("fields", [])
            actual = {f["ie_name"]: f for f in sib2_fields if f.get("depth", 0) == 0}
            if actual:
                # APP-06
                for name in ("cellReselectionInfoCommon", "cellReselectionServingFreqInfo", "intraFreqCellReselectionInfo"):
                    if name in actual:
                        self.assertEqual(actual[name]["presence"], "M", f"{name} must be mandatory M")

                # APP-07
                if "cellEquivalentSize-r17" in actual:
                    self.assertEqual(actual["cellEquivalentSize-r17"]["format"], "INTEGER")
                if "uav-PrioritizedFrequency-r19" in actual:
                    self.assertEqual(actual["uav-PrioritizedFrequency-r19"]["format"], "ENUMERATED")

                # APP-08: SIB2 description must not refer to InterFreqCarrierFreqInfo-v1900
                for fname in ("uav-PrioritizedFrequency-r19", "uav-PrioritizedFrequencyAltitudeRange-r19"):
                    if fname in actual:
                        desc = str(actual[fname].get("ie_description", "") or "")
                        self.assertNotIn("Clause InterFreqCarrierFreqInfo-v1900", desc)


if __name__ == "__main__":
    unittest.main()
