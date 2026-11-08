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
    build_external_agent_tool_registry,
)
from modules.nas.core.nas_db import NASDatabase
from modules.spec_search.core.spec_search_db import SpecSearchDatabase
from modules.specifications.core.database import SpecsDatabase


class TestMcpAndCliTools(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp_path = Path(self._tmp.name)

        self.specs_db_path = self.tmp_path / "specs.db"
        self.spec_search_db_path = self.tmp_path / "spec_search.db"
        self.protocol_db_path = self.tmp_path / "protocol.db"

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

        self.db_kwargs = {
            "specs_db_path": self.specs_db_path,
            "spec_search_db_path": self.spec_search_db_path,
            "protocol_db_path": self.protocol_db_path,
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


if __name__ == "__main__":
    unittest.main()
