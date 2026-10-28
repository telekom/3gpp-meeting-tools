# --- File: tests/test_plantuml_presets_config.py ---
"""Unit tests for PlantUML Presets Configuration & Stencil Management.

Verifies:
  1. Standalone "UE" presets and individual 3GPP Network Functions exist.
  2. All standard PlantUML presets (architectures, messages, fragments, notes) are preserved.
  3. JSON configuration persistence and roundtrip fidelity.
  4. Git-ignored user presets isolation and merge mechanism.
  5. 3GPP Call Flow starter template presence in PLANTUML_TYPES.
"""

import json
from pathlib import Path
import tempfile
import unittest

from modules.puml2visio.templates.plantuml_snippets import (
    BUILTIN_SNIPPETS,
    PLANTUML_SNIPPETS,
    PlantUMLSnippet,
    add_user_snippet,
    delete_user_snippet,
    extract_participants,
    get_default_presets_path,
    get_user_presets_path,
    load_snippets_from_file,
    load_user_snippets_from_file,
    reload_snippets,
    resolve_snippet_text,
    save_snippets_to_file,
    save_user_snippets_to_file,
)
from modules.puml2visio.templates.plantuml_templates import PLANTUML_TYPES


class TestPlantUMLPresetsConfig(unittest.TestCase):
    """Test suite for PlantUML presets and configuration architecture."""

    def test_ue_presets_presence(self):
        """Verify that UE is now available as standalone presets and variants."""
        snippet_ids = {s.id for s in PLANTUML_SNIPPETS}
        snippet_labels = {s.label for s in PLANTUML_SNIPPETS}

        # 1. Standard UE participant
        self.assertIn("part_3gpp_ue", snippet_ids, "Standard UE participant snippet must exist")
        ue_snip = next(s for s in PLANTUML_SNIPPETS if s.id == "part_3gpp_ue")
        self.assertIn('participant "UE" as UE', ue_snip.code)
        self.assertEqual(ue_snip.category, "3GPP Network Entities & NFs")

        # 2. UE as Actor
        self.assertIn("part_3gpp_ue_actor", snippet_ids, "UE as Actor snippet must exist")
        ue_actor = next(s for s in PLANTUML_SNIPPETS if s.id == "part_3gpp_ue_actor")
        self.assertIn('actor "UE" as UE', ue_actor.code)

        # 3. Dual UEs
        self.assertIn("part_3gpp_ue_dual", snippet_ids, "Dual UEs snippet must exist")

        # 4. UE split subcomponents
        self.assertIn("part_3gpp_ue_split", snippet_ids, "UE split subcomponents snippet must exist")

    def test_3gpp_network_functions_presence(self):
        """Verify that individual 3GPP 5G and 4G Network Functions are present."""
        snippet_ids = {s.id for s in PLANTUML_SNIPPETS}
        expected_nfs = [
            "part_3gpp_gnb",
            "part_3gpp_amf",
            "part_3gpp_smf",
            "part_3gpp_upf",
            "part_3gpp_udm",
            "part_3gpp_ausf",
            "part_3gpp_pcf",
            "part_3gpp_nrf",
            "part_3gpp_nssf",
            "part_3gpp_nef",
            "part_3gpp_udr",
            "part_3gpp_af",
            "part_3gpp_sepp",
            "part_3gpp_dn",
            "part_3gpp_enb",
            "part_3gpp_mme",
            "part_3gpp_sgw",
            "part_3gpp_pgw",
            "part_3gpp_hss",
        ]
        for nf_id in expected_nfs:
            self.assertIn(nf_id, snippet_ids, f"Expected 3GPP NF preset '{nf_id}' to be present")

    def test_all_original_presets_preserved(self):
        """Verify that all previous architectural, message, logic, and note presets are retained."""
        snippet_ids = {s.id for s in PLANTUML_SNIPPETS}
        original_sample_ids = [
            # 3GPP Skeletons
            "3gpp_5gs_skeleton",
            "3gpp_5gs_basic",
            "3gpp_eps_skeleton",
            "3gpp_roaming_skeleton",
            "3gpp_oran_skeleton",
            "3gpp_ims_skeleton",
            # Standard UML elements
            "part_standard",
            "part_actor",
            "part_database",
            "part_boundary",
            "part_box",
            # Messages
            "msg_sync_req",
            "msg_async_resp",
            "msg_colored",
            "msg_reject",
            "msg_lost",
            # Logic
            "frag_alt",
            "frag_opt",
            "frag_loop",
            "frag_par",
            "frag_critical",
            # Notes & Lifelines
            "note_over",
            "note_hnote",
            "note_rnote",
            "lifeline_activate",
            "lifeline_deactivate",
            # Multi-diagram
            "state_trans",
            "act_if_else",
            "timing_signals",
            "packet_fields",
            "dir_title_header",
        ]
        for orig_id in original_sample_ids:
            self.assertIn(orig_id, snippet_ids, f"Original preset '{orig_id}' must remain preserved")

    def test_snippet_serialization_roundtrip(self):
        """Verify dataclass to_dict and from_dict serialization roundtrip."""
        original = PlantUMLSnippet(
            id="test_custom_id",
            label="My Test Snippet",
            category="Test Category",
            diagram_scope=("Sequence", "State"),
            description="A test description",
            code="participant \"Test\" as T\n{cursor}",
            is_block_wrapper=True,
            is_user_defined=True,
        )
        d = original.to_dict()
        restored = PlantUMLSnippet.from_dict(d)

        self.assertEqual(original.id, restored.id)
        self.assertEqual(original.label, restored.label)
        self.assertEqual(original.category, restored.category)
        self.assertEqual(original.diagram_scope, restored.diagram_scope)
        self.assertEqual(original.description, restored.description)
        self.assertEqual(original.code, restored.code)
        self.assertEqual(original.is_block_wrapper, restored.is_block_wrapper)
        self.assertEqual(original.is_user_defined, restored.is_user_defined)

    def test_default_config_file_exists_and_matches(self):
        """Verify that src/config/plantuml_presets.json exists, is valid JSON, and has all builtins."""
        config_path = get_default_presets_path()
        self.assertTrue(config_path.exists(), "Default presets file plantuml_presets.json must exist on disk")

        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        self.assertIn("snippets", data)
        self.assertGreaterEqual(len(data["snippets"]), len(BUILTIN_SNIPPETS))

        file_ids = {s["id"] for s in data["snippets"]}
        for builtin in BUILTIN_SNIPPETS:
            self.assertIn(builtin.id, file_ids, f"Builtin '{builtin.id}' must be serialized in plantuml_presets.json")

    def test_user_presets_add_and_delete(self):
        """Verify adding and deleting user presets dynamically without touching built-in presets."""
        initial_count = len(PLANTUML_SNIPPETS)
        test_id = "user_unit_test_preset_99"

        user_snippet = PlantUMLSnippet(
            id=test_id,
            label="Temporary Unit Test Preset",
            category="User Defined Presets",
            diagram_scope=("Sequence",),
            description="Should be removed after test",
            code="participant \"TempTest\" as TT\n{cursor}",
            is_block_wrapper=False,
            is_user_defined=True,
        )

        try:
            # 1. Add user snippet
            add_user_snippet(user_snippet)
            self.assertEqual(len(PLANTUML_SNIPPETS), initial_count + 1)
            matching = [s for s in PLANTUML_SNIPPETS if s.id == test_id]
            self.assertEqual(len(matching), 1)
            self.assertTrue(matching[0].is_user_defined)

            # 2. Check that it exists in the user config file
            user_path = get_user_presets_path()
            self.assertTrue(user_path.exists())
            user_list = load_user_snippets_from_file(user_path)
            self.assertTrue(any(s.id == test_id for s in user_list))

        finally:
            # 3. Clean up by deleting
            deleted = delete_user_snippet(test_id)
            self.assertTrue(deleted)
            self.assertEqual(len(PLANTUML_SNIPPETS), initial_count)
            self.assertFalse(any(s.id == test_id for s in PLANTUML_SNIPPETS))

    def test_resolve_snippet_text_with_ue(self):
        """Verify dynamic participant resolution defaults to UE when participants are empty."""
        msg_snip = next(s for s in PLANTUML_SNIPPETS if s.id == "msg_sync_req")
        resolved, start_idx, end_idx = resolve_snippet_text("", msg_snip)
        self.assertIn("UE -> gNB: Request Message", resolved)

    def test_3gpp_template_in_plantuml_types(self):
        """Verify that '3GPP Call Flow (5GS)' is available in PLANTUML_TYPES."""
        self.assertIn("3GPP Call Flow (5GS)", PLANTUML_TYPES)
        template_info = PLANTUML_TYPES["3GPP Call Flow (5GS)"]
        self.assertIn("template", template_info)
        template_text = template_info["template"]
        self.assertIn('participant "UE" as UE', template_text)
        self.assertIn('participant "gNB" as gNB', template_text)
        self.assertIn('participant "AMF" as AMF', template_text)


if __name__ == "__main__":
    unittest.main()
