# --- File: src/modules/puml2visio/templates/plantuml_snippets.py ---
"""Declarative catalog of PlantUML syntax elements, snippets, and 3GPP presets.

Supports dynamic placeholder resolution (participants, selection wrapping, cursor placement),
JSON configuration persistence (`src/config/plantuml_presets.json`), and git-ignored
user-defined presets (`src/config/plantuml_presets_user.json`).
"""

from dataclasses import dataclass, field
import json
import logging
from pathlib import Path
import re
from typing import Callable, Dict, List, Optional, Tuple

import core.utils.paths


@dataclass(frozen=True)
class PlantUMLSnippet:
    """Represents a reusable PlantUML syntax element or pattern."""
    id: str
    label: str
    category: str
    diagram_scope: Tuple[str, ...]
    description: str
    code: str
    is_block_wrapper: bool = False
    is_user_defined: bool = False

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "label": self.label,
            "category": self.category,
            "diagram_scope": list(self.diagram_scope),
            "description": self.description,
            "code": self.code,
            "is_block_wrapper": self.is_block_wrapper,
            "is_user_defined": self.is_user_defined,
        }

    @classmethod
    def from_dict(cls, data: dict, is_user_defined: bool = False) -> "PlantUMLSnippet":
        scope_raw = data.get("diagram_scope", ["All"])
        if isinstance(scope_raw, list):
            scope = tuple(str(s) for s in scope_raw)
        elif isinstance(scope_raw, str):
            scope = (scope_raw,)
        else:
            scope = tuple(scope_raw)

        return cls(
            id=str(data.get("id", "")),
            label=str(data.get("label", "")),
            category=str(data.get("category", "General")),
            diagram_scope=scope,
            description=str(data.get("description", "")),
            code=str(data.get("code", "")),
            is_block_wrapper=bool(data.get("is_block_wrapper", False)),
            is_user_defined=bool(data.get("is_user_defined", is_user_defined)),
        )


# Regex to detect declared sequence participants/actors/entities/etc.
_PARTICIPANT_DECL_RE = re.compile(
    r'^\s*(?:participant|actor|boundary|control|entity|database|collections|queue)\s+'
    r'(?:"[^"]*"\s+as\s+([A-Za-z0-9_]+)|([A-Za-z0-9_]+))',
    re.MULTILINE
)

# Regex to detect implicit participants from arrow statements (A -> B)
_ARROW_PARTICIPANTS_RE = re.compile(
    r'^\s*([A-Za-z0-9_]+)\s*[-<~=]+(?:\[#[^\]]+\])?[-~=>x*o\\\/]+\s*([A-Za-z0-9_]+)\s*:',
    re.MULTILINE
)


def extract_participants(puml_text: str) -> List[str]:
    """Scans PlantUML text and returns an ordered, deduplicated list of participant aliases."""
    if not puml_text:
        return []

    participants: List[str] = []
    seen = set()

    def add(alias: str):
        if not alias:
            return
        if alias.lower() in {"start", "stop", "end", "return", "note", "activate", "deactivate", "alt", "opt"}:
            return
        if alias not in seen:
            seen.add(alias)
            participants.append(alias)

    # 1. Scan explicit declarations
    for match in _PARTICIPANT_DECL_RE.finditer(puml_text):
        alias = match.group(1) or match.group(2)
        if alias:
            add(alias.strip())

    # 2. Scan arrow message exchanges for any additional participants
    for match in _ARROW_PARTICIPANTS_RE.finditer(puml_text):
        p1 = match.group(1).strip()
        p2 = match.group(2).strip()
        add(p1)
        add(p2)

    return participants


def resolve_snippet_text(
    puml_text: str,
    snippet: PlantUMLSnippet,
    selected_text: str = "",
    base_indent: str = ""
) -> Tuple[str, int, int]:
    """Resolves dynamic placeholders in snippet code.

    Placeholders supported:
      {p1}: First detected participant (defaults to 'UE')
      {p2}: Second detected participant (defaults to 'gNB' or 'NW')
      {selection}: Selected text if wrapping, or default placeholder content
      {cursor}: Position where the editor cursor should be placed

    Returns:
      (resolved_text, selection_start_offset, selection_end_offset)
    """
    participants = extract_participants(puml_text)
    p1 = participants[0] if len(participants) > 0 else "UE"
    p2 = participants[1] if len(participants) > 1 else ("gNB" if p1 == "UE" else "NW")

    raw_code = snippet.code

    # Handle {selection} placeholder
    if "{selection}" in raw_code:
        if selected_text and snippet.is_block_wrapper:
            lines = selected_text.splitlines()
            indented_lines = ["  " + line if line.strip() else line for line in lines]
            inner_content = "\n".join(indented_lines)
        else:
            inner_content = "  ' TODO: Add messages or logic here"
        raw_code = raw_code.replace("{selection}", inner_content)

    # Replace participant tokens
    resolved = raw_code.replace("{p1}", p1).replace("{p2}", p2)

    # Indent multi-line snippets to match base_indent
    if base_indent:
        lines = resolved.splitlines()
        if lines:
            indented = [lines[0]] + [base_indent + l if l.strip() else l for l in lines[1:]]
            resolved = "\n".join(indented)

    # Locate cursor offset
    cursor_marker = "{cursor}"
    if cursor_marker in resolved:
        start_idx = resolved.find(cursor_marker)
        resolved = resolved.replace(cursor_marker, "", 1)
        end_marker = "{/cursor}"
        if end_marker in resolved:
            end_idx = resolved.find(end_marker)
            resolved = resolved.replace(end_marker, "", 1)
            return resolved, start_idx, end_idx
        return resolved, start_idx, start_idx

    return resolved, len(resolved), len(resolved)


# =========================================================================
# --- CANONICAL BUILT-IN PRESETS (SHIPPED DEFAULTS) ---
# =========================================================================

BUILTIN_SNIPPETS: List[PlantUMLSnippet] = [
    # ---------------------------------------------------------------------
    # 1. 3GPP NETWORK ENTITIES & NFS (Individual Nodes & Stencils)
    # ---------------------------------------------------------------------
    PlantUMLSnippet(
        id="part_3gpp_ue",
        label="UE (User Equipment)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="Standard 3GPP User Equipment participant",
        code="""participant "UE" as UE
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_ue_actor",
        label="UE as Actor (Human / Mobile Device)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence", "Use Case"),
        description="User Equipment declared as a stick-figure actor",
        code="""actor "UE" as UE
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_ue_dual",
        label="Dual UEs (UE-A & UE-B for Sidelink / ProSe / P2P)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="Two distinct UEs for sidelink (PC5), ProSe, or peer-to-peer call flows",
        code="""participant "UE A" as UEA
participant "UE B" as UEB
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_ue_split",
        label="UE Subcomponents (App/TE, MT, USIM)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="UE split into Application (TE), Mobile Termination (MT), and USIM",
        code="""box "User Equipment (UE)" #F8FAFC
  participant "App / TE" as TE
  participant "UE / MT" as UE
  participant "USIM" as USIM
end box
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_gnb",
        label="gNB / (R)AN (5G Radio Access Network)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="5G NR Next Generation NodeB / Radio Access Network node",
        code="""participant "gNB" as gNB
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_amf",
        label="AMF (Access and Mobility Management Function)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="5G Core Access and Mobility Management Function",
        code="""participant "AMF" as AMF
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_smf",
        label="SMF (Session Management Function)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="5G Core Session Management Function",
        code="""participant "SMF" as SMF
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_upf",
        label="UPF (User Plane Function)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="5G Core User Plane Function (GTP-U termination)",
        code="""participant "UPF" as UPF
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_udm",
        label="UDM (Unified Data Management)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="5G Core Unified Data Management node",
        code="""participant "UDM" as UDM
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_ausf",
        label="AUSF (Authentication Server Function)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="5G Core Authentication Server Function (5G AKA)",
        code="""participant "AUSF" as AUSF
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_pcf",
        label="PCF (Policy Control Function)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="5G Core Policy Control Function",
        code="""participant "PCF" as PCF
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_nrf",
        label="NRF (Network Repository Function)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="5G Core Network Repository Function (NF discovery)",
        code="""participant "NRF" as NRF
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_nssf",
        label="NSSF (Network Slice Selection Function)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="5G Core Network Slice Selection Function",
        code="""participant "NSSF" as NSSF
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_nef",
        label="NEF (Network Exposure Function)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="5G Core Network Exposure Function",
        code="""participant "NEF" as NEF
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_udr",
        label="UDR (Unified Data Repository)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="5G Core Unified Data Repository database",
        code="""database "UDR" as UDR
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_af",
        label="AF (Application Function)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="5G Application Function (external or internal)",
        code="""participant "AF" as AF
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_sepp",
        label="SEPP (Security Edge Protection Proxy)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="5G Roaming Security Edge Protection Proxy (N32 interface)",
        code="""boundary "SEPP" as SEPP
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_dn",
        label="DN (Data Network / Internet)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="External Data Network (Internet / IMS / Corporate)",
        code="""participant "DN" as DN
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_enb",
        label="eNB (E-UTRAN Node B / 4G LTE RAN)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="4G LTE Evolved NodeB base station",
        code="""participant "eNB" as eNB
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_mme",
        label="MME (Mobility Management Entity)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="4G EPC Mobility Management Entity",
        code="""participant "MME" as MME
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_sgw",
        label="S-GW (Serving Gateway)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="4G EPC Serving Gateway",
        code="""participant "S-GW" as SGW
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_pgw",
        label="P-GW (PDN Gateway)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="4G EPC Packet Data Network Gateway",
        code="""participant "P-GW" as PGW
{cursor}""",
    ),
    PlantUMLSnippet(
        id="part_3gpp_hss",
        label="HSS (Home Subscriber Server)",
        category="3GPP Network Entities & NFs",
        diagram_scope=("Sequence",),
        description="4G EPC Home Subscriber Server database",
        code="""database "HSS" as HSS
{cursor}""",
    ),

    # ---------------------------------------------------------------------
    # 2. 3GPP ARCHITECTURE PRESETS (Multi-Node Composite Skeletons)
    # ---------------------------------------------------------------------
    PlantUMLSnippet(
        id="3gpp_5gs_skeleton",
        label="5GS Core & RAN Architecture (Full)",
        category="3GPP Architecture Presets",
        diagram_scope=("Sequence",),
        description="Standard 3GPP 5G System participants (UE, (R)AN, AMF, SMF, UPF, AUSF, UDM, PCF, NRF)",
        code="""participant "UE" as UE
box "Radio Access Network" #F8FAFC
  participant "(R)AN" as RAN
end box
box "5G Core Control Plane" #F1F5F9
  participant "AMF" as AMF
  participant "SMF" as SMF
  participant "AUSF" as AUSF
  participant "UDM" as UDM
  participant "PCF" as PCF
  participant "NRF" as NRF
end box
box "User Plane" #F8FAFC
  participant "UPF" as UPF
  participant "DN" as DN
end box

{cursor}""",
    ),
    PlantUMLSnippet(
        id="3gpp_5gs_basic",
        label="5GS Basic Call Flow (UE, gNB, AMF, SMF, UPF)",
        category="3GPP Architecture Presets",
        diagram_scope=("Sequence",),
        description="Minimal 5GS participant set for typical signaling and PDU session flows",
        code="""participant "UE" as UE
participant "gNB" as gNB
participant "AMF" as AMF
participant "SMF" as SMF
participant "UPF" as UPF

{cursor}""",
    ),
    PlantUMLSnippet(
        id="3gpp_eps_skeleton",
        label="EPS / 4G Core & RAN Architecture",
        category="3GPP Architecture Presets",
        diagram_scope=("Sequence",),
        description="Standard 3GPP 4G EPS entities (UE, eNB, MME, SGW, PGW, HSS, PCRF)",
        code="""participant "UE" as UE
participant "eNB" as eNB
box "Evolved Packet Core (EPC)" #F1F5F9
  participant "MME" as MME
  participant "S-GW" as SGW
  participant "P-GW" as PGW
  participant "HSS" as HSS
  participant "PCRF" as PCRF
end box

{cursor}""",
    ),
    PlantUMLSnippet(
        id="3gpp_roaming_skeleton",
        label="5GS Roaming (VPLMN & HPLMN)",
        category="3GPP Architecture Presets",
        diagram_scope=("Sequence",),
        description="5G Roaming architecture with VPLMN, HPLMN and SEPP security proxies",
        code="""box "Visited PLMN (VPLMN)" #F8FAFC
  participant "UE" as UE
  participant "V-RAN" as VRAN
  participant "V-AMF" as VAMF
  participant "V-SMF" as VSMF
  participant "V-UPF" as VUPF
  participant "vSEPP" as VSEPP
end box
box "Home PLMN (HPLMN)" #F1F5F9
  participant "hSEPP" as HSEPP
  participant "H-UDM" as HUDM
  participant "H-AUSF" as HAUSF
  participant "H-SMF" as HSMF
  participant "H-UPF" as HUPF
end box

{cursor}""",
    ),
    PlantUMLSnippet(
        id="3gpp_oran_skeleton",
        label="O-RAN Architecture (RU, DU, CU-CP, CU-UP, RIC)",
        category="3GPP Architecture Presets",
        diagram_scope=("Sequence",),
        description="O-RAN split RAN nodes and Near-RT RIC",
        code="""participant "UE" as UE
box "O-RAN Nodes" #F8FAFC
  participant "O-RU" as ORU
  participant "O-DU" as ODU
  participant "O-CU-UP" as CUUP
  participant "O-CU-CP" as CUCP
  participant "Near-RT RIC" as RIC
end box
participant "5GC" as 5GC

{cursor}""",
    ),
    PlantUMLSnippet(
        id="3gpp_ims_skeleton",
        label="IMS VoNR / VoLTE Architecture",
        category="3GPP Architecture Presets",
        diagram_scope=("Sequence",),
        description="IP Multimedia Subsystem core nodes for voice & messaging",
        code="""participant "UE" as UE
participant "gNB" as gNB
participant "UPF" as UPF
box "IMS Core" #F1F5F9
  participant "P-CSCF" as PCSCF
  participant "I-CSCF" as ICSCF
  participant "S-CSCF" as SCSCF
  participant "HSS" as HSS
  participant "TAS" as TAS
end box

{cursor}""",
    ),

    # ---------------------------------------------------------------------
    # 3. PARTICIPANTS & GROUPING (Generic UML)
    # ---------------------------------------------------------------------
    PlantUMLSnippet(
        id="part_standard",
        label="Participant (with alias)",
        category="Participants & Grouping",
        diagram_scope=("Sequence",),
        description="Declare a participant box with human display name and code alias",
        code='participant "{cursor}Entity Name" as myEntity',
    ),
    PlantUMLSnippet(
        id="part_actor",
        label="Actor (Human User)",
        category="Participants & Grouping",
        diagram_scope=("Sequence", "Use Case"),
        description="Declare an actor stick-figure icon",
        code='actor "{cursor}User" as user',
    ),
    PlantUMLSnippet(
        id="part_database",
        label="Database (Cylinder Entity)",
        category="Participants & Grouping",
        diagram_scope=("Sequence", "Component"),
        description="Declare a database entity (e.g. UDR, MySQL, Redis)",
        code='database "{cursor}Database / UDR" as db',
    ),
    PlantUMLSnippet(
        id="part_boundary",
        label="Boundary / Interface Element",
        category="Participants & Grouping",
        diagram_scope=("Sequence",),
        description="Declare a boundary circle entity (e.g. API Gateway, SEPP)",
        code='boundary "{cursor}Gateway" as gw',
    ),
    PlantUMLSnippet(
        id="part_box",
        label="Box Group (Enclosing Background)",
        category="Participants & Grouping",
        diagram_scope=("Sequence",),
        description="Enclose participants inside a visual colored box",
        code="""box "{cursor}Group Title" #F8FAFC
{selection}
end box""",
        is_block_wrapper=True,
    ),
    PlantUMLSnippet(
        id="autonumber_start",
        label="Autonumber Steps [01]",
        category="Participants & Grouping",
        diagram_scope=("Sequence",),
        description="Enable automatic step numbering on all sequence messages",
        code='autonumber 1 1 "<b>[00]</b>"',
    ),
    PlantUMLSnippet(
        id="autonumber_pause",
        label="Autonumber Stop / Resume",
        category="Participants & Grouping",
        diagram_scope=("Sequence",),
        description="Temporarily pause or resume step numbering",
        code="""autonumber stop
' {cursor}Unnumbered internal actions
autonumber resume""",
    ),

    # ---------------------------------------------------------------------
    # 4. MESSAGES & ARROWS (Sequence)
    # ---------------------------------------------------------------------
    PlantUMLSnippet(
        id="msg_sync_req",
        label="Synchronous Request (->)",
        category="Messages & Arrows",
        diagram_scope=("Sequence",),
        description="Solid arrow for request or synchronous procedure call",
        code="{p1} -> {p2}: {cursor}Request Message",
    ),
    PlantUMLSnippet(
        id="msg_async_resp",
        label="Asynchronous / Response (-->)",
        category="Messages & Arrows",
        diagram_scope=("Sequence",),
        description="Dashed arrow for response, acknowledgement, or return signal",
        code="{p2} --> {p1}: {cursor}Response / ACK",
    ),
    PlantUMLSnippet(
        id="msg_colored",
        label="Colored / Highlighted Message (-[#color]->)",
        category="Messages & Arrows",
        diagram_scope=("Sequence",),
        description="Solid colored arrow for primary or emphasized signaling",
        code="{p1} -[#1E5C99]-> {p2}: {cursor}Highlighted Message",
    ),
    PlantUMLSnippet(
        id="msg_reject",
        label="Reject / Failure Message (-[#red]->)",
        category="Messages & Arrows",
        diagram_scope=("Sequence",),
        description="Red arrow for error, reject, or negative response",
        code="{p2} -[#D32F2F]-> {p1}: {cursor}Reject / Failure (Cause)",
    ),
    PlantUMLSnippet(
        id="msg_lost",
        label="Lost / Dropped Message (->x)",
        category="Messages & Arrows",
        diagram_scope=("Sequence",),
        description="Message that is lost in transit or dropped by channel",
        code="{p1} ->x {p2}: {cursor}Lost / Dropped Packet",
    ),
    PlantUMLSnippet(
        id="msg_self",
        label="Self-Call / Internal Action",
        category="Messages & Arrows",
        diagram_scope=("Sequence",),
        description="Internal processing step within the same entity",
        code="{p1} -> {p1}: {cursor}Internal verification / processing",
    ),
    PlantUMLSnippet(
        id="msg_bidir",
        label="Bi-directional Data Flow (<->)",
        category="Messages & Arrows",
        diagram_scope=("Sequence",),
        description="Two-way communication or active GTP/IPsec tunnel",
        code="{p1} <-> {p2}: {cursor}Bi-directional Data / Tunnel",
    ),
    PlantUMLSnippet(
        id="msg_broadcast",
        label="Broadcast to All (->*)",
        category="Messages & Arrows",
        diagram_scope=("Sequence",),
        description="Broadcast message sent to all listening participants",
        code="{p1} ->*: {cursor}System Information Broadcast (SIB)",
    ),
    PlantUMLSnippet(
        id="msg_create",
        label="Create Instance Dynamic Participant",
        category="Messages & Arrows",
        diagram_scope=("Sequence",),
        description="Dynamically instantiate a participant halfway through the flow",
        code="""create {p2}
{p1} -> {p2}: <<create>> {cursor}Initialize Session""",
    ),

    # ---------------------------------------------------------------------
    # 5. LOGIC & GROUPING FRAGMENTS (Sequence)
    # ---------------------------------------------------------------------
    PlantUMLSnippet(
        id="frag_alt",
        label="Alt / Else (Alternative Branches)",
        category="Logic & Grouping",
        diagram_scope=("Sequence",),
        description="Conditional branching (e.g. Success vs Failure)",
        code="""alt {cursor}Success Case
{selection}
else Failure Case
  {p2} --> {p1}: Reject
end""",
        is_block_wrapper=True,
    ),
    PlantUMLSnippet(
        id="frag_opt",
        label="Opt (Optional Procedure)",
        category="Logic & Grouping",
        diagram_scope=("Sequence",),
        description="Optional block executed only if conditions are met",
        code="""opt {cursor}If Security Capabilities Required
{selection}
end""",
        is_block_wrapper=True,
    ),
    PlantUMLSnippet(
        id="frag_loop",
        label="Loop (Repeated Signaling)",
        category="Logic & Grouping",
        diagram_scope=("Sequence",),
        description="Repeated loop (e.g. retransmission, periodic keepalive)",
        code="""loop {cursor}For each Bearer / Until T3512 expires
{selection}
end""",
        is_block_wrapper=True,
    ),
    PlantUMLSnippet(
        id="frag_par",
        label="Par (Parallel Execution)",
        category="Logic & Grouping",
        diagram_scope=("Sequence",),
        description="Concurrent signaling branches running simultaneously",
        code="""par {cursor}Parallel Branch A
{selection}
else Parallel Branch B
  {p1} -> {p2}: Concurrent signaling
end""",
        is_block_wrapper=True,
    ),
    PlantUMLSnippet(
        id="frag_critical",
        label="Critical (Atomic Section)",
        category="Logic & Grouping",
        diagram_scope=("Sequence",),
        description="Atomic procedure that cannot be interrupted",
        code="""critical {cursor}Atomic Key Handshake
{selection}
end""",
        is_block_wrapper=True,
    ),
    PlantUMLSnippet(
        id="frag_break",
        label="Break (Early Exit / Abort)",
        category="Logic & Grouping",
        diagram_scope=("Sequence",),
        description="Early termination or exception handling",
        code="""break {cursor}On Integrity Check Failure
  {p1} ->x {p2}: Abort Connection
end""",
    ),
    PlantUMLSnippet(
        id="frag_group",
        label="Group (Custom Labeled Frame)",
        category="Logic & Grouping",
        diagram_scope=("Sequence",),
        description="Frame arbitrary messages with a custom label header",
        code="""group {cursor}Mutual Authentication Procedure [TS 33.501]
{selection}
end""",
        is_block_wrapper=True,
    ),
    PlantUMLSnippet(
        id="frag_ref",
        label="Ref (Reference to Other Spec / Clause)",
        category="Logic & Grouping",
        diagram_scope=("Sequence",),
        description="Reference an external 3GPP procedure or specification clause",
        code="ref over {p1}, {p2} : {cursor}See 3GPP TS 23.502 Clause 4.2.2.2.2",
    ),

    # ---------------------------------------------------------------------
    # 6. NOTES, LIFELINES & STATE (Sequence)
    # ---------------------------------------------------------------------
    PlantUMLSnippet(
        id="note_over",
        label="Note over Participants",
        category="Notes & Lifelines",
        diagram_scope=("Sequence",),
        description="Annotation spanning two or more participants",
        code="note over {p1}, {p2}: {cursor}Verify Subscription & Slice (S-NSSAI)",
    ),
    PlantUMLSnippet(
        id="note_hnote",
        label="Hexagonal State Note (hnote over)",
        category="Notes & Lifelines",
        diagram_scope=("Sequence",),
        description="Hexagonal state box commonly used for 3GPP RRC/NAS states (e.g. 5GMM-REGISTERED)",
        code="hnote over {p1}: {cursor}5GMM-REGISTERED / RRC_CONNECTED",
    ),
    PlantUMLSnippet(
        id="note_rnote",
        label="Rectangle State Note (rnote over)",
        category="Notes & Lifelines",
        diagram_scope=("Sequence",),
        description="Rectangular state box for entity status",
        code="rnote over {p1}: {cursor}CM-CONNECTED",
    ),
    PlantUMLSnippet(
        id="note_right",
        label="Note Right of Participant",
        category="Notes & Lifelines",
        diagram_scope=("Sequence",),
        description="Single-line note docked to the right of an entity",
        code="note right of {p1}: {cursor}Start Timer T3510",
    ),
    PlantUMLSnippet(
        id="note_left",
        label="Note Left of Participant",
        category="Notes & Lifelines",
        diagram_scope=("Sequence",),
        description="Single-line note docked to the left of an entity",
        code="note left of {p1}: {cursor}Store Security Context",
    ),
    PlantUMLSnippet(
        id="note_multiline",
        label="Multi-line Note Block",
        category="Notes & Lifelines",
        diagram_scope=("Sequence",),
        description="Detailed multi-line note block with bullet points",
        code="""note over {p1}, {p2}
  {cursor}* IE: 5GS Mobile Identity (SUCI)
  * IE: UE Security Capabilities
  * IE: Requested NSSAI
end note""",
        is_block_wrapper=True,
    ),
    PlantUMLSnippet(
        id="note_across",
        label="Note Across All Lifelines",
        category="Notes & Lifelines",
        diagram_scope=("Sequence",),
        description="Wide banner note spanning across the entire diagram",
        code="note across: {cursor}=== End of Phase 1 ===\\nAll security keys established",
    ),
    PlantUMLSnippet(
        id="lifeline_activate",
        label="Activate Lifeline",
        category="Notes & Lifelines",
        diagram_scope=("Sequence",),
        description="Show participant active processing bar",
        code="activate {p1}\n{cursor}",
    ),
    PlantUMLSnippet(
        id="lifeline_deactivate",
        label="Deactivate Lifeline",
        category="Notes & Lifelines",
        diagram_scope=("Sequence",),
        description="End participant active processing bar",
        code="deactivate {p1}\n{cursor}",
    ),
    PlantUMLSnippet(
        id="lifeline_block",
        label="Activation Block (activate ... deactivate)",
        category="Notes & Lifelines",
        diagram_scope=("Sequence",),
        description="Wrap actions inside an active processing bar",
        code="""activate {p1}
{selection}
deactivate {p1}""",
        is_block_wrapper=True,
    ),
    PlantUMLSnippet(
        id="msg_return",
        label="Return Message (from active lifeline)",
        category="Notes & Lifelines",
        diagram_scope=("Sequence",),
        description="Return message that automatically deactivates the target lifeline",
        code="return {cursor}Status OK / Response Payload",
    ),
    PlantUMLSnippet(
        id="divider_section",
        label="Section Divider (== Title ==)",
        category="Notes & Lifelines",
        diagram_scope=("Sequence",),
        description="Visual horizontal header dividing call flow phases",
        code="== {cursor}Phase 1: Initial Registration & Security ==",
    ),
    PlantUMLSnippet(
        id="divider_delay",
        label="Time Delay (... Delay ...)",
        category="Notes & Lifelines",
        diagram_scope=("Sequence",),
        description="Indicate elapsed time or timer expiration between messages",
        code="... {cursor}500ms delay / Timer T3510 expires ...",
    ),
    PlantUMLSnippet(
        id="divider_spacing",
        label="Vertical Spacing (||N||)",
        category="Notes & Lifelines",
        diagram_scope=("Sequence",),
        description="Add vertical breathing room between messages",
        code="||45||\n{cursor}",
    ),

    # ---------------------------------------------------------------------
    # 7. MULTI-DIAGRAM ELEMENTS (State, Activity, Timing, Packet, Common)
    # ---------------------------------------------------------------------
    PlantUMLSnippet(
        id="state_trans",
        label="State Transition with Guard/Action",
        category="State Diagram",
        diagram_scope=("State",),
        description="Transition between states with event, guard condition, and action",
        code="{cursor}StateA --> StateB : Event [Guard] / Action",
    ),
    PlantUMLSnippet(
        id="state_composite",
        label="Composite Sub-State",
        category="State Diagram",
        diagram_scope=("State",),
        description="Nested internal sub-states inside a parent state",
        code="""state "{cursor}5GMM-REGISTERED" as RegState {
  [*] --> Initiated
  Initiated --> NormalService : Confirm
  NormalService --> [*]
}""",
    ),
    PlantUMLSnippet(
        id="state_fork_join",
        label="Fork / Join Concurrency",
        category="State Diagram",
        diagram_scope=("State",),
        description="State machine fork and join synchronization bars",
        code="""state fork_state <<fork>>
[*] --> fork_state
fork_state --> State1
fork_state --> State2

state join_state <<join>>
State1 --> join_state
State2 --> join_state
join_state --> [*]""",
    ),
    PlantUMLSnippet(
        id="act_if_else",
        label="If / Then / Else / Endif",
        category="Activity Diagram",
        diagram_scope=("Activity (Beta)",),
        description="Conditional decision branch in activity flow",
        code="""if ({cursor}Check Subscription?) then (yes)
  :Grant Slice Access;
else (no)
  :Reject Registration;
  detach
endif""",
    ),
    PlantUMLSnippet(
        id="act_while",
        label="While Loop (Repeat Action)",
        category="Activity Diagram",
        diagram_scope=("Activity (Beta)",),
        description="Looping activity while condition remains true",
        code="""while ({cursor}Retransmit count < Max?) is (yes)
  :Send PDU;
endwhile (no)
:Declare Radio Link Failure;""",
    ),
    PlantUMLSnippet(
        id="act_fork",
        label="Fork / Fork Again / End Fork",
        category="Activity Diagram",
        diagram_scope=("Activity (Beta)",),
        description="Parallel concurrent activity execution",
        code="""fork
  :{cursor}Authenticate at AUSF;
fork again
  :Fetch Profiles from UDM;
end fork""",
    ),
    PlantUMLSnippet(
        id="timing_signals",
        label="Timing Signals (Robust / Concise / Binary)",
        category="Timing Diagram",
        diagram_scope=("Timing",),
        description="Define clock, signal, and state waveforms",
        code="""robust "{cursor}Radio State" as RS
concise "NAS State" as NS
binary "Clock / Strobe" as CLK

@0
RS is IDLE
NS is DEREGISTERED
CLK is low

@100
RS is TX_ACTIVE
NS is REGISTERING
CLK is high""",
    ),
    PlantUMLSnippet(
        id="packet_fields",
        label="Packetdiag Octet Fields",
        category="Packet Diagram",
        diagram_scope=("Packet (packetdiag)",),
        description="Define protocol message header bit fields",
        code="""packetdiag {
  colwidth = 32
  node_height = 40

  0-7: {cursor}Protocol Discriminator
  8-15: Security Header Type
  16-31: Message Authentication Code (MAC)
  32-39: Sequence Number
  40-63: Payload [colheight = 2]
}""",
    ),

    # ---------------------------------------------------------------------
    # 8. FORMATTING & DIRECTIVES (All Diagrams)
    # ---------------------------------------------------------------------
    PlantUMLSnippet(
        id="dir_title_header",
        label="Title, Header & Footer",
        category="Formatting & Directives",
        diagram_scope=("All",),
        description="Add a top title, right-aligned header, and copyright footer",
        code="""title {cursor}3GPP TS 23.502 Call Flow
header 3GPP Technical Specification
footer Page %page% of %lastpage%""",
    ),
    PlantUMLSnippet(
        id="dir_legend",
        label="Legend Box",
        category="Formatting & Directives",
        diagram_scope=("All",),
        description="Render an explanatory legend box in the bottom right",
        code="""legend right
  <b>Legend:</b>
  * <font color="#1E5C99">■</font> Control Plane (NAS/NGAP)
  * <font color="#15803D">■</font> User Plane (GTP-U)
  * <font color="#D32F2F">■</font> Error / Reject
endlegend""",
    ),
    PlantUMLSnippet(
        id="dir_skinparam_tuning",
        label="Sequence Styling & Layout Directives",
        category="Formatting & Directives",
        diagram_scope=("Sequence", "All"),
        description="Popular enterprise skinparams (message wrapping, box padding, response alignment)",
        code="""skinparam responseMessageBelowArrow true
skinparam maxMessageSize 180
skinparam BoxPadding 15
skinparam ParticipantPadding 10
!pragma teoz true""",
    ),
]


# =========================================================================
# --- SAMPLE 3GPP PROCEDURE SPECIFICATION TEXTS ---
# =========================================================================

SAMPLE_3GPP_PROCEDURES = {
    "5GS Initial Registration (TS 23.502 §4.2.2)": """1. UE sends Registration Request (SUCI, 5G-GUTI, Requested NSSAI) to (R)AN.
2. (R)AN selects AMF based on RAT and NSSAI, and sends N2 Initial UE Message (Registration Request).
3. If identity cannot be resolved, AMF sends Identity Request to UE, and UE responds with Identity Response (SUCI).
4. AMF initiates Authentication: AMF invokes Nausf_UEAuthentication_Authenticate Request to AUSF.
5. AUSF contacts UDM (Nudm_UEAuthentication_Get Request) to retrieve Authentication Vectors (5G-AKA/EAP-AKA').
6. UDM returns vectors to AUSF, which challenges UE via AMF (NAS Authentication Request / Response).
7. Upon successful authentication, AUSF returns confirmation to AMF.
8. AMF initiates NAS Security Mode Command / Complete to establish ciphering and integrity protection.
9. AMF invokes Nudm_SDM_Get to retrieve Access and Mobility Subscription data from UDM.
10. AMF establishes Policy Association with PCF (Npcf_AMPolicyControl_Create).
11. AMF transmits Registration Accept (5G-GUTI, Allowed NSSAI, T3512) to (R)AN, which delivers it to UE.
12. UE sends Registration Complete to AMF.""",

    "5GS PDU Session Establishment (TS 23.502 §4.3.2)": """1. UE sends NAS message (PDU Session Establishment Request) to AMF with PDU Session ID, S-NSSAI, DNN.
2. AMF selects SMF and sends Nsmf_PDUSession_CreateSMContext Request to SMF.
3. SMF retrieves session management subscription data from UDM (Nudm_SDM_Get).
4. SMF responds to AMF with Nsmf_PDUSession_CreateSMContext Response.
5. SMF performs session policy association with PCF (Npcf_SMPolicyControl_Create).
6. SMF selects UPF and establishes N4 Session via N4 Session Establishment Request / Response.
7. SMF sends Namf_Communication_N1N2MessageTransfer (N2 SM Information: CN tunnel info, QoS Profile; N1 SM: PDU Session Establishment Accept) to AMF.
8. AMF sends N2 PDU Session Request to (R)AN.
9. (R)AN performs RRC Reconfiguration with UE to set up Data Radio Bearers (DRB).
10. (R)AN sends N2 PDU Session Response (AN tunnel info) to AMF.
11. AMF forwards AN tunnel info to SMF via Nsmf_PDUSession_UpdateSMContext Request.
12. SMF updates UPF with downlink AN tunnel info via N4 Session Modification Request / Response.
13. SMF responds to AMF with Nsmf_PDUSession_UpdateSMContext Response. User Plane data path is now active.""",

    "Xn-based Handover (TS 38.300 §9.2.3)": """1. Source gNB evaluates UE Measurement Reports and makes handover decision.
2. Source gNB sends Xn Handover Request to Target gNB with UE context and QoS flow details.
3. Target gNB performs Admission Control and allocates radio resources.
4. Target gNB responds to Source gNB with Xn Handover Request Acknowledge.
5. Source gNB sends RRCReconfiguration with sync (Handover Command) to UE.
6. Source gNB delivers SN Status Transfer to Target gNB and starts data forwarding.
7. UE synchronizes to Target gNB and transmits RRCReconfigurationComplete.
8. Target gNB sends Path Switch Request to AMF via NGAP.
9. AMF invokes Nsmf_PDUSession_UpdateSMContext Request to SMF.
10. SMF instructs UPF via N4 Session Modification to switch downlink path to Target gNB.
11. UPF sends End Marker packets to Source gNB and switches path to Target gNB.
12. AMF confirms Path Switch to Target gNB with Path Switch Request Ack.
13. Target gNB sends Xn UE Context Release to Source gNB to free resources.""",

    "5GS Service Request Procedure (TS 23.502 §4.2.3)": """1. UE in CM-IDLE sends RRC Setup Complete with NAS Service Request to (R)AN.
2. (R)AN sends N2 Initial UE Message (Service Request, User Location Info) to AMF.
3. If security not active, AMF performs Authentication and NAS Security Mode procedures.
4. AMF sends Nsmf_PDUSession_UpdateSMContext Request to SMF to activate User Plane connections.
5. SMF initiates N4 Session Modification with UPF to allocate downlink tunnel endpoint.
6. SMF sends N2 SM Information (QoS profiles, UPF Core Tunnel Info) back to AMF.
7. AMF sends N2 Request (Security Context, N2 SM Information) to (R)AN.
8. (R)AN performs RRC Reconfiguration to establish Data Radio Bearers with UE.
9. (R)AN sends N2 Response (AN Tunnel Info) to AMF.
10. AMF sends Nsmf_PDUSession_UpdateSMContext Request to SMF with AN Tunnel Info.
11. SMF modifies UPF with AN Tunnel Info. Uplink and Downlink data path are active."""
}


# =========================================================================
# --- CONFIGURATION FILE PATH RESOLUTION ---
# =========================================================================

def get_default_presets_path() -> Path:
    """Returns path to the base, shipped presets configuration file (tracked in Git)."""
    return core.utils.paths.get_project_root() / "config" / "plantuml_presets.json"


def get_user_presets_path() -> Path:
    """Returns path to local workspace user presets file (strictly untracked in Git via .gitignore)."""
    return core.utils.paths.get_project_root() / "config" / "plantuml_presets_user.json"


def get_app_data_presets_path() -> Path:
    """Returns path to machine-level user presets file in user home app data root."""
    return core.utils.paths.get_app_data_root() / "plantuml_presets_user.json"


# =========================================================================
# --- PERSISTENCE & DYNAMIC HOT-RELOAD LOGIC ---
# =========================================================================

_RELOAD_CALLBACKS: List[Callable[[], None]] = []


def register_snippets_reloaded_callback(callback: Callable[[], None]):
    """Register a callback invoked whenever presets are reloaded or modified."""
    if callback not in _RELOAD_CALLBACKS:
        _RELOAD_CALLBACKS.append(callback)


def unregister_snippets_reloaded_callback(callback: Callable[[], None]):
    """Unregister a reload callback."""
    if callback in _RELOAD_CALLBACKS:
        _RELOAD_CALLBACKS.remove(callback)


def save_snippets_to_file(path: Path, snippets: List[PlantUMLSnippet]) -> None:
    """Writes a list of PlantUMLSnippet objects to a formatted JSON file."""
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        data = {
            "version": 1,
            "snippets": [s.to_dict() for s in snippets],
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception as e:
        logging.error(f"Failed saving PlantUML snippets to {path}: {e}")


def load_snippets_from_file(path: Path, is_user_defined: bool = False) -> List[PlantUMLSnippet]:
    """Reads PlantUMLSnippet objects from a JSON file."""
    if not path.exists():
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            raw_list = data.get("snippets", [])
            return [PlantUMLSnippet.from_dict(d, is_user_defined=is_user_defined) for d in raw_list]
    except Exception as e:
        logging.error(f"Failed reading PlantUML snippets from {path}: {e}")
        return []


def load_user_snippets_from_file(path: Path) -> List[PlantUMLSnippet]:
    """Reads user-defined snippets from a JSON file, flagging them as user defined."""
    return load_snippets_from_file(path, is_user_defined=True)


def save_user_snippets_to_file(path: Path, snippets: List[PlantUMLSnippet]) -> None:
    """Saves user-defined snippets to a JSON file."""
    save_snippets_to_file(path, snippets)


def _load_and_merge_user_file(
    path: Path,
    snippets_by_id: Dict[str, PlantUMLSnippet],
    ordered_list: List[PlantUMLSnippet]
):
    """Loads snippets from path and merges or appends them into ordered_list."""
    if not path.exists():
        return
    user_snips = load_user_snippets_from_file(path)
    for snip in user_snips:
        if snip.id in snippets_by_id:
            # Override existing built-in or previous snippet with user version
            idx = ordered_list.index(snippets_by_id[snip.id])
            ordered_list[idx] = snip
            snippets_by_id[snip.id] = snip
        else:
            snippets_by_id[snip.id] = snip
            ordered_list.append(snip)


def reload_snippets() -> List[PlantUMLSnippet]:
    """Loads default and user presets from disk, updating PLANTUML_SNIPPETS in-place.

    Loading Order:
      1. Default presets (`src/config/plantuml_presets.json`). If missing, creates it
         with all built-in 3GPP and UML presets.
      2. Workspace user presets (`src/config/plantuml_presets_user.json`, Git-ignored).
      3. AppData user presets (`~/3GPP_Delegate_Helper/plantuml_presets_user.json`).
    """
    default_path = get_default_presets_path()
    ordered_list: List[PlantUMLSnippet] = []
    snippets_by_id: Dict[str, PlantUMLSnippet] = {}

    if default_path.exists():
        loaded = load_snippets_from_file(default_path, is_user_defined=False)
        for s in loaded:
            snippets_by_id[s.id] = s
            ordered_list.append(s)

    # If default file is missing or empty, initialize with BUILTIN_SNIPPETS
    if not ordered_list:
        ordered_list = list(BUILTIN_SNIPPETS)
        for s in ordered_list:
            snippets_by_id[s.id] = s
        save_snippets_to_file(default_path, ordered_list)
    else:
        # Guarantee any newly added built-in presets (e.g. part_3gpp_ue) are present
        missing_builtins = [b for b in BUILTIN_SNIPPETS if b.id not in snippets_by_id]
        if missing_builtins:
            for b in missing_builtins:
                snippets_by_id[b.id] = b
                ordered_list.append(b)
            save_snippets_to_file(default_path, ordered_list)

    # Merge Workspace User Presets (plantuml_presets_user.json)
    user_path = get_user_presets_path()
    _load_and_merge_user_file(user_path, snippets_by_id, ordered_list)

    # Merge AppData User Presets
    app_data_path = get_app_data_presets_path()
    if app_data_path != user_path:
        _load_and_merge_user_file(app_data_path, snippets_by_id, ordered_list)

    # Update in-place so all imported references stay current
    PLANTUML_SNIPPETS[:] = ordered_list

    # Notify all registered GUI listeners to refresh their trees / tables / menus
    for cb in list(_RELOAD_CALLBACKS):
        try:
            cb()
        except Exception as e:
            logging.error(f"Error in snippets reload callback: {e}")

    return PLANTUML_SNIPPETS


def add_user_snippet(snippet: PlantUMLSnippet) -> None:
    """Adds or updates a user snippet in plantuml_presets_user.json and reloads."""
    user_path = get_user_presets_path()
    user_snippets = load_user_snippets_from_file(user_path)

    # Guarantee is_user_defined is True
    if not snippet.is_user_defined:
        snippet = PlantUMLSnippet(
            id=snippet.id,
            label=snippet.label,
            category=snippet.category,
            diagram_scope=snippet.diagram_scope,
            description=snippet.description,
            code=snippet.code,
            is_block_wrapper=snippet.is_block_wrapper,
            is_user_defined=True,
        )

    updated = False
    for i, s in enumerate(user_snippets):
        if s.id == snippet.id:
            user_snippets[i] = snippet
            updated = True
            break
    if not updated:
        user_snippets.append(snippet)

    save_user_snippets_to_file(user_path, user_snippets)
    reload_snippets()


def delete_user_snippet(snippet_id: str) -> bool:
    """Removes a user snippet from plantuml_presets_user.json and reloads."""
    user_path = get_user_presets_path()
    user_snippets = load_user_snippets_from_file(user_path)
    count_before = len(user_snippets)
    user_snippets = [s for s in user_snippets if s.id != snippet_id]
    if len(user_snippets) != count_before:
        save_user_snippets_to_file(user_path, user_snippets)
        reload_snippets()
        return True
    return False


# =========================================================================
# --- MODULE INITIALIZATION ---
# =========================================================================

# Global list of snippets imported by external modules:
# `from modules.puml2visio.templates.plantuml_snippets import PLANTUML_SNIPPETS`
PLANTUML_SNIPPETS: List[PlantUMLSnippet] = []

# Populate immediately on module load
reload_snippets()
