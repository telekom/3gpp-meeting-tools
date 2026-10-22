# --- File: src/modules/puml2visio/templates/plantuml_snippets.py ---
"""Declarative catalog of PlantUML syntax elements, snippets, and 3GPP presets

with dynamic placeholder resolution (participants, selection wrapping, cursor placement).
"""

from dataclasses import dataclass
import re
from typing import List, Optional, Tuple


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
        # Ignore common keywords that might be accidentally caught
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
            # Indent each line of the selection by 2 spaces relative to block
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
            # First line is placed at cursor (already at base_indent)
            # Subsequent lines get base_indent prefixed
            indented = [lines[0]] + [base_indent + l if l.strip() else l for l in lines[1:]]
            resolved = "\n".join(indented)

    # Locate cursor offset
    cursor_marker = "{cursor}"
    if cursor_marker in resolved:
        start_idx = resolved.find(cursor_marker)
        resolved = resolved.replace(cursor_marker, "", 1)
        # Check if there is an end marker or selected token
        end_marker = "{/cursor}"
        if end_marker in resolved:
            end_idx = resolved.find(end_marker)
            resolved = resolved.replace(end_marker, "", 1)
            return resolved, start_idx, end_idx
        return resolved, start_idx, start_idx

    return resolved, len(resolved), len(resolved)


# =========================================================================
# --- COMPREHENSIVE PLANTUML SNIPPET CATALOG ---
# =========================================================================

PLANTUML_SNIPPETS: List[PlantUMLSnippet] = [
    # ---------------------------------------------------------------------
    # 1. 3GPP ARCHITECTURE PRESETS (Sequence)
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
    # 2. PARTICIPANTS & BOXES (Sequence)
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
    # 3. MESSAGES & ARROWS (Sequence)
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
    # 4. LOGIC & GROUPING FRAGMENTS (Sequence)
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
    # 5. NOTES, LIFELINES & STATE (Sequence)
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
    # 6. MULTI-DIAGRAM ELEMENTS (State, Activity, Timing, Packet, Common)
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
    # 7. FORMATTING & DIRECTIVES (All Diagrams)
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
