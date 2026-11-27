"""Data models for cross-spec protocol packet stack composition and rendering."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass
class PacketField:
    """Single bit- or octet-aligned field within a protocol layer."""

    name: str
    short_label: str
    bits: int
    presence: str = "M"
    value_hint: str = ""
    spec_number: str = ""
    clause: str = ""
    condition: str = ""
    flag_key: str = ""
    description: str = ""
    is_padding: bool = False
    is_spare: bool = False
    is_variable: bool = False
    variable_display_octets: int = 0
    def_name: str = ""

    @property
    def is_octet_aligned(self) -> bool:
        return self.bits > 0 and self.bits % 8 == 0

    @property
    def octets(self) -> int:
        if self.bits <= 0:
            return 0
        return (self.bits + 7) // 8

    @property
    def size_label(self) -> str:
        if self.is_variable:
            if self.variable_display_octets > 0:
                return f"Var (~{self.variable_display_octets} B)"
            return "Variable"
        if self.bits < 8:
            return "1 bit" if self.bits == 1 else f"{self.bits} bits"
        if self.bits % 8 == 0:
            n_oct = self.bits // 8
            return "1 octet" if n_oct == 1 else f"{n_oct} octets"
        return f"{self.bits} bits"


@dataclass
class PacketLayer:
    """Contiguous protocol layer or extension header wrapper inside a composed packet."""

    layer_id: str
    title: str
    spec_number: str
    clause: str
    role: str
    fields: List[PacketField] = field(default_factory=list)
    collapsible: bool = False
    collapsed: bool = False
    notes: List[str] = field(default_factory=list)

    @property
    def fixed_bits(self) -> int:
        return sum(f.bits for f in self.fields if not f.is_variable)

    @property
    def fixed_octets(self) -> int:
        return (self.fixed_bits + 7) // 8

    @property
    def has_variable_fields(self) -> bool:
        return any(f.is_variable for f in self.fields)


@dataclass(frozen=True)
class FlagOption:
    """Interactive option/flag pill exposed in the Packet Stack toolbar."""

    key: str
    label: str
    tooltip: str
    group: str
    default_enabled: bool = False
    requires_flag: str = ""
    is_release_present: bool = True


@dataclass(frozen=True)
class PacketScenarioPreset:
    """Pre-configured User Plane interface/scenario profile."""

    preset_id: str
    label: str
    interface: str
    description: str
    default_flags: Dict[str, bool] = field(default_factory=dict)
    transport_mode: str = "collapsed"
    include_tpdu: bool = True


@dataclass
class PacketConfig:
    """User-editable configuration driving `compose_up_packet()` / `compose_ngap_packet()`."""

    target_frame: str = "dl_pdu_session"
    preset_id: str = "n3_n9"
    transport_mode: str = "collapsed"  # "collapsed", "ipv4_udp"/"ipv4_sctp", "ipv6_udp"/"ipv6_sctp", "sctp_only", "none"
    flags: Dict[str, bool] = field(default_factory=dict)
    include_tpdu: bool = True
    tpdu_label: str = "T-PDU Payload (IPv4 / IPv6 / Ethernet)"
    spec_version_38415: str = ""
    spec_version_29281: str = ""
    highlight_field: str = ""
    stack_family: str = "up"
    message_name: str = ""
    nas_message_name: str = ""
    spec_version_38413: str = ""
    spec_version_24501: str = ""
    spec_version_policy: str = ""
    payload_container_type: str = ""
    policy_part_type: str = "ursp"


@dataclass
class ComposedPacket:
    """Fully resolved, octet-indexed packet stack ready for HTML/Markdown/PlantUML export."""

    title: str
    subtitle: str
    direction: str
    interface: str
    spec_refs: List[str]
    layers: List[PacketLayer]
    available_flags: List[FlagOption]
    available_presets: List[PacketScenarioPreset]
    config: PacketConfig
    overhead_3gpp_octets: int
    overhead_transport_octets: int
    overhead_total_octets: int
    ext_header_summaries: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    def all_fields(self) -> List[PacketField]:
        out: List[PacketField] = []
        for layer in self.layers:
            out.extend(layer.fields)
        return out

    def find_field(self, field_name: str) -> Optional[PacketField]:
        target = (field_name or "").strip().lower()
        if not target:
            return None
        for layer in self.layers:
            for fld in layer.fields:
                if fld.name.strip().lower() == target or fld.short_label.strip().lower() == target:
                    return fld
        return None
