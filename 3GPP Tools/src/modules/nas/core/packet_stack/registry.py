"""Unified multi-stack dispatcher (5G User Plane, N2 NGAP Control Plane, and 5GS NAS/Policies)."""

from __future__ import annotations

from typing import Any, Dict, List, Optional

from .composer import compose_up_packet
from .models import ComposedPacket, PacketConfig, PacketScenarioPreset
from .nas_catalog import NAS_PRESET_MAP, NAS_SCENARIO_PRESETS, resolve_nas_packet_target
from .nas_composer import compose_nas_packet
from .ngap_catalog import NGAP_PRESET_MAP, NGAP_SCENARIO_PRESETS, resolve_ngap_packet_target
from .ngap_composer import compose_ngap_packet
from .up_catalog import UP_PRESET_MAP, UP_SCENARIO_PRESETS, resolve_up_packet_target


def get_presets_for_family(stack_family: str) -> List[PacketScenarioPreset]:
    if stack_family == "ngap":
        return NGAP_SCENARIO_PRESETS
    if stack_family == "nas":
        return NAS_SCENARIO_PRESETS
    return UP_SCENARIO_PRESETS


def get_preset_map_for_family(stack_family: str) -> Dict[str, PacketScenarioPreset]:
    if stack_family == "ngap":
        return NGAP_PRESET_MAP
    if stack_family == "nas":
        return NAS_PRESET_MAP
    return UP_PRESET_MAP


def resolve_packet_target(
    spec_number: Optional[str],
    clause: str = "",
    ie_name: str = "",
    message_name: str = "",
) -> Optional[Dict[str, Any]]:
    """Resolve any supported specification selection into a UP, NGAP, or 5GS NAS packet target."""
    up_target = resolve_up_packet_target(
        spec_number=spec_number,
        clause=clause,
        ie_name=ie_name,
        message_name=message_name,
    )
    if up_target is not None:
        up_target.setdefault("stack_family", "up")
        return up_target

    ngap_target = resolve_ngap_packet_target(
        spec_number=spec_number,
        clause=clause,
        ie_name=ie_name,
        message_name=message_name,
    )
    if ngap_target is not None:
        return ngap_target

    return resolve_nas_packet_target(
        spec_number=spec_number,
        clause=clause,
        ie_name=ie_name,
        message_name=message_name,
    )


def compose_packet(config: PacketConfig, db: Any = None) -> ComposedPacket:
    """Compose a 5G UP, N2 NGAP, or 5GS NAS packet stack based on `config.stack_family`."""
    if config.stack_family == "ngap":
        return compose_ngap_packet(config, db=db)
    if config.stack_family == "nas":
        return compose_nas_packet(config, db=db)
    return compose_up_packet(config, db=db)
