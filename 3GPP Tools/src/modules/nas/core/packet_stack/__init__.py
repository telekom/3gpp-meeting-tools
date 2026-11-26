"""Cross-specification protocol packet stack composition and exporters."""

from __future__ import annotations

from .composer import (
    build_gtpu_header_layers,
    build_transport_layers,
    compose_gtpu_extension_header,
    compose_up_packet,
)
from .exporters import (
    copy_packet_to_clipboard,
    render_packet_html,
    render_packet_markdown,
    render_packet_plain_text,
    render_packet_plantuml_packetdiag,
)
from .models import (
    ComposedPacket,
    FlagOption,
    PacketConfig,
    PacketField,
    PacketLayer,
    PacketScenarioPreset,
)
from .ngap_catalog import (
    NGAP_PRESET_MAP,
    NGAP_SCENARIO_PRESETS,
    SAMPLE_NAS_MESSAGES,
    default_nas_message_for_ngap,
    is_non_ue_ngap_message,
    load_nas_message_fields_from_db,
    load_ngap_message_from_db,
    resolve_ngap_packet_target,
)
from .ngap_composer import (
    build_ngap_pdu_layers,
    build_sctp_transport_layers,
    compose_ngap_packet,
)
from .registry import (
    compose_packet,
    get_preset_map_for_family,
    get_presets_for_family,
    resolve_packet_target,
)
from .up_catalog import (
    UP_PRESET_MAP,
    UP_SCENARIO_PRESETS,
    get_flags_for_target,
    load_up_release_capabilities,
    resolve_up_packet_target,
)

__all__ = [
    "ComposedPacket",
    "FlagOption",
    "NGAP_PRESET_MAP",
    "NGAP_SCENARIO_PRESETS",
    "PacketConfig",
    "PacketField",
    "PacketLayer",
    "PacketScenarioPreset",
    "SAMPLE_NAS_MESSAGES",
    "UP_PRESET_MAP",
    "UP_SCENARIO_PRESETS",
    "build_gtpu_header_layers",
    "build_ngap_pdu_layers",
    "build_sctp_transport_layers",
    "build_transport_layers",
    "compose_gtpu_extension_header",
    "compose_ngap_packet",
    "compose_packet",
    "compose_up_packet",
    "copy_packet_to_clipboard",
    "default_nas_message_for_ngap",
    "get_flags_for_target",
    "get_preset_map_for_family",
    "get_presets_for_family",
    "is_non_ue_ngap_message",
    "load_nas_message_fields_from_db",
    "load_ngap_message_from_db",
    "load_up_release_capabilities",
    "render_packet_html",
    "render_packet_markdown",
    "render_packet_plain_text",
    "render_packet_plantuml_packetdiag",
    "resolve_ngap_packet_target",
    "resolve_packet_target",
    "resolve_up_packet_target",
]
