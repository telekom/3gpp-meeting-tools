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
    "PacketConfig",
    "PacketField",
    "PacketLayer",
    "PacketScenarioPreset",
    "UP_PRESET_MAP",
    "UP_SCENARIO_PRESETS",
    "build_gtpu_header_layers",
    "build_transport_layers",
    "compose_gtpu_extension_header",
    "compose_up_packet",
    "copy_packet_to_clipboard",
    "get_flags_for_target",
    "load_up_release_capabilities",
    "render_packet_html",
    "render_packet_markdown",
    "render_packet_plain_text",
    "render_packet_plantuml_packetdiag",
    "resolve_up_packet_target",
]
