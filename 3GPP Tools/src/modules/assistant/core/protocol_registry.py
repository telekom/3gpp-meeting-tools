"""Static protocol capability registry.

The registry describes parser/application capability.  It does *not* claim that
structured data for a protocol is currently populated in the protocol DB.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Iterable, Optional, Tuple


@dataclass(frozen=True)
class ProtocolDescriptor:
    id: str
    display_name: str
    specifications: Tuple[str, ...]
    aliases: Tuple[str, ...] = ()
    reference_points: Tuple[str, ...] = ()
    parser_supported: bool = True
    notes: str = ""


_PROTOCOLS: Tuple[ProtocolDescriptor, ...] = (
    ProtocolDescriptor(
        id="pfcp",
        display_name="PFCP",
        specifications=("29.244",),
        aliases=("packet forwarding control protocol",),
        reference_points=("N4",),
    ),
    ProtocolDescriptor(
        id="gtpu",
        display_name="GTP-U",
        specifications=("29.281",),
        aliases=("gtp-u", "gtpu", "gtpv1-u"),
    ),
    ProtocolDescriptor(
        id="nr_rrc",
        display_name="NR RRC",
        specifications=("38.331",),
        aliases=("nr rrc",),
    ),
    ProtocolDescriptor(
        id="lte_rrc",
        display_name="LTE RRC",
        specifications=("36.331",),
        aliases=("e-utra rrc", "lte rrc"),
    ),
    ProtocolDescriptor(
        id="ngap",
        display_name="NGAP",
        specifications=("38.413",),
        aliases=("ng application protocol",),
        reference_points=("N2",),
    ),
    ProtocolDescriptor(
        id="xnap",
        display_name="XnAP",
        specifications=("38.423",),
        aliases=("xn application protocol",),
        reference_points=("Xn",),
    ),
    ProtocolDescriptor(
        id="f1ap",
        display_name="F1AP",
        specifications=("38.473",),
        aliases=("f1 application protocol",),
        reference_points=("F1",),
    ),
    ProtocolDescriptor(
        id="e1ap",
        display_name="E1AP",
        specifications=("38.463",),
        aliases=("e1 application protocol",),
        reference_points=("E1",),
    ),
    ProtocolDescriptor(
        id="pdu_session_up",
        display_name="PDU Session UP",
        specifications=("38.415",),
        aliases=("pdu session user plane", "pdu set user plane"),
    ),
    ProtocolDescriptor(
        id="5gs_nas",
        display_name="5GS NAS",
        specifications=("24.501",),
        aliases=("5g nas", "5gs nas"),
        notes="TS 24.501 contains both 5GMM and 5GSM. Phase 1 keeps the registry at specification level.",
    ),
    ProtocolDescriptor(
        id="eps_nas",
        display_name="EPS NAS",
        specifications=("24.301",),
        aliases=("eps nas", "lte nas"),
    ),
    ProtocolDescriptor(
        id="core_network_nas",
        display_name="Core Network NAS",
        specifications=("24.008",),
        aliases=("24.008 nas",),
    ),
)


def iter_protocols() -> Iterable[ProtocolDescriptor]:
    return _PROTOCOLS


def get_protocol(protocol_id: str) -> Optional[ProtocolDescriptor]:
    key = _normalize(protocol_id)
    for descriptor in _PROTOCOLS:
        if _normalize(descriptor.id) == key:
            return descriptor
    return None


def resolve_protocol(value: str) -> Tuple[ProtocolDescriptor, ...]:
    """Resolve a protocol name/alias, without treating interfaces as aliases."""
    key = _normalize(value)
    if not key:
        return ()

    exact = []
    partial = []
    for descriptor in _PROTOCOLS:
        names = (descriptor.id, descriptor.display_name, *descriptor.aliases)
        normalized = {_normalize(name) for name in names if name}
        if key in normalized:
            exact.append(descriptor)
        elif any(key in candidate or candidate in key for candidate in normalized):
            partial.append(descriptor)

    return tuple(exact or partial)


def find_by_specification(specification: str) -> Tuple[ProtocolDescriptor, ...]:
    spec = _normalize_spec(specification)
    return tuple(p for p in _PROTOCOLS if spec in p.specifications)


def registry_by_id() -> Dict[str, ProtocolDescriptor]:
    return {item.id: item for item in _PROTOCOLS}


def _normalize(value: str) -> str:
    return "".join(ch.lower() for ch in str(value or "") if ch.isalnum())


def _normalize_spec(value: str) -> str:
    text = str(value or "").strip().upper()
    for prefix in ("3GPP", "TS", "TR"):
        text = text.replace(prefix, "")
    return text.strip()
