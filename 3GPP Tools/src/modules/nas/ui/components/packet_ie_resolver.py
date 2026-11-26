"""IE definition lookup candidates and side-panel HTML builder for PacketStackWidget."""

from __future__ import annotations

import html
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from core.ui.theme import C
from modules.nas.core.nas_db import parse_version_tuple
from modules.nas.core.packet_stack.exporters import _format_spec_ref
from modules.nas.core.packet_stack.models import ComposedPacket, PacketField


@dataclass
class FieldDefinitionResult:
    """Resolved DB definition(s) plus fallback wire summary for a clicked packet field."""

    title: str
    spec_number: str
    clause: str
    defs: List[Dict[str, Any]] = field(default_factory=list)
    selected: Optional[Dict[str, Any]] = None
    summary_html: str = ""
    found: bool = False


def field_lookup_candidates(field_name: str, fld: Optional[PacketField] = None) -> List[str]:
    """Build ordered, de-duplicated IE definition lookup candidates for a packet field."""
    candidates: List[str] = []

    def _add(val: str) -> None:
        clean = (val or "").strip()
        if not clean:
            return
        if clean not in candidates:
            candidates.append(clean)
        if "-" in clean:
            spaced = clean.replace("-", " ")
            if spaced not in candidates:
                candidates.append(spaced)
        elif " " in clean and clean.replace(" ", "-") not in candidates:
            hyphened = clean.replace(" ", "-")
            candidates.append(hyphened)

    if fld and fld.def_name:
        _add(fld.def_name)

    raw = (field_name or "").strip()
    if raw.endswith(" Criticality"):
        _add("Criticality")
        return candidates

    # Extract embedded IE name from `IE ID = N (Name)`
    m_ie = re.match(r"^IE\s+ID\s*=\s*\d+\s*\((.+)\)$", raw, re.IGNORECASE)
    if m_ie:
        _add(m_ie.group(1))

    # Strip known wire-envelope suffixes
    stripped = raw
    for suffix in (" Value Length", " Value", " Len"):
        if stripped.endswith(suffix) and len(stripped) > len(suffix):
            stripped = stripped[: -len(suffix)].strip()
            break
    if stripped != raw:
        _add(stripped)

    # Drop trailing parenthetical annotation if any
    no_paren = re.sub(r"\s*\([^)]*\)\s*$", "", stripped).strip()
    if no_paren and no_paren != stripped:
        _add(no_paren)

    if raw:
        _add(raw)
    return candidates


def render_field_summary_html(
    fld: Optional[PacketField],
    spec_number: str,
    clause: str,
    field_name: str,
) -> str:
    """Render compact wire-field metadata card using `C` tokens."""
    display_name = (fld.name if fld else field_name) or field_name or "Field"
    spec_ref = _format_spec_ref(spec_number, clause)
    meta_chips: List[str] = []
    if fld:
        if fld.size_label:
            meta_chips.append(f"<b>Size:</b> {html.escape(fld.size_label)}")
        if fld.presence:
            meta_chips.append(f"<b>Presence:</b> {html.escape(fld.presence)}")
        if fld.value_hint:
            meta_chips.append(f"<b>Value:</b> <code>{html.escape(fld.value_hint)}</code>")
    if spec_ref:
        meta_chips.append(f"<b>Ref:</b> {html.escape(spec_ref)}")

    wire_line = ""
    if fld and fld.short_label and fld.short_label.strip() != display_name.strip():
        wire_line = (
            f'<div style="font-size: 10.5px; color: {C.TEXT_SECONDARY}; margin-top: 2px;">'
            f'<b>Wire label:</b> <code>{html.escape(fld.short_label)}</code></div>'
        )
    cond_line = ""
    if fld and fld.condition:
        cond_line = (
            f'<div style="font-size: 10.5px; color: {C.TEXT_SECONDARY}; margin-top: 2px;">'
            f'<b>Condition:</b> {html.escape(fld.condition)}</div>'
        )
    desc_line = ""
    if fld and fld.description:
        desc_line = (
            f'<div style="font-size: 11px; color: {C.TEXT_BODY}; margin-top: 4px; line-height: 1.4;">'
            f'{html.escape(fld.description)}</div>'
        )

    chips_html = (
        f'<div style="font-size: 10.5px; color: {C.TEXT_SECONDARY}; margin-top: 3px;">'
        + " &nbsp;·&nbsp; ".join(meta_chips)
        + "</div>"
        if meta_chips
        else ""
    )

    return (
        f'<div style="background-color: {C.SKY_SURFACE}; border: 1px solid {C.SKY_BORDER}; '
        f'border-left: 4px solid {C.ACCENT_SKY}; border-radius: 5px; padding: 8px 10px; margin-bottom: 10px;">'
        f'<div style="font-size: 12.5px; font-weight: bold; color: {C.ACCENT_SKY_DARK};">'
        f'{html.escape(display_name)}</div>'
        f'{chips_html}{wire_line}{cond_line}{desc_line}'
        f'</div>'
    )


def build_side_panel_html(result: FieldDefinitionResult) -> str:
    """Compose the full HTML document shown inside `PacketIESidePanel`."""
    parts: List[str] = [
        f'<div style="font-family: Segoe UI, Calibri, Arial, sans-serif; color: {C.TEXT_PRIMARY};">',
        result.summary_html,
    ]
    sel = result.selected
    if result.found and sel and (sel.get("raw_description") or "").strip():
        spec = str(sel.get("spec_number") or result.spec_number or "").strip()
        ver = str(sel.get("version") or "").strip()
        cl = str(sel.get("clause") or result.clause or "").strip()
        hdr_bits = []
        if spec:
            hdr_bits.append(f"TS {spec}" + (f" v{ver}" if ver else ""))
        if cl and cl != "-":
            hdr_bits.append(f"Clause {cl}")
        hdr_text = " · ".join(hdr_bits)
        parts.append(
            f'<div style="font-size: 11px; font-weight: bold; color: {C.TEXT_SECONDARY}; '
            f'border-bottom: 1px solid {C.BORDER_SUBTLE}; padding-bottom: 3px; margin-bottom: 6px;">'
            f'{html.escape(hdr_text)}</div>'
        )
        parts.append(f'<div>{sel["raw_description"]}</div>')
    else:
        spec = (result.spec_number or "").strip()
        if spec and not spec[0].isdigit():
            note = (
                f'<b>{html.escape(result.title)}</b> is defined in external specification '
                f'<b>{html.escape(spec)}</b> (not indexed in the local 3GPP protocol database) — showing wire-field summary.'
            )
        elif spec:
            note = (
                f'No Clause definition for <b>"{html.escape(result.title)}"</b> found in the local '
                f'<b>TS {html.escape(spec)}</b> database — showing wire-field summary.'
            )
        else:
            note = "Showing wire-field summary."
        parts.append(
            f'<div style="background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; '
            f'border-radius: 5px; padding: 8px 10px; font-size: 11px; color: {C.TEXT_SECONDARY}; line-height: 1.45;">'
            f'ℹ️ {note}</div>'
        )
    parts.append("</div>")
    return "".join(parts)


def resolve_packet_field(
    db: Any,
    packet: Optional[ComposedPacket],
    spec_number: str,
    clause: str,
    field_name: str,
    version_hint: str = "",
) -> FieldDefinitionResult:
    """Resolve DB definition(s) for a clicked packet field, with graceful wire-summary fallback."""
    clean_spec = (spec_number or "").strip()
    clean_clause = (clause or "").strip()
    clean_name = (field_name or "").strip()

    fld = packet.find_field(clean_name) if packet and clean_name else None
    if not clean_name and fld:
        clean_name = fld.name
    if fld and not clean_spec:
        clean_spec = fld.spec_number
    if fld and (not clean_clause or clean_clause == "-"):
        clean_clause = fld.clause

    candidates = field_lookup_candidates(clean_name or clean_clause, fld)
    defs: List[Dict[str, Any]] = []

    if db is not None and hasattr(db, "get_ie_definitions_by_clause") and clean_spec[:1].isdigit():
        try:
            for cand in candidates:
                hit = db.get_ie_definitions_by_clause(
                    clause="",
                    alt_name=cand,
                    spec_number=clean_spec,
                )
                if hit:
                    defs = hit
                    break
            if not defs and clean_clause and clean_clause != "-":
                defs = db.get_ie_definitions_by_clause(
                    clause=clean_clause,
                    alt_name=candidates[0] if candidates else "",
                    spec_number=clean_spec,
                )
        except Exception:
            defs = []

    if defs:
        defs = sorted(
            defs,
            key=lambda d: parse_version_tuple(str(d.get("version") or "")),
            reverse=True,
        )

    selected: Optional[Dict[str, Any]] = None
    if defs:
        v_clean = (version_hint or "").lstrip("vV").strip()
        if v_clean:
            for d in defs:
                if str(d.get("version") or "").lstrip("vV").strip() == v_clean:
                    selected = d
                    break
        if selected is None:
            selected = defs[0]

    found = bool(selected)
    resolved_clause = str((selected or {}).get("clause") or clean_clause or "").strip()
    if found and selected and selected.get("ie_name"):
        title = str(selected["ie_name"]).strip()
    elif candidates:
        title = candidates[0]
    else:
        title = clean_name or (f"Clause {clean_clause}" if clean_clause else "Field Details")

    summary_html = render_field_summary_html(fld, clean_spec, resolved_clause, clean_name or title)
    return FieldDefinitionResult(
        title=title,
        spec_number=clean_spec,
        clause=resolved_clause,
        defs=defs,
        selected=selected,
        summary_html=summary_html,
        found=found,
    )
