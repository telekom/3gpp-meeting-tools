"""Multi-format renderers and clipboard exporter (HTML, Word, Markdown, Plain Text, PlantUML packetdiag)."""

from __future__ import annotations

import html
from typing import Dict, List, Tuple
from urllib.parse import quote

from PyQt5.QtCore import QMimeData
from PyQt5.QtWidgets import QApplication

from core.ui.theme import C
from modules.nas.core.word_html import compact_html_for_word
from .models import ComposedPacket, PacketField, PacketLayer


_ROLE_COLORS: Dict[str, Tuple[str, str, str]] = {
    # (banner_bg, banner_fg, cell_tint)
    "transport": (C.BG_MUTED, C.TEXT_STRONG, C.BG_SURFACE),
    "sctp_chunk": (C.TEAL_SOFT, C.TEAL_DARK, C.TEAL_SURFACE),
    "gtpu_header": (C.SKY_SOFT, C.ACCENT_SKY_DARK, C.SKY_SURFACE),
    "gtpu_ext_wrapper": (C.SUCCESS_SOFT_ALT, C.SUCCESS_DARK, C.SUCCESS_SOFT),
    "up_frame": (C.WARNING_YELLOW_SOFT, C.WARNING_YELLOW_TEXT, C.WARNING_SOFT),
    "ngap_envelope": (C.SKY_SOFT, C.ACCENT_SKY_DARK, C.SKY_SURFACE),
    "ngap_ie": (C.INFO_SOFT, C.INFO_DARK, C.BLUE_TINT_SOFT),
    "nas_security": (C.ORANGE_SOFT, C.ORANGE_DARK, C.WARNING_SOFT),
    "nas_message": (C.WARNING_YELLOW_SOFT, C.WARNING_YELLOW_TEXT, C.WARNING_SOFT),
    "ue_policy_delivery": (C.TEAL_SOFT, C.TEAL_DARK, C.TEAL_SURFACE),
    "ue_policy_rule": (C.VIOLET_SOFT, C.VIOLET_DARK, C.VIOLET_SURFACE),
    "uu_l2": (C.INFO_SOFT, C.INFO_DARK, C.BLUE_TINT_SOFT),
    "uu_sdap": (C.SUCCESS_SOFT_ALT, C.SUCCESS_DARK, C.SUCCESS_SOFT),
    "uu_pdcp": (C.SKY_SOFT, C.ACCENT_SKY_DARK, C.SKY_SURFACE),
    "uu_rlc": (C.TEAL_SOFT, C.TEAL_DARK, C.TEAL_SURFACE),
    "uu_mac": (C.ORANGE_SOFT, C.ORANGE_DARK, C.WARNING_SOFT),
    "segmentation_divider": (C.WARNING_YELLOW_SOFT, C.WARNING_DARK, C.WARNING_SOFT),
    "ran_container": (C.VIOLET_SOFT, C.VIOLET_DARK, C.VIOLET_SURFACE),
    "payload": (C.BG_CANVAS, C.TEXT_SECONDARY, C.BG_ALT_LIGHT),
}


def _family_labels(stack_family: str) -> Tuple[str, str, str]:
    """Return `(overhead_label, html_chain_label, md_chain_label)` for the stack family."""
    if stack_family == "ngap":
        return ("N2 / NGAP", "NGAP Summary", "NGAP Summary")
    if stack_family == "nas":
        return ("5GS NAS", "Encapsulation Chain", "Encapsulation Chain")
    if stack_family == "ran_uu":
        return ("NR Uu", "Layer Stack", "Layer Stack")
    return ("3GPP UP", "Ext Header Chain", "Extension Header Chain")


def _is_3gpp_spec(spec_number: str) -> bool:
    s = (spec_number or "").strip()
    return bool(s) and s[0].isdigit()


def _format_spec_ref(spec_number: str, clause: str) -> str:
    spec = (spec_number or "").strip()
    cl = (clause or "").strip()
    if not spec:
        return cl
    if _is_3gpp_spec(spec):
        return f"TS {spec} {cl}".strip() if cl and cl != "-" else f"TS {spec}"
    return f"{spec} §{cl}".strip() if cl and cl != "-" else spec


def _pack_layer_rows(layer: PacketLayer, start_octet: int) -> Tuple[List[Dict[str, object]], int]:
    """Pack bit/octet fields of a layer into visual 8-bit rows with continuous wire octet offsets."""
    rows: List[Dict[str, object]] = []
    cur_octet = start_octet
    bit_buf: List[PacketField] = []
    bit_sum = 0

    def _flush_bit_buf() -> None:
        nonlocal cur_octet, bit_buf, bit_sum
        if not bit_buf:
            return
        rows.append({
            "kind": "bits",
            "octet_label": str(cur_octet),
            "fields": list(bit_buf),
            "size_label": "1 octet",
            "spec_ref": _format_spec_ref(bit_buf[0].spec_number, bit_buf[0].clause),
        })
        cur_octet += 1
        bit_buf = []
        bit_sum = 0

    for fld in layer.fields:
        if fld.is_variable:
            _flush_bit_buf()
            rows.append({
                "kind": "full",
                "octet_label": "Var",
                "field": fld,
                "size_label": fld.size_label,
                "spec_ref": _format_spec_ref(fld.spec_number, fld.clause),
            })
            continue

        if fld.bits <= 0:
            _flush_bit_buf()
            rows.append({
                "kind": "full",
                "octet_label": "—",
                "field": fld,
                "size_label": fld.size_label,
                "spec_ref": _format_spec_ref(fld.spec_number, fld.clause),
            })
            continue

        if fld.bits <= 8 and (bit_sum > 0 or fld.bits < 8):
            if bit_sum + fld.bits > 8:
                _flush_bit_buf()
            if fld.bits == 8 and bit_sum == 0:
                rows.append({
                    "kind": "full",
                    "octet_label": str(cur_octet),
                    "field": fld,
                    "size_label": fld.size_label,
                    "spec_ref": _format_spec_ref(fld.spec_number, fld.clause),
                })
                cur_octet += 1
                continue
            bit_buf.append(fld)
            bit_sum += fld.bits
            if bit_sum == 8:
                _flush_bit_buf()
            continue

        _flush_bit_buf()
        n_oct = max(1, fld.octets)
        oct_label = str(cur_octet) if n_oct == 1 else f"{cur_octet}–{cur_octet + n_oct - 1}"
        rows.append({
            "kind": "full",
            "octet_label": oct_label,
            "field": fld,
            "size_label": fld.size_label,
            "spec_ref": _format_spec_ref(fld.spec_number, fld.clause),
        })
        cur_octet += n_oct

    _flush_bit_buf()
    return rows, cur_octet


def _is_highlighted(fld: PacketField, highlight_name: str) -> bool:
    if not highlight_name:
        return False
    h = highlight_name.strip().lower()
    return h in fld.name.strip().lower() or fld.name.strip().lower() in h


def render_packet_html(packet: ComposedPacket, interactive_links: bool = True) -> str:
    """Render unified octet/bit stack table in HTML (`C` tokens only; interactive links optional)."""
    hl = packet.config.highlight_field
    ov_lbl, chain_lbl, _md_chain_lbl = _family_labels(packet.config.stack_family)
    parts: List[str] = [
        f'<div style="font-family: Segoe UI, Calibri, Arial, sans-serif; color: {C.TEXT_PRIMARY};">'
    ]

    # Summary Header Banner
    parts.append(
        f'<div style="background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_STRONG}; '
        f'border-left: 4px solid {C.ACCENT_SKY}; border-radius: 5px; padding: 8px 10px; margin-bottom: 8px;">'
        f'<div style="font-size: 13px; font-weight: bold; color: {C.ACCENT_SKY_DARK};">{html.escape(packet.title)}</div>'
        f'<div style="font-size: 11px; color: {C.TEXT_BODY}; margin-top: 2px;">'
        f'<b>Direction:</b> {html.escape(packet.direction)} &nbsp;|&nbsp; '
        f'<b>Interface:</b> {html.escape(packet.interface)} &nbsp;|&nbsp; '
        f'<b>Specs:</b> {html.escape(", ".join(packet.spec_refs))}</div>'
    )
    if packet.channel_mapping_chain:
        pill_style = (
            f"background-color: {C.SKY_SOFT}; color: {C.ACCENT_SKY_DARK}; "
            f"border: 1px solid {C.SKY_BORDER}; border-radius: 3px; "
            f"padding: 1px 5px; font-weight: 600;"
        )
        map_pills = " &nbsp;→&nbsp; ".join(
            f'<span style="{pill_style}">{html.escape(step)}</span>'
            for step in packet.channel_mapping_chain
        )
        parts.append(
            f'<div style="font-size: 10.5px; color: {C.TEXT_BODY}; margin-top: 4px;">'
            f'<b>Channel Mapping:</b> {map_pills}</div>'
        )
    if packet.ext_header_summaries:
        chain_txt = " &nbsp;→&nbsp; ".join(html.escape(s) for s in packet.ext_header_summaries)
        parts.append(
            f'<div style="font-size: 10.5px; color: {C.SUCCESS_DARK}; margin-top: 4px;">'
            f'<b>{chain_lbl}:</b> {chain_txt}</div>'
        )
    parts.append(
        f'<div style="font-size: 10.5px; color: {C.TEXT_SECONDARY}; margin-top: 3px;">'
        f'<b>Fixed Overhead:</b> {ov_lbl} = <b>{packet.overhead_3gpp_octets} octets</b>'
        + (
            f' &nbsp;|&nbsp; Transport = <b>{packet.overhead_transport_octets} octets</b>'
            f' &nbsp;|&nbsp; Total Header = <b>{packet.overhead_total_octets} octets</b>'
            if packet.overhead_transport_octets > 0
            else ""
        )
        + "</div></div>"
    )

    # Bit-Grid Table (HTML width="100%" required so Qt QTextDocument expands cleanly to available viewport width)
    parts.append(
        f'<table width="100%" border="1" cellspacing="0" cellpadding="4" style="border-collapse: collapse; '
        f'width: 100%; border: 1px solid {C.BORDER_STRONG}; font-family: Consolas, \'Courier New\', monospace; font-size: 11px;">'
        f'<thead>'
        f'<tr style="background-color: {C.BG_CANVAS}; color: {C.TEXT_PRIMARY};">'
        f'<th rowspan="2" width="7%" style="border: 1px solid {C.BORDER_STRONG}; padding: 4px 8px; text-align: center; white-space: nowrap; width: 7%;">Octet</th>'
        f'<th colspan="8" style="border: 1px solid {C.BORDER_STRONG}; padding: 3px 6px; text-align: center;">Bits</th>'
        f'<th rowspan="2" width="9%" style="border: 1px solid {C.BORDER_STRONG}; padding: 4px 8px; text-align: center; white-space: nowrap; width: 9%;">Size</th>'
        f'<th rowspan="2" width="16%" style="border: 1px solid {C.BORDER_STRONG}; padding: 4px 8px; text-align: left; white-space: nowrap; width: 16%;">Spec / Clause</th>'
        f'</tr>'
        f'<tr style="background-color: {C.BG_SURFACE}; color: {C.TEXT_SECONDARY};">'
        + "".join(
            f'<th width="8.5%" style="border: 1px solid {C.BORDER_STRONG}; padding: 2px 4px; text-align: center; width: 8.5%;">{b}</th>'
            for b in range(7, -1, -1)
        )
        + "</tr></thead><tbody>"
    )

    cur_octet = 1
    for layer in packet.layers:
        b_bg, b_fg, cell_tint = _ROLE_COLORS.get(layer.role, (C.BG_CANVAS, C.TEXT_STRONG, C.BG_CARD))
        spec_tag = _format_spec_ref(layer.spec_number, f"Cl. {layer.clause}" if layer.clause and layer.clause != "-" else "")
        parts.append(
            f'<tr style="background-color: {b_bg}; color: {b_fg}; font-family: Segoe UI, Arial, sans-serif;">'
            f'<td colspan="11" style="border: 1px solid {C.BORDER_STRONG}; padding: 4px 8px; font-weight: bold; font-size: 11px;">'
            f'{html.escape(layer.title)}'
            + (f' <span style="font-weight: normal; font-size: 10px;">({html.escape(spec_tag)})</span>' if spec_tag else "")
            + "</td></tr>"
        )

        rows, cur_octet = _pack_layer_rows(layer, cur_octet)
        for row in rows:
            parts.append("<tr>")
            parts.append(
                f'<td style="border: 1px solid {C.BORDER_STRONG}; padding: 4px 8px; text-align: center; '
                f'background-color: {C.BG_SURFACE}; color: {C.TEXT_SECONDARY}; font-weight: bold; white-space: nowrap;">'
                f'{html.escape(str(row["octet_label"]))}</td>'
            )
            if row["kind"] == "bits":
                used_bits = 0
                for fld in row["fields"]:  # type: ignore[attr-defined]
                    used_bits += fld.bits
                    bg = C.HIGHLIGHT_YELLOW if _is_highlighted(fld, hl) else (C.BG_MUTED if fld.is_spare else cell_tint)
                    fg = C.TEXT_PRIMARY if not fld.is_spare else C.TEXT_MUTED
                    lbl = html.escape(fld.short_label)
                    if interactive_links and not fld.is_spare:
                        href = (
                            f"action:packet_field:{quote(fld.spec_number, safe='')}:"
                            f"{quote(fld.clause, safe='')}:{quote(fld.name, safe='')}"
                        )
                        lbl_html = f'<a href="{href}" style="color: {fg}; text-decoration: none; font-weight: 600;">{lbl}</a>'
                    else:
                        lbl_html = f'<span style="font-weight: 600;">{lbl}</span>'
                    nowrap = " white-space: nowrap;" if len(fld.short_label) <= 8 else ""
                    parts.append(
                        f'<td colspan="{fld.bits}" style="border: 1px solid {C.BORDER_STRONG}; padding: 4px 3px; '
                        f'text-align: center; background-color: {bg}; color: {fg};{nowrap}">{lbl_html}</td>'
                    )
                if used_bits < 8:
                    parts.append(
                        f'<td colspan="{8 - used_bits}" style="border: 1px solid {C.BORDER_STRONG}; padding: 4px 3px; '
                        f'text-align: center; background-color: {C.BG_MUTED}; color: {C.TEXT_MUTED};">—</td>'
                    )
            else:
                fld = row["field"]  # type: ignore[assignment]
                bg = C.HIGHLIGHT_YELLOW if _is_highlighted(fld, hl) else (C.ORANGE_SOFT if fld.is_padding else cell_tint)
                fg = C.ORANGE_DARK if fld.is_padding else C.TEXT_PRIMARY
                lbl = html.escape(fld.short_label)
                if interactive_links and not fld.is_spare and not fld.is_padding:
                    href = (
                        f"action:packet_field:{quote(fld.spec_number, safe='')}:"
                        f"{quote(fld.clause, safe='')}:{quote(fld.name, safe='')}"
                    )
                    lbl_html = f'<a href="{href}" style="color: {fg}; text-decoration: none; font-weight: 600;">{lbl}</a>'
                else:
                    lbl_html = f'<span style="font-weight: 600;">{lbl}</span>'
                parts.append(
                    f'<td colspan="8" style="border: 1px solid {C.BORDER_STRONG}; padding: 4px 6px; '
                    f'text-align: center; background-color: {bg}; color: {fg};">{lbl_html}</td>'
                )

            parts.append(
                f'<td style="border: 1px solid {C.BORDER_STRONG}; padding: 4px 8px; text-align: center; '
                f'background-color: {C.BG_CARD}; color: {C.TEXT_SECONDARY}; white-space: nowrap;">{html.escape(str(row["size_label"]))}</td>'
            )
            spec_cell = html.escape(str(row["spec_ref"]))
            parts.append(
                f'<td style="border: 1px solid {C.BORDER_STRONG}; padding: 4px 8px; text-align: left; '
                f'background-color: {C.BG_CARD}; color: {C.TEXT_SECONDARY}; white-space: nowrap;">{spec_cell}</td>'
            )
            parts.append("</tr>")

    parts.append("</tbody></table></div>")
    return "".join(parts)


def render_packet_markdown(packet: ComposedPacket) -> str:
    """Render packet stack as GitHub-Flavored Markdown."""
    ov_lbl, _html_chain_lbl, chain_lbl = _family_labels(packet.config.stack_family)
    lines = [
        f"### {packet.title}",
        "",
        f"- **Direction:** {packet.direction}",
        f"- **Interface:** {packet.interface}",
        f"- **Specs:** {', '.join(packet.spec_refs)}",
        f"- **Fixed {ov_lbl} Overhead:** {packet.overhead_3gpp_octets} octets",
    ]
    if packet.channel_mapping_chain:
        lines.append(f"- **Channel Mapping:** {' -> '.join(packet.channel_mapping_chain)}")
    if packet.overhead_transport_octets > 0:
        lines.append(f"- **Transport Overhead:** {packet.overhead_transport_octets} octets (Total = {packet.overhead_total_octets} octets)")
    if packet.ext_header_summaries:
        lines.append(f"- **{chain_lbl}:** {' -> '.join(packet.ext_header_summaries)}")
    lines.extend(["", "| Octet | Layer / Field | Size | Presence | Spec / Clause |", "| --- | --- | --- | --- | --- |"])

    cur_octet = 1
    for layer in packet.layers:
        rows, cur_octet = _pack_layer_rows(layer, cur_octet)
        spec_hdr = _format_spec_ref(layer.spec_number, f"Cl. {layer.clause}" if layer.clause and layer.clause != "-" else "")
        lines.append(f"| — | **{layer.title}** | — | — | {spec_hdr} |")
        for row in rows:
            if row["kind"] == "bits":
                summary = " | ".join(f"{f.short_label} ({f.bits}b)" for f in row["fields"])  # type: ignore[attr-defined]
                lines.append(f"| {row['octet_label']} | {summary} | {row['size_label']} | M/C | {row['spec_ref']} |")
            else:
                fld = row["field"]  # type: ignore[assignment]
                lines.append(
                    f"| {row['octet_label']} | {fld.short_label} | {row['size_label']} | {fld.presence} | {row['spec_ref']} |"
                )
    return "\n".join(lines)


def render_packet_plain_text(packet: ComposedPacket) -> str:
    """Render clean plain-text packet stack summary."""
    ov_lbl, _h, _m = _family_labels(packet.config.stack_family)
    lines = [
        packet.title,
        "=" * len(packet.title),
        f"Direction : {packet.direction}",
        f"Interface : {packet.interface}",
        f"Specs     : {', '.join(packet.spec_refs)}",
        f"Overhead  : {ov_lbl} = {packet.overhead_3gpp_octets} octets"
        + (f" (+ {packet.overhead_transport_octets} transport = {packet.overhead_total_octets} octets)" if packet.overhead_transport_octets else ""),
    ]
    if packet.channel_mapping_chain:
        lines.append(f"Channels  : {' -> '.join(packet.channel_mapping_chain)}")
    lines.append("")
    cur_octet = 1
    for layer in packet.layers:
        spec_hdr = _format_spec_ref(layer.spec_number, f"Cl. {layer.clause}" if layer.clause and layer.clause != "-" else "")
        lines.append(f"[{layer.title}]" + (f" ({spec_hdr})" if spec_hdr else ""))
        rows, cur_octet = _pack_layer_rows(layer, cur_octet)
        for row in rows:
            if row["kind"] == "bits":
                summary = " | ".join(f"{f.short_label}({f.bits}b)" for f in row["fields"])  # type: ignore[attr-defined]
                lines.append(f"  Oct {str(row['octet_label']):>5} : {summary}")
            else:
                fld = row["field"]  # type: ignore[assignment]
                lines.append(f"  Oct {str(row['octet_label']):>5} : {fld.short_label} [{row['size_label']}] ({row['spec_ref']})")
        lines.append("")
    return "\n".join(lines).strip()


def _sanitize_puml_label(label: str) -> str:
    return (label or "").replace(":", " -").replace("\n", " ").strip()


def render_packet_plantuml_packetdiag(packet: ComposedPacket) -> str:
    """
    Generate an editable PlantUML `@startpacketdiag ... @endpacketdiag` diagram
    representing the active full packet stack with contiguous bit offsets.
    """
    lines = [
        "@startpacketdiag",
        f"title {_sanitize_puml_label(packet.title)}",
        "colwidth = 32",
        "node_height = 48",
        "scale_interval = 8",
    ]

    bit_pos = 0
    for layer in packet.layers:
        lines.append(f"' === {layer.title} (TS {layer.spec_number} Cl. {layer.clause}) ===")
        for fld in layer.fields:
            lbl = _sanitize_puml_label(fld.short_label or fld.name)
            if fld.is_variable:
                disp_bits = max(16, (fld.variable_display_octets or 4) * 8)
                end_bit = bit_pos + disp_bits - 1
                lines.append(f"{bit_pos}-{end_bit}: {lbl} [colheight = 2]")
                bit_pos = end_bit + 1
                continue

            if fld.bits <= 0:
                lines.append(f"' (0-octet transparent pass-through: {lbl})")
                continue
            start_bit = bit_pos
            end_bit = bit_pos + fld.bits - 1
            bit_range = f"{start_bit}" if start_bit == end_bit else f"{start_bit}-{end_bit}"
            attrs: List[str] = []
            if fld.bits <= 2 and len(lbl) > 3:
                attrs.append("rotate = 270")
            elif fld.bits >= 64:
                attrs.append("colheight = 2")
            attr_str = f" [{', '.join(attrs)}]" if attrs else ""
            lines.append(f"{bit_range}: {lbl}{attr_str}")
            bit_pos = end_bit + 1

    lines.append("@endpacketdiag")
    return "\n".join(lines)


def copy_packet_to_clipboard(packet: ComposedPacket, mode: str = "word") -> str:
    """Copy the composed packet diagram to the clipboard in the chosen format."""
    clipboard = QApplication.clipboard()
    if not clipboard:
        return "Clipboard unavailable"

    if mode == "word":
        html_doc = (
            f'<div style="font-family: Calibri, \'Segoe UI\', Arial, sans-serif; max-width: 900px;">'
            f"{render_packet_html(packet, interactive_links=False)}</div>"
        )
        mime = QMimeData()
        mime.setHtml(compact_html_for_word(html_doc, full_document=True))
        mime.setText(render_packet_plain_text(packet))
        clipboard.setMimeData(mime)
        return "Copied Packet Diagram for Word / Outlook!"
    if mode == "packetdiag":
        clipboard.setText(render_packet_plantuml_packetdiag(packet))
        return "Copied PlantUML packetdiag!"
    if mode == "markdown":
        clipboard.setText(render_packet_markdown(packet))
        return "Copied Packet as Markdown!"
    clipboard.setText(render_packet_plain_text(packet))
    return "Copied Packet Plain Text!"
