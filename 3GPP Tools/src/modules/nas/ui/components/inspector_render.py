"""Definition rendering, cross-ref parsing, and usage menu mixin for NASInspectorWidget."""

from __future__ import annotations

import html
import re
from typing import Any, Dict, List, Optional, Tuple

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QAction, QMenu

from core.ui.theme import C
from modules.nas.core.nas_db import parse_version_tuple
from modules.nas.ui.nas_dialogs import get_spec_metadata


RE_NAS_CROSS_REF = re.compile(
    r"(?:see\s+)?(?:subclause|clause|sub-clause)\s+([0-9A-Za-z\.]+)\s+in\s+(?:3GPP\s+)?TS\s+(24\.(?:301|501|526|588|555|578|514))",
    re.IGNORECASE,
)
RE_EXPLICIT_SPEC_REF = re.compile(
    r"^TS\s+((?:24|25|29|36|37|38)\.[0-9]{3})\s+([0-9A-Za-z\.]+)$",
    re.IGNORECASE,
)


class NASInspectorRenderMixin:
    """Renders Clause 9 HTML/ASN.1 definitions, inline fallbacks, and containing-message menus."""

    def display_definitions(
            self,
            clause: str,
            ie_name: str,
            spec_number: Optional[str],
            defs: List[Dict[str, Any]],
            containing_msgs: List[Dict[str, Any]],
            target_version_hint: Optional[str] = None,
            fallback_type_ref: str = "",
            field_description: str = "",
            **kwargs,
    ):
        self._current_ie_clause = clause
        self._current_ie_name = ie_name
        self._current_ie_spec = spec_number
        self._current_containing_msgs = containing_msgs
        self._current_message_name_hint = str(kwargs.get("message_name", "") or "")

        meta = get_spec_metadata(spec_number) if spec_number else {}
        acronym_str = f" [{meta.get('acronym', f'TS {spec_number}')}]" if spec_number else ""

        if defs:
            resolved_name = defs[0]["ie_name"]
            full_badge = f"{resolved_name}{acronym_str}"
            elided_badge = self.title_lbl.fontMetrics().elidedText(
                full_badge, Qt.ElideRight, max(140, self.title_lbl.maximumWidth() - 18)
            )
            self.title_lbl.setText(elided_badge)
            self.title_lbl.setToolTip(full_badge)
            self._current_clause_defs = {f"{d['spec_number']} v{d['version']}": d for d in defs}

            if containing_msgs:
                num_msgs = len(containing_msgs)
                self.usage_btn.setText(f"Used in: {num_msgs} msg{'s' if num_msgs != 1 else ''} ▾")
                self.usage_btn.setEnabled(True)
                self.usage_btn.setVisible(True)
            else:
                self.usage_btn.setEnabled(False)
                self.usage_btn.setVisible(False)

            self._updating_combo = True
            self.version_combo.clear()
            for d in defs:
                clean_v = str(d.get("version", "")).lstrip("vV")
                parts = clean_v.split(".")
                rel_str = f"· R{parts[0]} " if parts and parts[0].isdigit() else ""
                key_name = f"TS {d['spec_number']} {rel_str}(v{d['version']})" if rel_str else f"TS {d['spec_number']} v{d['version']}"
                self.version_combo.addItem(key_name, f"{d['spec_number']} v{d['version']}")

            target_version_key: Optional[str] = None
            if target_version_hint:
                clean_hint = target_version_hint.replace("TS ", "").strip()
                if clean_hint in self._current_clause_defs:
                    target_version_key = clean_hint
                else:
                    for k in self._current_clause_defs:
                        if target_version_hint in k or k.endswith(f"v{target_version_hint}"):
                            target_version_key = k
                            break

            if not target_version_key and defs:
                target_version_key = f"{defs[0]['spec_number']} v{defs[0]['version']}"

            for idx in range(self.version_combo.count()):
                if self.version_combo.itemData(idx) == target_version_key:
                    self.version_combo.setCurrentIndex(idx)
                    break

            self._updating_combo = False
            selected_def = self._current_clause_defs.get(target_version_key) or defs[0]
            self._render_selected_definition(selected_def)
            self._update_open_spec_menu()
            self._update_packet_stack_availability(
                spec_number=spec_number,
                clause=clause,
                ie_name=ie_name,
                version_hint=selected_def.get("version", "") or target_version_hint or "",
            )
            self.stack.setCurrentIndex(1)
            self.copy_btn.setEnabled(True)
            return

        self.usage_btn.setVisible(False)
        self.version_combo.clear()
        self._current_clause_defs.clear()

        self.title_lbl.setText(f"{ie_name}{acronym_str}")

        is_inline_type = any(
            fallback_type_ref.strip().upper().startswith(kw)
            for kw in ("ENUMERATED", "INTEGER", "BOOLEAN", "BIT STRING", "OCTET STRING", "NULL", "CHOICE", "SEQUENCE")
        )

        match_explicit = RE_EXPLICIT_SPEC_REF.match(fallback_type_ref.strip())
        if match_explicit:
            ext_spec = match_explicit.group(1)
            ext_clause = match_explicit.group(2)
            ext_meta = get_spec_metadata(ext_spec)
            html_output = (
                f'<h2 style="color: {C.ACCENT_SKY_DARK}; margin-top: 2px; margin-bottom: 8px; '
                f'font-family: Segoe UI, sans-serif; font-size: 14px; border-bottom: 1px solid {C.BORDER_SUBTLE}; padding-bottom: 4px;">'
                f'{html.escape(ie_name)} (TS {html.escape(ext_spec)} Clause {html.escape(ext_clause)})</h2>'
                f'<div style="margin-top: 10px; border: 1px solid {C.WARNING_BORDER_STRONG}; background-color: {C.WARNING_SOFT}; border-radius: 6px; padding: 12px; border-left: 4px solid {C.WARNING_ACCENT}; font-family: Segoe UI, sans-serif;">'
                f'<div style="font-weight: bold; color: {C.WARNING_DARK}; font-size: 12px; margin-bottom: 4px;">'
                f'ℹ️ Cross-Specification Reference: TS {html.escape(ext_spec)} — {html.escape(ext_meta.get("title", ""))} (Clause {html.escape(ext_clause)})'
                f'</div>'
                f'<div style="color: {C.WARNING_BROWN_DEEP}; font-size: 11px; line-height: 1.5;">'
                f'This UE policy structure is defined in <b>3GPP TS {html.escape(ext_spec)} Clause {html.escape(ext_clause)}</b>, which is not currently present in your local database.<br/>'
                f'<p style="margin-top: 8px; margin-bottom: 2px;">'
                f'<a href="action:import:{ext_spec}:0" style="display: inline-block; background-color: {C.ACCENT_SKY}; color: {C.BG_CARD}; padding: 4px 12px; border-radius: 4px; text-decoration: none; font-weight: bold; font-size: 11px;">'
                f'📥 Import TS {html.escape(ext_spec)} from 3GPP Archive'
                f'</a>'
                f'</p>'
                f'</div>'
                f'</div>'
            )
            self._current_html_content = html_output
            self.text_area.setHtml(html_output)
            self._update_open_spec_menu()
            self.stack.setCurrentIndex(1)
            self.copy_btn.setEnabled(True)
        elif is_inline_type or fallback_type_ref:
            html_output = self._render_inline_type_html(
                ie_name=ie_name,
                type_ref=fallback_type_ref,
                field_description=field_description,
                spec_number=spec_number,
            )
            self._current_html_content = html_output
            self.text_area.setHtml(html_output)
            self._update_open_spec_menu()
            self.stack.setCurrentIndex(1)
            self.copy_btn.setEnabled(True)
        else:
            fb_text = f"Type / Reference: {fallback_type_ref}\n(No structure definition found in database)"
            self._current_html_content = fb_text
            self.text_area.setPlainText(fb_text)
            self._update_open_spec_menu()
            self.stack.setCurrentIndex(1)
            self.copy_btn.setEnabled(bool(fallback_type_ref))
        self._update_packet_stack_availability(
            spec_number=spec_number,
            clause=clause,
            ie_name=ie_name,
            version_hint=target_version_hint or "",
        )

    def _render_selected_definition(self, definition: Dict[str, Any]):
        base_html = definition.get("raw_description", "")
        current_spec = definition.get("spec_number", "")
        current_ver_str = definition.get("version", "")
        current_clause = definition.get("clause", "")

        # Check for 4G <-> 5G and UE Policy (24.501 -> 24.526, 24.588, 24.555, 24.514, 24.578) cross-specification references
        if self.db and current_spec in ("24.501", "24.301"):
            cross_targets: List[Tuple[str, str]] = []
            match = RE_NAS_CROSS_REF.search(base_html)
            if match:
                target_clause = match.group(1).rstrip(".[] ")
                target_spec = match.group(2).strip()
                if target_spec != current_spec:
                    cross_targets.append((target_spec, target_clause))

            if current_spec == "24.501" and current_clause in ("D.6.2", "D.6.8"):
                policy_clause_map = {
                    "24.526": "5.2",
                    "24.588": "5.2",
                    "24.555": "5.2",
                    "24.514": "12.2",
                    "24.578": "5.2",
                }
                for p_spec, p_clause in policy_clause_map.items():
                    if p_spec in base_html and not any(t[0] == p_spec for t in cross_targets):
                        cross_targets.append((p_spec, p_clause))

            ver_tuple = parse_version_tuple(current_ver_str)
            major_ver = (ver_tuple[0][1] if isinstance(ver_tuple[0], tuple) else ver_tuple[0]) if ver_tuple else 15

            for target_spec, target_clause in cross_targets:
                cross_def = self.db.get_cross_referenced_ie_definition(
                    target_spec=target_spec,
                    target_clause=target_clause,
                    alt_name=definition.get("ie_name", ""),
                    major_version=major_ver,
                )
                if not cross_def:
                    cross_def = self.db.get_cross_referenced_ie_definition(
                        target_spec=target_spec,
                        target_clause=target_clause,
                        alt_name="",
                        major_version=None,
                    )

                if cross_def and cross_def.get("raw_description"):
                    cross_body = cross_def.get("raw_description", "").strip()
                    c_ver = cross_def.get("version", "")
                    c_cl = cross_def.get("clause", target_clause)
                    cross_html = (
                        f'<div style="margin-top: 14px; border: 1.5px solid {C.ACCENT_SKY}; background-color: {C.SKY_SURFACE}; border-radius: 6px; padding: 12px;">'
                        f'<div style="margin-bottom: 10px; border-bottom: 1px solid {C.SKY_BORDER}; padding-bottom: 6px;">'
                        f'<span style="font-size: 12px; font-weight: bold; color: {C.ACCENT_SKY_DARK};">'
                        f'<a href="action:open_spec:{target_spec}:{c_ver}:{c_cl}" style="color: {C.ACCENT_SKY_DARK}; text-decoration: underline;" title="Click to open TS {target_spec} v{c_ver} in Word">'
                        f'🔗 Cross-Referenced from TS {target_spec} v{c_ver} (Clause {c_cl}) [Open Word]'
                        f'</a>'
                        f'</span>'
                        f'<span style="font-size: 10px; color: {C.ACCENT_SKY}; margin-left: 8px; background-color: {C.SKY_SOFT}; padding: 2px 6px; border-radius: 3px; border: 1px solid {C.SKY_BORDER_STRONG};">'
                        f'v{c_ver}'
                        f'</span>'
                        f'</div>'
                        f'<div style="background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; padding: 10px;">'
                        f'{cross_body}'
                        f'</div>'
                        f'</div>'
                    )
                    base_html += cross_html
                else:
                    missing_html = (
                        f'<div style="margin-top: 14px; border: 1px solid {C.WARNING_BORDER_STRONG}; background-color: {C.WARNING_SOFT}; border-radius: 6px; padding: 10px; border-left: 4px solid {C.WARNING_ACCENT}; font-family: Segoe UI, sans-serif;">'
                        f'<div style="font-weight: bold; color: {C.WARNING_DARK}; font-size: 12px; margin-bottom: 4px;">'
                        f'ℹ️ Cross-Specification Reference: TS {target_spec} (Clause {target_clause})'
                        f'</div>'
                        f'<div style="color: {C.WARNING_BROWN_DEEP}; font-size: 11px; line-height: 1.5;">'
                        f'This Information Element references <b>TS {target_spec} (Release {major_ver})</b>, which is not currently present in your local database.<br/>'
                        f'<p style="margin-top: 6px; margin-bottom: 2px;">'
                        f'<a href="action:import:{target_spec}:{major_ver}" style="display: inline-block; background-color: {C.ACCENT_SKY}; color: {C.BG_CARD}; padding: 4px 12px; border-radius: 4px; text-decoration: none; font-weight: bold; font-size: 11px;">'
                        f'📥 Import TS {target_spec} (Rel-{major_ver}) from 3GPP Archive'
                        f'</a>'
                        f'</p>'
                        f'</div>'
                        f'</div>'
                    )
                    base_html += missing_html

        self._current_html_content = base_html
        self.text_area.setHtml(base_html)

    def _on_version_changed(self, index: int):
        if self._updating_combo or index < 0:
            return
        version_key = self.version_combo.itemData(index)
        if version_key in self._current_clause_defs:
            selected_def = self._current_clause_defs[version_key]
            self._render_selected_definition(selected_def)
            self._update_open_spec_menu()
            self._update_packet_stack_availability(
                spec_number=selected_def.get("spec_number") or self._current_ie_spec,
                clause=self._current_ie_clause or "",
                ie_name=self._current_ie_name or "",
                version_hint=selected_def.get("version", ""),
            )

    @staticmethod
    def _render_inline_type_html(
        ie_name: str,
        type_ref: str,
        field_description: str = "",
        spec_number: Optional[str] = None,
    ) -> str:
        meta = get_spec_metadata(spec_number) if spec_number else {}
        spec_title = f" ({meta.get('acronym', f'TS {spec_number}')})" if spec_number else ""
        header_html = (
            f'<h2 style="color: {C.ACCENT_SKY_DARK}; margin-top: 2px; margin-bottom: 8px; '
            f'font-family: Segoe UI, sans-serif; font-size: 14px; border-bottom: 1px solid {C.BORDER_SUBTLE}; padding-bottom: 4px;">'
            f'{html.escape(ie_name)}{spec_title}</h2>'
        )

        clean_type_ref = re.sub(r"\s+", " ", type_ref).strip()
        formatted_asn1 = f"{ie_name} ::= {clean_type_ref}"
        if "ENUMERATED" in clean_type_ref:
            enum_match = re.search(r"ENUMERATED\s*\{([^}]+)\}", clean_type_ref, re.IGNORECASE)
            if enum_match:
                inner_vals = [v.strip() for v in enum_match.group(1).split(",") if v.strip()]
                indented_vals = ",\n    ".join(inner_vals)
                formatted_asn1 = f"{ie_name} ::= ENUMERATED {{\n    {indented_vals}\n}}"

        asn1_html = (
            f'<div style="margin-bottom: 10px;">'
            f'<div style="font-size: 11px; font-weight: bold; color: {C.TEXT_SECONDARY}; margin-bottom: 3px;">Inline ASN.1 Definition:</div>'
            f'<pre style="background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; padding: 8px; '
            f'border-radius: 4px; font-family: Consolas, monospace; font-size: 11px; color: {C.TEXT_PRIMARY}; '
            f'white-space: pre-wrap; margin: 0;">{html.escape(formatted_asn1)}</pre>'
            f'</div>'
        )

        desc_html = ""
        if field_description and field_description.strip():
            desc_html = (
                f'<div style="margin-top: 8px;">'
                f'<div style="font-size: 11px; font-weight: bold; color: {C.TEXT_SECONDARY}; margin-bottom: 4px;">'
                f'Field Description & Semantics:</div>'
                f'<div style="background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_SUBTLE}; border-left: 3px solid {C.ACCENT_SKY}; '
                f'padding: 8px; border-radius: 3px; font-family: Segoe UI, sans-serif; font-size: 11px; color: {C.TEXT_BODY}; line-height: 1.4;">'
                f'{html.escape(field_description)}'
                f'</div>'
                f'</div>'
            )

        return f"{header_html}{asn1_html}{desc_html}"

    def clear_display(self):
        self.title_lbl.setText("No Selection")
        self.usage_btn.setVisible(False)
        self.version_combo.clear()
        self.text_area.clear()
        self._current_html_content = ""
        self._current_clause_defs.clear()
        self._packet_available = False
        if hasattr(self, "mode_segment"):
            self.mode_segment.setVisible(False)
        if hasattr(self, "content_stack"):
            self.content_stack.setCurrentIndex(0)
        self.copy_btn.setEnabled(False)
        self.open_spec_btn.setEnabled(False)
        self.open_spec_btn.setText("📖 Open Spec")
        self.open_spec_btn.setMenu(None)
        self.stack.setCurrentIndex(0)

    def set_inspector_mode(self, mode: str) -> None:
        clean = "packet" if (mode == "packet" and getattr(self, "_packet_available", False)) else "definition"
        self._inspector_mode = clean
        if hasattr(self, "btn_mode_definition"):
            self.btn_mode_definition.blockSignals(True)
            self.btn_mode_definition.setChecked(clean == "definition")
            self.btn_mode_definition.blockSignals(False)
        if hasattr(self, "btn_mode_packet"):
            self.btn_mode_packet.blockSignals(True)
            self.btn_mode_packet.setChecked(clean == "packet")
            self.btn_mode_packet.blockSignals(False)
        if hasattr(self, "content_stack"):
            self.content_stack.setCurrentIndex(1 if clean == "packet" else 0)

    def _update_packet_stack_availability(
        self,
        *,
        spec_number: Optional[str],
        clause: str = "",
        ie_name: str = "",
        version_hint: str = "",
    ) -> None:
        if not hasattr(self, "packet_stack_widget"):
            return
        ver_clean = str(version_hint or "").strip()
        if " v" in ver_clean:
            ver_clean = ver_clean.split(" v", 1)[1].strip()
        m_ver = re.search(r"([0-9]+\.[0-9]+\.[0-9]+)", ver_clean)
        ver_clean = m_ver.group(1) if m_ver else ""

        available = self.packet_stack_widget.set_context(
            db=self.db,
            spec_number=spec_number,
            clause=clause,
            ie_name=ie_name,
            message_name=getattr(self, "_current_message_name_hint", ""),
            version_hint=ver_clean,
        )
        self._packet_available = available
        self.mode_segment.setVisible(available)
        if available and getattr(self, "_inspector_mode", "definition") == "packet":
            self.set_inspector_mode("packet")
        else:
            self.set_inspector_mode("definition")

    def _on_inspect_packet_field(self, spec_number: str, clause: str, field_name: str) -> None:
        defs = []
        if self.db and spec_number in ("29.281", "38.415", "38.413", "38.412", "24.501"):
            defs = self.db.get_ie_definitions_by_clause(
                clause=clause,
                alt_name=field_name,
                spec_number=spec_number,
            )
        self.set_inspector_mode("definition")
        if defs:
            self.display_definitions(
                clause=clause,
                ie_name=field_name,
                spec_number=spec_number,
                defs=defs,
                containing_msgs=[],
                fallback_type_ref=f"{field_name} (Clause {clause})",
            )

    def _show_usage_menu(self):
        if not self._current_ie_clause:
            return

        clause = self._current_ie_clause
        name = self._current_ie_name or ""
        spec_num = self._current_ie_spec

        menu = QMenu(self)
        menu.setStyleSheet(f"""
            QMenu {{
                font-size: 11px;
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                padding: 4px;
            }}
            QMenu::item {{
                padding: 4px 20px 4px 10px;
                border-radius: 3px;
            }}
            QMenu::item:selected {{
                background-color: {C.SKY_SOFT};
                color: {C.ACCENT_SKY_DARK};
            }}
            QMenu::separator {{
                height: 1px;
                background-color: {C.BORDER_SUBTLE};
                margin: 4px 0;
            }}
        """)

        header_title = f"Messages referencing {name or clause}"
        if spec_num:
            meta = get_spec_metadata(spec_num)
            header_title += f" [{meta['acronym']}]"
        header_action = QAction(f"{header_title}:", self)
        header_action.setEnabled(False)
        menu.addAction(header_action)
        menu.addSeparator()

        if self._current_containing_msgs:
            for m in self._current_containing_msgs:
                msg_name = m["message_name"]
                clause_ref = m["clause"]
                m_spec = m.get("spec_number", "")
                spec_tag = f" [{m_spec}]" if m_spec and not spec_num else ""
                action = QAction(f"{msg_name} ({clause_ref}){spec_tag}", self)
                action.triggered.connect(lambda checked, mn=msg_name: self.jump_to_message_requested.emit(mn))
                menu.addAction(action)
        else:
            none_act = QAction("No messages found in active releases", self)
            none_act.setEnabled(False)
            menu.addAction(none_act)

        menu.addSeparator()
        filter_action = QAction(f"🔍 Filter message list by '{name or clause}'", self)
        filter_action.triggered.connect(lambda: self.filter_by_ie_requested.emit(name or clause))
        menu.addAction(filter_action)

        menu.exec_(self.usage_btn.mapToGlobal(self.usage_btn.rect().bottomLeft()))

    def set_containing_messages(self, containing_msgs: List[Dict[str, Any]]):
        self._current_containing_msgs = containing_msgs
        num_msgs = len(containing_msgs)
        if num_msgs > 0:
            self.usage_btn.setText(f"Used in: {num_msgs} msg{'s' if num_msgs != 1 else ''} ▾")
            self.usage_btn.setVisible(True)
            self.usage_btn.setEnabled(True)
        else:
            self.usage_btn.setVisible(False)
            self.usage_btn.setEnabled(False)
