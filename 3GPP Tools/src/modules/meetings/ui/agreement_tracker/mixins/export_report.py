"""Markdown/HTML Study Item dossier compilation and formatted Excel export mixin."""

import html
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List

import pandas as pd
from PyQt5.QtWidgets import QFileDialog, QMessageBox

from core.ui.theme import C
from core.utils.folder_opener import open_local_file
from core.utils.paths import to_extended_path


class _TrackerExportMixin:
    """Exports filtered agreements to styled HTML/Markdown dossiers and Excel workbooks."""

    def _export_study_item_report(self):
        rows = self.table_model.get_filtered_rows()
        if not rows:
            return

        wis = self.wi_token_widget.get_tokens()
        wi_slug = "_".join(wis[:2]) if wis else "All_Topics"
        wi_slug = re.sub(r'[^A-Za-z0-9_\-]', '_', wi_slug)
        default_name = f"3GPP_Agreements_Report_{wi_slug}_{datetime.now().strftime('%Y%m%d')}.md"

        out_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Study Item & Topic Agreements Report",
            default_name,
            "Markdown & HTML Dossier (*.md);;HTML Report Only (*.html)",
        )
        if not out_path:
            return

        try:
            md_text, html_text = self._compile_study_item_report_content(rows)
            p = Path(out_path)
            if p.suffix.lower() == ".html":
                with open(to_extended_path(p), "w", encoding="utf-8") as f:
                    f.write(html_text)
                open_local_file(str(p))
            else:
                with open(to_extended_path(p), "w", encoding="utf-8") as f:
                    f.write(md_text)
                html_companion = p.with_suffix(".html")
                with open(to_extended_path(html_companion), "w", encoding="utf-8") as f:
                    f.write(html_text)
                open_local_file(str(html_companion))
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export report:\n{e}")

    def _compile_study_item_report_content(self, rows: List[dict]):
        """Builds both a structured Markdown dossier and a rich standalone HTML report."""
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
        filter_desc = self.filter_summary_lbl.text()

        # Group by Work Item -> Meeting -> Agenda Item
        grouped: Dict[str, Dict[str, Dict[str, List[dict]]]] = {}
        for r in rows:
            wi = str(r.get("Work Item", "")).strip() or "Unspecified Work Item"
            mtg = f"{r.get('Meeting', '')} ({r.get('start_date', '')})"
            ai = str(r.get("AI", "")).strip() or "General"
            ai_title = str(r.get("AI Title", "")).strip()
            ai_header = f"AI {ai}: {ai_title}" if ai_title else f"AI {ai}"
            grouped.setdefault(wi, {}).setdefault(mtg, {}).setdefault(ai_header, []).append(r)

        md_lines = [
            "# ⚖️ 3GPP Cross-Meeting Study Item & Topic Agreements Dossier",
            "",
            f"- **Generated:** {now_str}",
            f"- **Filter Scope:** {filter_desc}",
            f"- **Total Matched Items:** {len(rows)}",
            "",
            "---",
            "",
        ]

        html_sections = []
        for wi, mtg_map in grouped.items():
            md_lines.append(f"## 🎯 Work Item / Study Item: `{wi}`")
            md_lines.append("")
            html_sections.append(
                f"<div class='wi-card'><h2>🎯 Work Item / Study Item: <code>{html.escape(wi)}</code></h2>"
            )

            for mtg_label, ai_map in mtg_map.items():
                md_lines.append(f"### 🗓️ {mtg_label}")
                md_lines.append("")
                html_sections.append(f"<h3>🗓️ {html.escape(mtg_label)}</h3>")

                for ai_header, items in ai_map.items():
                    md_lines.append(f"#### 📂 {ai_header}")
                    md_lines.append("")
                    html_sections.append(f"<h4>📂 {html.escape(ai_header)}</h4>")

                    for item in items:
                        kind = str(item.get("Kind", "Agreement"))
                        tdoc = str(item.get("TDoc", "")).strip()
                        ttitle = str(item.get("TDoc Title", "")).strip()
                        src = str(item.get("Source", "")).strip()
                        txt = str(item.get("Agreement Text", "")).strip()

                        icon = "🟢"
                        css_cls = "ag-green"
                        if "ffs" in kind.lower() or "open issue" in kind.lower():
                            icon = "🔵"
                            css_cls = "ag-blue"
                        elif "working assumption" in kind.lower():
                            icon = "🟡"
                            css_cls = "ag-amber"
                        elif "agreed tdoc" in kind.lower():
                            icon = "📄"
                            css_cls = "ag-indigo"

                        tdoc_suffix = ""
                        if tdoc:
                            tdoc_suffix = f" *(Related TDoc: **{tdoc}**{f' — {ttitle}' if ttitle else ''}{f' [{src}]' if src else ''})*"

                        md_lines.append(f"- {icon} **[{kind}]**{tdoc_suffix}")
                        for line in txt.split("\n"):
                            md_lines.append(f"  > {line}")
                        md_lines.append("")

                        body_html = "<br>".join(html.escape(l) for l in txt.split("\n"))
                        tdoc_html = ""
                        if tdoc:
                            tdoc_html = (
                                f"<div class='tdoc-ref'>Related TDoc: <b>{html.escape(tdoc)}</b>"
                                f"{f' — {html.escape(ttitle)}' if ttitle else ''}"
                                f"{f' <i>({html.escape(src)})</i>' if src else ''}</div>"
                            )
                        html_sections.append(
                            f"<div class='ag-box {css_cls}'>"
                            f"<div class='ag-badge'>{icon} {html.escape(kind)}</div>"
                            f"<div class='ag-text'>{body_html}</div>"
                            f"{tdoc_html}"
                            f"</div>"
                        )

            html_sections.append("</div>")

        html_doc = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>3GPP Study Item & Topic Agreements Dossier</title>
<style>
  body {{ font-family: 'Segoe UI', system-ui, sans-serif; background: {C.BG_SURFACE}; color: {C.TEXT_PRIMARY}; margin: 24px auto; max-width: 1120px; padding: 0 20px; }}
  h1 {{ color: {C.TEXT_PRIMARY}; border-bottom: 2px solid {C.BORDER_STRONG}; padding-bottom: 8px; }}
  .meta {{ background: {C.BG_CARD}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; padding: 12px 16px; margin-bottom: 20px; color: {C.TEXT_SECONDARY}; }}
  .wi-card {{ background: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 8px; padding: 16px 20px; margin-bottom: 24px; box-shadow: 0 1px 2px rgba(0,0,0,0.03); }}
  h2 {{ color: {C.INFO_NAVY}; margin-top: 0; }}
  h3 {{ color: {C.TEXT_BODY}; border-bottom: 1px solid {C.BORDER_SUBTLE}; padding-bottom: 4px; margin-top: 18px; }}
  h4 {{ color: {C.TEXT_SECONDARY}; margin: 12px 0 8px 0; }}
  .ag-box {{ border-radius: 6px; padding: 10px 14px; margin-bottom: 10px; border-left: 4px solid {C.GREEN_ACCENT}; background: {C.SUCCESS_SOFT}; }}
  .ag-blue {{ border-left-color: {C.ACCENT_ALT}; background: {C.INFO_SOFT}; }}
  .ag-amber {{ border-left-color: {C.WARNING_AMBER}; background: {C.WARNING_SOFT}; }}
  .ag-indigo {{ border-left-color: {C.INDIGO_DARK}; background: {C.WG_SA_BG}; }}
  .ag-badge {{ font-weight: bold; font-size: 12px; margin-bottom: 4px; color: {C.TEXT_STRONG}; }}
  .ag-text {{ font-size: 13px; line-height: 1.5; color: {C.TEXT_PRIMARY}; }}
  .tdoc-ref {{ margin-top: 6px; font-size: 12px; color: {C.TEXT_BODY}; }}
</style>
</head>
<body>
  <h1>⚖️ 3GPP Cross-Meeting Study Item & Topic Agreements Dossier</h1>
  <div class="meta">
    <b>Generated:</b> {html.escape(now_str)}<br>
    <b>Filter Scope:</b> {html.escape(filter_desc)}<br>
    <b>Total Matched Items:</b> {len(rows)}
  </div>
  {''.join(html_sections)}
</body>
</html>"""

        return "\n".join(md_lines), html_doc

    def _export_to_excel(self):
        rows = self.table_model.get_filtered_rows()
        if not rows:
            return

        wis = self.wi_token_widget.get_tokens()
        wi_slug = "_".join(wis[:2]) if wis else "All_Topics"
        wi_slug = re.sub(r'[^A-Za-z0-9_\-]', '_', wi_slug)
        default_name = f"3GPP_Agreements_{wi_slug}_{datetime.now().strftime('%Y%m%d')}.xlsx"

        out_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Agreements to Excel",
            default_name,
            "Excel Workbook (*.xlsx)",
        )
        if not out_path:
            return

        try:
            export_records = []
            for r in rows:
                export_records.append({
                    "WG": r.get("WG", ""),
                    "Meeting": r.get("Meeting", ""),
                    "Date": r.get("start_date", ""),
                    "Agenda Item": r.get("AI", ""),
                    "AI / Topic Title": r.get("AI Title", ""),
                    "Work Item / Study Item": r.get("Work Item", ""),
                    "Release": r.get("Release", ""),
                    "Type": r.get("Kind", ""),
                    "Agreement / Conclusion Text": r.get("Agreement Text", ""),
                    "Related TDoc": r.get("TDoc", ""),
                    "TDoc Title": r.get("TDoc Title", ""),
                    "Source": r.get("Source", ""),
                    "Chair Notes File": r.get("Chair Notes File", ""),
                })
            df = pd.DataFrame(export_records)
            df.to_excel(to_extended_path(out_path), index=False)
            open_local_file(str(out_path))
        except Exception as e:
            QMessageBox.critical(self, "Excel Export Error", f"Failed to export Excel workbook:\n{e}")
