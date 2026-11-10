# --- File: src/modules/meetings/core/markdown_exporter.py ---
import json
import os
import re
import logging
from pathlib import Path
from typing import Dict, List, Optional
from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.company_sanitizer import CompanySanitizer
from core.utils.paths import to_extended_path
from modules.meetings.core.tdocs.tdoc_url_resolver import resolve_tdoc_download_url

logger = logging.getLogger(__name__)


def natural_ai_sort_key(ai_str: str) -> list:
    """Sort key for agenda items (e.g. '8.1', '8.1.1', '8.2', '9', '10')."""
    parts = re.split(r'(\d+)', str(ai_str or ""))
    return [int(p) if p.isdigit() else p.lower() for p in parts]


class MarkdownExporterThread(QThread):
    finished = pyqtSignal(bool, str)

    def __init__(
        self,
        meeting_dir: Path,
        tdocs_data: list,
        docs_url: str,
        mtg_info: dict,
        include_chair_agreements: Optional[bool] = None,
        agenda_map: Optional[dict] = None,
    ):
        super().__init__()
        self.meeting_dir = Path(meeting_dir)
        self.tdocs_data = tdocs_data or []
        self.docs_url = docs_url or ""
        self.mtg_info = mtg_info or {}
        self.agenda_map = agenda_map or {}
        self.export_dir = self.meeting_dir / "Export"

        # Reliably resolve the root `src/` directory
        # src/modules/meetings/core/exporters/markdown_exporter.py -> parents[4] is src
        src_dir = Path(__file__).resolve().parents[4]
        config_path = src_dir / "config" / "export_config.json"

        # Define default configuration
        default_config = {
            "columns_for_3gu_tdoc_export": ["TDoc", "Agenda Item", "Type", "For", "Title", "Source", "Abstract"],
            "columns_for_3gu_tdoc_export_ls": ["TDoc", "Agenda Item", "Type", "For", "Title", "Source", "Abstract"],
            "columns_for_3gu_tdoc_export_ls_out": ["Abstract", "TDoc", "Title", "Agenda Item", "Reply to"],
            "columns_for_3gu_tdoc_export_pcr": ["TDoc", "Agenda Item", "Type", "For", "Title", "Source", "Abstract"],
            "columns_for_3gu_tdoc_export_cr": ["TDoc", "Agenda Item", "Type", "For", "Title", "Source", "Abstract"],
            "columns_for_3gu_tdoc_export_contributor": ["TDoc", "Type", "Title", "Source", "Agenda Item", "TDoc Status",
                                                        "Spec", "CR category"],
            "company_name_regex_for_report": "Deutsche Telekom",
            "include_chair_agreements": True,
            "include_agreements_summary": True,
        }

        self.config = {}

        if not config_path.exists():
            try:
                with open(to_extended_path(config_path), "w", encoding="utf-8") as f:
                    json.dump(default_config, f, indent=4)
                self.config = default_config
            except Exception as e:
                logger.error(f"[MarkdownExporter] Failed to create default config file: {e}")
                self.config = default_config
        else:
            try:
                with open(to_extended_path(config_path), "r", encoding="utf-8") as f:
                    self.config = json.load(f)
            except Exception as e:
                logger.error(f"[MarkdownExporter] Failed to load config, falling back to defaults: {e}")
                self.config = default_config

        # Check whether this is a RAN meeting
        wg_name = str(self.mtg_info.get("wg_name", "")).strip().upper()
        is_ran = wg_name.startswith("RAN") or "RAN" in wg_name

        # Resolve include_chair_agreements preference:
        # 1. Explicit argument if passed
        # 2. Config setting if defined
        # 3. Default True for RAN meetings (and True by default)
        if include_chair_agreements is not None:
            self.include_chair_agreements = bool(include_chair_agreements)
        elif "include_chair_agreements" in self.config:
            self.include_chair_agreements = bool(self.config.get("include_chair_agreements", True))
        else:
            self.include_chair_agreements = True if is_ran else True

        self.include_agreements_summary = bool(self.config.get("include_agreements_summary", True))

    def run(self):
        try:
            os.makedirs(to_extended_path(self.export_dir), exist_ok=True)

            # Load column configurations with fallbacks to defaults
            cols_ls_in = self.config.get(
                "columns_for_3gu_tdoc_export_ls",
                ["TDoc", "Agenda Item", "Type", "For", "Title", "Source", "Abstract"],
            )
            cols_ls_out = self.config.get(
                "columns_for_3gu_tdoc_export_ls_out",
                ["Abstract", "TDoc", "Title", "Agenda Item", "Reply to"],
            )
            cols_pcr = self.config.get(
                "columns_for_3gu_tdoc_export_pcr",
                ["TDoc", "Agenda Item", "Type", "For", "Title", "Source", "Abstract"],
            )
            cols_cr = self.config.get(
                "columns_for_3gu_tdoc_export_cr",
                ["TDoc", "Agenda Item", "Type", "For", "Title", "Source", "Abstract"],
            )
            cols_company = self.config.get(
                "columns_for_3gu_tdoc_export_contributor",
                ["TDoc", "Title", "Source", "Agenda Item", "TDoc Status", "Spec", "CR category"],
            )
            company_regex = self.config.get("company_name_regex_for_report", "Deutsche Telekom")

            # Extract Meeting info for the Markdown Header
            wg_name = str(self.mtg_info.get('wg_name', 'WG'))
            mtg_num = str(self.mtg_info.get('meeting_number', ''))
            mtg_url = self.docs_url.replace('/Docs', '').replace('/docs', '').rstrip('/') if self.docs_url else ""
            header_comment = f"<!--- [{wg_name}#{mtg_num}]({mtg_url}) --->\n\n"

            ai_groups = {}
            company_docs = []

            # 1. Parse and categorize the TDoc data
            for row in self.tdocs_data:
                status = str(row.get('TDoc Status', '')).lower()
                doc_type = str(row.get('Type', '')).lower()
                source = str(row.get('Source', ''))

                is_withdrawn = 'withdrawn' in status
                is_agreed = any(x in status for x in ['agreed', 'approved'])

                is_ls_in = 'ls in' in doc_type or doc_type == 'ls'
                is_ls_out = 'ls out' in doc_type or 'ls_out' in doc_type
                is_pcr = 'pcr' in doc_type or 'p-cr' in doc_type
                is_cr = ('cr' in doc_type and not is_pcr)

                # Add to Company Report only if a CONFIRMED source segment matches (excludes ? and [...])
                confirmed_source_text = " ".join(
                    seg for seg, is_tent in CompanySanitizer.tokenize_source_segments(source) if not is_tent
                )
                if re.search(company_regex, confirmed_source_text, re.IGNORECASE) and not is_withdrawn:
                    company_docs.append(row)

                # Process Agenda Items (Handling overlapping AIs like "20.1, 20.2")
                raw_ais = str(row.get('Agenda Item', ''))
                ais = [ai.strip() for ai in re.split(r'[,/&]', raw_ais) if ai.strip()]

                for ai in ais:
                    if ai not in ai_groups:
                        ai_groups[ai] = {'ls_in': [], 'ls_out': [], 'pcr': [], 'cr': []}

                    if is_ls_in:
                        ai_groups[ai]['ls_in'].append(row)
                    elif is_ls_out and is_agreed:
                        ai_groups[ai]['ls_out'].append(row)
                    elif is_pcr and is_agreed:
                        ai_groups[ai]['pcr'].append(row)
                    elif is_cr and is_agreed:
                        ai_groups[ai]['cr'].append(row)

            # 2. Extract / load Chairman's Notes agreements if enabled
            agreements_by_ai: Dict[str, List[dict]] = {}
            all_agreements: List[dict] = []
            company_agreements: List[dict] = []

            if self.include_chair_agreements:
                from modules.meetings.core.tdocs.chair_notes_agreements_parser import ChairNotesAgreementsParser
                _, raw_agreements, _ = ChairNotesAgreementsParser.scan_and_cache_chair_notes(
                    self.meeting_dir, self.mtg_info, self.agenda_map
                )
                all_agreements = [a for a in (raw_agreements or []) if not a.get("is_previous_meeting_recap")]

                for ag in all_agreements:
                    ai_num = str(ag.get("ai", "")).strip() or "Unassigned"
                    agreements_by_ai.setdefault(ai_num, []).append(ag)

                    # Check if company contribution is referenced by TDoc ID or in text
                    tid = str(ag.get("tdoc_id", "")).strip().upper()
                    t_row = next((r for r in self.tdocs_data if str(r.get("TDoc", "")).strip().upper() == tid), None) if tid else None
                    is_comp = False
                    if t_row:
                        src = str(t_row.get("Source", ""))
                        conf_src = " ".join(seg for seg, tent in CompanySanitizer.tokenize_source_segments(src) if not tent)
                        if re.search(company_regex, conf_src, re.IGNORECASE):
                            is_comp = True
                    elif re.search(company_regex, ag.get("text", ""), re.IGNORECASE):
                        is_comp = True

                    if is_comp:
                        company_agreements.append(ag)

            # 3. Write Agenda Item Markdown Files
            # Combine all Agenda Items from TDocs and from Chairman's Notes agreements
            all_ais = sorted(list(set(ai_groups.keys()) | set(agreements_by_ai.keys())), key=natural_ai_sort_key)

            for ai in all_ais:
                groups = ai_groups.get(ai, {'ls_in': [], 'ls_out': [], 'pcr': [], 'cr': []})
                agreements_for_ai = agreements_by_ai.get(ai, [])

                # If no documents and no agreements, skip
                if not any(groups.values()) and not agreements_for_ai:
                    continue

                safe_ai = re.sub(r'[\\/*?:"<>|]', "_", ai)
                filepath = self.export_dir / f"{safe_ai}.md"

                with open(to_extended_path(filepath), "w", encoding="utf-8") as f:
                    f.write(header_comment)

                    # Table sections for documents
                    self._write_table(f, "Following LS were received:", groups['ls_in'], cols_ls_in)
                    self._write_table(f, "Following LS OUTs were agreed:", groups['ls_out'], cols_ls_out)
                    self._write_table(f, "Following pCRs were agreed:", groups['pcr'], cols_pcr)
                    self._write_table(f, "Following CRs were agreed:", groups['cr'], cols_cr)

                    # Chairman's Notes Agreements section
                    if agreements_for_ai:
                        self._write_agreements_section(
                            f,
                            "Following Agreements & Working Assumptions were captured in Chairman's Notes:",
                            agreements_for_ai,
                        )

            # 4. Write Master Agreements Markdown File if agreements exist
            if self.include_chair_agreements and self.include_agreements_summary and all_agreements:
                agreements_path = self.export_dir / "Agreements.md"
                with open(to_extended_path(agreements_path), "w", encoding="utf-8") as f:
                    f.write(header_comment)
                    f.write(f"# {wg_name}#{mtg_num} — Chairman's Notes Agreements & Working Assumptions\n\n")

                    total = len(all_agreements)
                    wa_count = sum(1 for a in all_agreements if "working assumption" in str(a.get("kind", "")).lower())
                    ffs_count = sum(1 for a in all_agreements if "ffs" in str(a.get("kind", "")).lower() or "open issue" in str(a.get("kind", "")).lower())
                    ag_count = total - wa_count - ffs_count

                    f.write(
                        f"> **Summary:** **{total}** items captured in Chairman's Notes "
                        f"(🟢 **{ag_count}** Agreements, 🟡 **{wa_count}** Working Assumptions, 🔵 **{ffs_count}** FFS / Open Issues).\n\n---\n\n"
                    )

                    for ai_key in sorted(agreements_by_ai.keys(), key=natural_ai_sort_key):
                        ags = agreements_by_ai[ai_key]
                        ai_title = ags[0].get("ai_title", "") if ags else ""
                        title_str = f" — {ai_title}" if ai_title else ""
                        f.write(f"## Agenda Item {ai_key}{title_str}\n\n")
                        self._write_agreements_section(f, "", ags)

            # 5. Write Company Markdown File
            if company_docs or company_agreements:
                company_path = self.export_dir / "Company.md"
                with open(to_extended_path(company_path), "w", encoding="utf-8") as f:
                    f.write(header_comment)
                    if company_docs:
                        self._write_table(f, "Following Company Contributions:", company_docs, cols_company)
                    if company_agreements:
                        self._write_agreements_section(
                            f,
                            "Following Company Contributions resulted in Agreements / Working Assumptions:",
                            company_agreements,
                        )

            ag_note = f"\n• Included {len(all_agreements)} Chairman's Notes agreement(s)" if all_agreements else ""
            self.finished.emit(True, f"Successfully exported Markdown reports to:\n{self.export_dir}{ag_note}")

        except Exception as e:
            logger.error(f"[MarkdownExporter] Export failed: {e}", exc_info=True)
            self.finished.emit(False, str(e))

    def _write_table(self, f, title, docs, columns):
        if not docs:
            return

        if title:
            f.write(f"{title}\n\n")

        # Write headers
        f.write("| " + " | ".join(columns) + " |\n")
        # Ensure left-aligned markdown columns
        f.write("|" + "|".join([":---"] * len(columns)) + "|\n")

        # Write rows
        for d in docs:
            row_data = []
            for col in columns:
                # Catch empty values and safely escape newlines and markdown pipes
                val = str(d.get(col, '')).replace('\n', ' ').replace('\r', '').replace('|', '\\|')

                # Format Source table cell with carry-over separation if unconfirmed sources exist
                if col == "Source" and val:
                    val = CompanySanitizer.format_source_markdown_table_cell(d.get(col, ''))
                # Render the TDoc URL hyperlink natively
                elif col == "TDoc" and val:
                    tdoc_url = resolve_tdoc_download_url(val, docs_url=self.docs_url, meta=d)
                    if tdoc_url:
                        val = f"[{val}]({tdoc_url})"

                row_data.append(val)

            f.write("| " + " | ".join(row_data) + " |\n")
        f.write("\n")

    def _write_agreements_section(self, f, title: str, agreements: list):
        """Renders Chairman's Notes agreements cleanly into Markdown with icons and callout quotes."""
        if not agreements:
            return

        if title:
            f.write(f"### ⚖️ {title}\n\n")

        for ag in agreements:
            kind = str(ag.get("kind", "Agreement")).strip()
            ai = str(ag.get("ai", "")).strip()
            ai_title = str(ag.get("ai_title", "")).strip()
            tdoc_id = str(ag.get("tdoc_id", "")).strip().upper()
            wi = str(ag.get("wi", "")).strip()
            text = str(ag.get("text", "")).strip()

            icon = "🟢"
            if "ffs" in kind.lower() or "open issue" in kind.lower():
                icon = "🔵"
            elif "working assumption" in kind.lower():
                icon = "🟡"

            header_parts = [f"- {icon} **[{kind}]**"]
            if ai and title:  # only repeat AI if this is an AI-spanning section
                header_parts.append(f"**AI {ai}**" + (f" (*{ai_title}*)" if ai_title else ""))
            if wi and wi != "General":
                header_parts.append(f"*(WI: `{wi}`)*")
            if tdoc_id:
                tdoc_url = resolve_tdoc_download_url(tdoc_id, docs_url=self.docs_url)
                link_str = f"[{tdoc_id}]({tdoc_url})" if tdoc_url else f"`{tdoc_id}`"
                header_parts.append(f"Related TDoc: {link_str}")

            f.write(" — ".join(header_parts) + "\n")
            for subline in text.split("\n"):
                f.write(f"  > {subline}\n")
            f.write("\n")
