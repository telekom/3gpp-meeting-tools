# --- File: src/modules/meetings/core/markdown_exporter.py ---
import json
import os
import re
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal


class MarkdownExporterThread(QThread):
    finished = pyqtSignal(bool, str)

    def __init__(self, meeting_dir: Path, tdocs_data: list, docs_url: str, mtg_info: dict):
        super().__init__()
        self.meeting_dir = meeting_dir
        self.tdocs_data = tdocs_data
        self.docs_url = docs_url
        self.mtg_info = mtg_info
        self.export_dir = self.meeting_dir / "Export"

        # Reliably resolve the root `src/` directory
        src_dir = Path(__file__).resolve().parents[3]
        config_path = src_dir / "export_config.json"

        # Define the exact default configuration provided
        default_config = {
            "columns_for_3gu_tdoc_export": ["TDoc", "Agenda Item", "Type", "For", "Title", "Source", "Abstract"],
            "columns_for_3gu_tdoc_export_ls": ["TDoc", "Agenda Item", "Type", "For", "Title", "Source", "Abstract"],
            "columns_for_3gu_tdoc_export_ls_out": ["Abstract", "TDoc", "Title", "Agenda Item", "Reply to"],
            "columns_for_3gu_tdoc_export_pcr": ["TDoc", "Agenda Item", "Type", "For", "Title", "Source", "Abstract"],
            "columns_for_3gu_tdoc_export_cr": ["TDoc", "Agenda Item", "Type", "For", "Title", "Source", "Abstract"],
            "columns_for_3gu_tdoc_export_contributor": ["TDoc", "Type", "Title", "Source", "Agenda Item", "TDoc Status",
                                                        "Spec", "CR category"],
            "company_name_regex_for_report": "Deutsche Telekom"
        }

        self.config = {}

        # Check if the config exists. If not, generate it so it works out-of-the-box!
        if not config_path.exists():
            try:
                with open(config_path, "w", encoding="utf-8") as f:
                    json.dump(default_config, f, indent=4)
                self.config = default_config
            except Exception as e:
                print(f"Failed to create default config file: {e}")
                self.config = default_config
        else:
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    self.config = json.load(f)
            except Exception as e:
                print(f"Failed to load config, falling back to defaults: {e}")
                self.config = default_config

    def run(self):
        try:
            self.export_dir.mkdir(parents=True, exist_ok=True)

            # Load column configurations with fallbacks to the defaults
            cols_ls_in = self.config.get("columns_for_3gu_tdoc_export_ls",
                                         ["TDoc", "Agenda Item", "Type", "For", "Title", "Source", "Abstract"])
            cols_ls_out = self.config.get("columns_for_3gu_tdoc_export_ls_out",
                                          ["Abstract", "TDoc", "Title", "Agenda Item", "Reply to"])
            cols_pcr = self.config.get("columns_for_3gu_tdoc_export_pcr",
                                       ["TDoc", "Agenda Item", "Type", "For", "Title", "Source", "Abstract"])
            cols_cr = self.config.get("columns_for_3gu_tdoc_export_cr",
                                      ["TDoc", "Agenda Item", "Type", "For", "Title", "Source", "Abstract"])
            cols_company = self.config.get("columns_for_3gu_tdoc_export_contributor",
                                           ["TDoc", "Title", "Source", "Agenda Item", "TDoc Status", "Spec",
                                            "CR category"])
            company_regex = self.config.get("company_name_regex_for_report", "Deutsche Telekom")

            # Extract Meeting info for the Markdown Header
            wg_name = str(self.mtg_info.get('wg_name', 'WG'))
            mtg_num = str(self.mtg_info.get('meeting_number', ''))
            mtg_url = self.docs_url.replace('/Docs', '').replace('/docs', '').rstrip('/') if self.docs_url else ""
            header_comment = f"<!--- [{wg_name}#{mtg_num}]({mtg_url}) --->\n\n"

            ai_groups = {}
            company_docs = []

            # 1. Parse and categorize the data
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

                # Add to Company Report if it matches regex and is not withdrawn
                if re.search(company_regex, source, re.IGNORECASE) and not is_withdrawn:
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

            # 2. Write the Agenda Item Markdown Files
            for ai, groups in ai_groups.items():
                if not any(groups.values()):
                    continue

                safe_ai = re.sub(r'[\\/*?:"<>|]', "_", ai)
                filepath = self.export_dir / f"{safe_ai}.md"

                with open(filepath, "w", encoding="utf-8") as f:
                    f.write(header_comment)

                    self._write_table(f, "Following LS were received:", groups['ls_in'], cols_ls_in)
                    self._write_table(f, "Following LS OUTs were agreed:", groups['ls_out'], cols_ls_out)
                    self._write_table(f, "Following pCRs were agreed:", groups['pcr'], cols_pcr)
                    self._write_table(f, "Following CRs were agreed:", groups['cr'], cols_cr)

            # 3. Write the Company Markdown File
            if company_docs:
                company_path = self.export_dir / "Company.md"
                with open(company_path, "w", encoding="utf-8") as f:
                    f.write(header_comment)
                    self._write_table(f, "Following Company Contributions:", company_docs, cols_company)

            self.finished.emit(True, f"Successfully exported Markdown reports to:\n{self.export_dir}")

        except Exception as e:
            self.finished.emit(False, str(e))

    def _write_table(self, f, title, docs, columns):
        if not docs:
            return

        if title:
            f.write(f"{title}\n\n")

        # Write headers
        f.write("| " + " | ".join(columns) + " |\n")
        # Ensure left-aligned markdown columns matching your examples
        f.write("|" + "|".join([":---"] * len(columns)) + "|\n")

        # Write rows
        for d in docs:
            row_data = []
            for col in columns:
                # Catch empty values and safely escape newlines and markdown pipes
                val = str(d.get(col, '')).replace('\n', ' ').replace('\r', '').replace('|', '\\|')

                # Render the TDoc URL hyperlink natively
                if col == "TDoc" and val and self.docs_url:
                    base_url = self.docs_url.rstrip('/')
                    val = f"[{val}]({base_url}/{val}.zip)"

                row_data.append(val)

            f.write("| " + " | ".join(row_data) + " |\n")
        f.write("\n")