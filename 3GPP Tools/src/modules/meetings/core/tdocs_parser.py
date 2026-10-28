# --- File: src/modules/meetings/core/tdocs_parser.py ---
import io
import re
import openpyxl
import logging
import json
import os
from pathlib import Path
import docx

logger = logging.getLogger(__name__)


class TDocsParser:
    @staticmethod
    def parse_tdocs_excel(filepath: str) -> list:
        json_cache = filepath + ".json"
        try:
            if os.path.exists(json_cache) and os.path.getmtime(json_cache) >= os.path.getmtime(filepath):
                with open(json_cache, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception as e:
            logging.warning(f"Could not read JSON cache: {e}")

        data = []
        try:
            with open(filepath, "rb") as f:
                in_mem_file = io.BytesIO(f.read())

            wb = openpyxl.load_workbook(in_mem_file, data_only=True, read_only=True)
            sheet = wb["TDoc_List"] if "TDoc_List" in wb.sheetnames else wb.worksheets[0]

            headers = []
            header_row_idx = 1

            for row_idx, row in enumerate(sheet.iter_rows(min_row=1, max_row=15, values_only=True), start=1):
                row_strs = [str(c).strip() if c is not None else "" for c in row]

                hits = 0
                for c in row_strs:
                    cu = c.upper()
                    if cu in ["TDOC", "TD#", "TDOC#"]: hits += 1
                    if cu == "TITLE": hits += 1
                    if cu == "SOURCE": hits += 1
                    if cu == "TYPE": hits += 1
                    if cu == "FOR": hits += 1
                    if "AGENDA ITEM" in cu or cu in ["AI", "AI#", "AI #"]: hits += 1
                    if "STATUS" in cu: hits += 1

                if hits >= 3:
                    header_row_idx = row_idx
                    for val in row_strs:
                        val_clean = re.sub(r'\s+', ' ', val).strip()
                        val_up = val_clean.upper()

                        if ("AGENDA ITEM" in val_up or val_up in ["AI", "AI#",
                                                                  "AI #"]) and "SORT" not in val_up and "DESCRIPTION" not in val_up:
                            val = "Agenda Item"
                        elif val_up in ["TD#", "TDOC#", "TDOC"]:
                            val = "TDoc"
                        # ---> THE FIX: Force the Excel column to map to "TDoc Status"
                        elif "STATUS" in val_up or "RESULT" in val_up:
                            val = "TDoc Status"

                        headers.append(val)
                    break

            if not headers:
                logging.warning("Could not find a valid header row in the TDocs Excel file.")
                wb.close()
                return []

            for row in sheet.iter_rows(min_row=header_row_idx + 1, values_only=True):
                row_dict = {}
                is_empty = True
                for i, value in enumerate(row):
                    if i < len(headers) and headers[i]:
                        val_str = str(value).strip() if value is not None else ""
                        row_dict[headers[i]] = val_str
                        if val_str: is_empty = False

                if not is_empty:
                    data.append(row_dict)

            wb.close()

            try:
                with open(json_cache, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2)
            except Exception as e:
                logging.warning(f"Could not save JSON cache: {e}")

            return data

        except Exception as e:
            logging.error(f"Failed to parse Excel file {filepath}: {e}")
            return []

    @classmethod
    def parse_tdocs_by_agenda(cls, filepath: str, ui_logger=None) -> dict:
        from bs4 import BeautifulSoup
        import logging

        if ui_logger: logger.log(logging.INFO, "⏳ Parsing TdocsByAgenda HTML (Word Export)...")
        data = {}

        try:
            with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
                soup = BeautifulSoup(f, 'html.parser')

            tables = soup.find_all('table')
            if not tables:
                if ui_logger: logger.log(logging.ERROR, "❌ No tables found in HTML.")
                return data

            target_table = None
            for table in tables:
                headers = [th.get_text(strip=True).lower() for th in table.find_all(['th', 'td'])[:20]]
                if any('td#' in h or 'td #' in h for h in headers):
                    target_table = table
                    break

            if not target_table:
                if ui_logger: logger.log(logging.ERROR, "❌ Could not identify the main TDoc table in HTML.")
                return data

            rows = target_table.find_all('tr')
            if not rows: return data

            header_row = rows[0].find_all(['th', 'td'])
            headers = [h.get_text(separator=' ', strip=True).lower() for h in header_row]

            td_idx = next((i for i, h in enumerate(headers) if 'td#' in h or 'td #' in h), -1)
            comments_idx = next((i for i, h in enumerate(headers) if 'comments' in h), -1)
            email_idx = next((i for i, h in enumerate(headers) if 'e-mail_discussion' in h), -1)
            source_idx = next((i for i, h in enumerate(headers) if 'source' in h), -1)

            # ---> THE FIX: Smarter hunting for HTML columns
            title_idx = next((i for i, h in enumerate(headers) if 'title' in h or 'subject' in h), -1)
            for_idx = next((i for i, h in enumerate(headers) if h == 'for' or 'doc for' in h), -1)
            result_idx = next((i for i, h in enumerate(headers) if 'result' in h or 'status' in h), -1)
            type_idx = next((i for i, h in enumerate(headers) if 'type' in h), -1)

            if td_idx == -1:
                if ui_logger: logger.log(logging.ERROR, "❌ 'TD#' column missing in HTML table.")
                return data

            for row in rows[1:]:
                cols = row.find_all(['td'])
                if len(cols) <= td_idx: continue

                tdoc_id = cols[td_idx].get_text(separator=' ', strip=True)
                if not tdoc_id or not tdoc_id.startswith(('S2-', 'R', 'C', 'S')):
                    continue

                comments = cols[comments_idx].get_text(separator='\n', strip=True) if comments_idx != -1 and len(
                    cols) > comments_idx else ""
                email_disc = cols[email_idx].get_text(separator='\n', strip=True) if email_idx != -1 and len(
                    cols) > email_idx else ""
                title = cols[title_idx].get_text(separator='\n', strip=True) if title_idx != -1 and len(
                    cols) > title_idx else ""
                source = cols[source_idx].get_text(separator='\n', strip=True) if source_idx != -1 and len(
                    cols) > source_idx else ""
                doc_for = cols[for_idx].get_text(separator='\n', strip=True) if for_idx != -1 and len(
                    cols) > for_idx else ""
                result = cols[result_idx].get_text(separator='\n', strip=True) if result_idx != -1 and len(
                    cols) > result_idx else ""
                doc_type = cols[type_idx].get_text(separator='\n', strip=True) if type_idx != -1 and len(
                    cols) > type_idx else ""

                if ui_logger and (comments or email_disc):
                    logger.log(logging.DEBUG, f"   ➔ Extracted agenda remarks for {tdoc_id}")

                data[tdoc_id] = {
                    'Comments': comments,
                    'e-mail_Discussion': email_disc,
                    'Title': title,
                    'Source': source,
                    'For': doc_for,
                    'Result': result,
                    'Type': doc_type
                }

            if ui_logger: logger.log(logging.INFO, f"✅ Successfully parsed {len(data)} TDocs from Agenda HTML.")

        except Exception as e:
            if ui_logger: logger.log(logging.ERROR, f"❌ Error parsing HTML: {str(e)}")

        return data

    @staticmethod
    def parse_tdocs_from_docx(docx_path: str) -> dict:
        """
        Parses 3GPP document lists and Chairman's Notes from a .docx table.
        Robustly accommodates diverse Working Group (SA2, RAN1/2/3, SA1/3, CT1/3/4)
        secretary formats, column naming variations, and multi-row header layouts.
        Returns a dictionary compatible with TDocsTableModel.merge_agenda_data().
        """
        path = Path(docx_path)
        if not path.exists():
            logging.error(f"[TDocsParser] File not found: {docx_path}")
            return {}

        try:
            doc = docx.Document(str(path))
        except Exception as e:
            logging.error(f"[TDocsParser] Failed to open docx file '{docx_path}': {e}")
            return {}

        agenda_data = {}

        def _map_header_row(header_cells: list) -> dict:
            col_map = {}
            for idx, text in enumerate(header_cells):
                clean_text = re.sub(r'\s+', ' ', text).strip().lower()
                if not clean_text:
                    continue

                # 1. TDoc identifier
                if (
                    clean_text in [
                        'td#', 'td #', 'tdoc', 'tdoc#', 'tdoc #', 'td no', 'td no.',
                        'td number', 'tdoc no', 'tdoc no.', 'tdoc number', 'temporary document',
                        't-doc', 't-doc#', 't-doc no', 't-doc no.', 'doc#', 'doc #', 'doc no',
                        'doc no.', 'document #', 'document no', 'document no.'
                    ]
                    or clean_text in ['td', 'doc']
                    or re.search(r'^(?:t(?:emp(?:orary)?)?\s*doc(?:ument)?|td)\s*(?:#|no\.?|num(?:ber)?)?$', clean_text)
                ):
                    col_map['tdoc'] = idx

                # 2. Title / Subject
                elif (
                    clean_text in ['subject', 'title', 'document title', 'tdoc title', 'subject / title', 'description', 'name']
                    or 'title' in clean_text
                    or 'subject' in clean_text
                ):
                    col_map['title'] = idx

                # 3. Source / Sponsoring Companies
                elif (
                    clean_text in ['source', 'sources', 'company', 'companies', 'source / company', 'proposer', 'proposers', 'submitted by']
                    or 'source' in clean_text
                    or 'proposer' in clean_text
                ):
                    col_map['source'] = idx

                # 4. Document Type
                elif clean_text in ['type', 'doc type', 'document type', 'kind']:
                    col_map['type'] = idx

                # 5. Document For / Purpose
                elif clean_text in ['doc for', 'for', 'for action', 'action', 'purpose']:
                    col_map['for'] = idx

                # 6. Comments / Discussion / Chair Remarks
                elif (
                    clean_text in [
                        'comments', 'comment', 'secretary remarks', 'discussion', 'discussions',
                        'discussion / comments', 'comments / discussion', 'chair notes', "chair's notes",
                        'chairman notes', "chairman's notes", 'session remarks', 'remarks', 'notes',
                        'chair comments', 'summary', 'minutes', 'offline discussion'
                    ]
                    or 'discussion' in clean_text
                    or 'remarks' in clean_text
                    or 'chair' in clean_text
                ):
                    col_map['comments'] = idx

                # 7. Result / Decision / Status
                elif (
                    clean_text in [
                        'result', 'results', 'status', 'decision', 'decisions', 'outcome',
                        'outcomes', 'conclusion', 'conclusions', 'final status', 'action taken'
                    ]
                    or 'result' in clean_text
                    or 'status' in clean_text
                    or 'decision' in clean_text
                    or 'conclusion' in clean_text
                ):
                    col_map['result'] = idx

                # 8. Agenda Item
                elif (
                    clean_text in [
                        'ai', 'agenda item', 'agenda', 'item', 'agenda item (ai)', 'agenda-item',
                        'ai#', 'ai #', 'agenda no', 'agenda no.', 'agenda item #'
                    ]
                    or clean_text.startswith('agenda item')
                    or clean_text.startswith('ai ')
                    or clean_text == 'ai'
                ):
                    col_map['ai'] = idx

            return col_map

        for table in doc.tables:
            if len(table.rows) < 2:
                continue

            # Identify header row by scanning the first 5 rows (accommodating title/metadata banners)
            col_map = {}
            header_row_idx = 0
            max_scan = min(5, len(table.rows))

            for r_idx in range(max_scan):
                candidate_cells = [cell.text for cell in table.rows[r_idx].cells]
                candidate_map = _map_header_row(candidate_cells)
                # Valid header row must identify the TDoc column and at least one other metadata column
                if 'tdoc' in candidate_map and len(candidate_map) >= 2:
                    col_map = candidate_map
                    header_row_idx = r_idx
                    break

            # If no multi-column header found, check if row 0 has at least a tdoc column
            if not col_map:
                candidate_cells = [cell.text for cell in table.rows[0].cells]
                candidate_map = _map_header_row(candidate_cells)
                if 'tdoc' in candidate_map:
                    col_map = candidate_map
                    header_row_idx = 0

            if not col_map or 'tdoc' not in col_map:
                continue

            current_ai = 'N/A'

            for row in table.rows[header_row_idx + 1:]:
                cells = [c.text.strip() for c in row.cells]
                if not cells or len(cells) <= col_map['tdoc']:
                    continue

                raw_tdoc = cells[col_map['tdoc']].strip()
                if not raw_tdoc or raw_tdoc == '-':
                    # Check for section break / Agenda Item banner in rows without TDocs
                    full_row_text = " ".join(cells).strip()
                    ai_banner_match = re.search(
                        r'(?:agenda\s*item|agenda\s*no\.?|ai)\s*[:#]?\s*([0-9]+(?:\.[0-9]+)*[a-zA-Z]?)',
                        full_row_text,
                        re.IGNORECASE,
                    )
                    if ai_banner_match:
                        current_ai = ai_banner_match.group(1).strip()
                    continue

                # Extract canonical 3GPP TDoc ID (e.g. S2-2401234, R2-2405678r01, SP-230890)
                td_match = re.search(r'\b([A-Za-z0-9]{1,4}-\d{4,8}(?:[rR]\d{1,2}[a-zA-Z]?)?)\b', raw_tdoc)
                if not td_match:
                    # Fallback to general prefix-digits pattern
                    td_match = re.search(r'([A-Za-z0-9]+-\d+)', raw_tdoc)
                if not td_match:
                    # Not a TDoc row; check if this row defines an Agenda Item section banner
                    full_row_text = " ".join(cells).strip()
                    ai_banner_match = re.search(
                        r'(?:agenda\s*item|agenda\s*no\.?|ai)\s*[:#]?\s*([0-9]+(?:\.[0-9]+)*[a-zA-Z]?)',
                        full_row_text,
                        re.IGNORECASE,
                    )
                    if ai_banner_match:
                        current_ai = ai_banner_match.group(1).strip()
                    continue

                tdoc_id = td_match.group(1).upper()

                # Track current Agenda Item across multi-row sub-headers
                if 'ai' in col_map and col_map['ai'] < len(cells):
                    ai_val = cells[col_map['ai']].strip()
                    if ai_val and ai_val != '-':
                        current_ai = ai_val

                title = cells[col_map['title']] if 'title' in col_map and col_map['title'] < len(cells) else ''
                source = cells[col_map['source']] if 'source' in col_map and col_map['source'] < len(cells) else ''
                doc_type = cells[col_map['type']] if 'type' in col_map and col_map['type'] < len(cells) else ''
                doc_for = cells[col_map['for']] if 'for' in col_map and col_map['for'] < len(cells) else ''
                comments = cells[col_map['comments']] if 'comments' in col_map and col_map['comments'] < len(cells) else ''
                result = cells[col_map['result']] if 'result' in col_map and col_map['result'] < len(cells) else ''

                if tdoc_id in agenda_data:
                    # Merge sequential entries (e.g. morning vs afternoon session notes)
                    existing = agenda_data[tdoc_id]
                    if comments and comments != existing.get('Comments', ''):
                        if existing.get('Comments'):
                            comments = f"{existing['Comments']}\n---\n{comments}"
                    else:
                        comments = comments or existing.get('Comments', '')

                    result = result or existing.get('Result', '')
                    title = title or existing.get('Title', '')
                    source = source or existing.get('Source', '')
                    doc_type = doc_type or existing.get('Type', '')
                    doc_for = doc_for or existing.get('For', '')
                    if current_ai == 'N/A' and existing.get('Agenda Item'):
                        current_ai = existing.get('Agenda Item')

                agenda_data[tdoc_id] = {
                    'Title': title,
                    'Source': source,
                    'Type': doc_type,
                    'For': doc_for,
                    'Comments': comments,
                    'Result': result,
                    'Agenda Item': current_ai,
                    'e-mail_Discussion': ''
                }

        logging.info(f"[TDocsParser] Successfully parsed {len(agenda_data)} TDocs from {path.name}")
        return agenda_data