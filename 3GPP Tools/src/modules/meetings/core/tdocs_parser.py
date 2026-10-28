# --- File: src/modules/meetings/core/tdocs_parser.py ---
import io
import re
import html
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

        def _clean_text(val: str, preserve_newlines: bool = False) -> str:
            if not val:
                return ""
            s = str(val).strip()
            # Double unescape to handle nested entities like &amp;apos; or &amp;amp;
            s = html.unescape(s)
            s = html.unescape(s)
            s = s.replace('&apos;', "'").replace('&amp;', '&').replace('&#39;', "'")
            if preserve_newlines:
                lines = [re.sub(r'[ \t]+', ' ', l).strip() for l in s.split('\n')]
                return '\n'.join(lines).strip()
            return re.sub(r'\s+', ' ', s).strip()

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
            # Match both 'comment' and 'comments' (e.g. SA2#164 uses singular 'comment')
            comments_idx = next((i for i, h in enumerate(headers) if 'comment' in h), -1)
            email_idx = next((i for i, h in enumerate(headers) if 'e-mail_discussion' in h), -1)
            source_idx = next((i for i, h in enumerate(headers) if 'source' in h), -1)

            # Smarter hunting for HTML columns
            title_idx = next((i for i, h in enumerate(headers) if 'title' in h or 'subject' in h), -1)
            for_idx = next((i for i, h in enumerate(headers) if h == 'for' or 'doc for' in h), -1)
            result_idx = next((i for i, h in enumerate(headers) if 'result' in h or 'status' in h), -1)
            type_idx = next((i for i, h in enumerate(headers) if 'type' in h), -1)
            ai_idx = next((i for i, h in enumerate(headers) if h in ('ai', 'ai#', 'ai #', 'agenda', 'agenda item') or 'agenda item' in h), -1)
            wi_idx = next((i for i, h in enumerate(headers) if 'work item' in h or h == 'wi'), -1)
            rel_idx = next((i for i, h in enumerate(headers) if h in ('rel', 'release') or 'release' in h), -1)

            if td_idx == -1:
                if ui_logger: logger.log(logging.ERROR, "❌ 'TD#' column missing in HTML table.")
                return data

            current_ai = 'N/A'
            current_sub_item = ''
            sub_item_order_counter = 0
            order_in_agenda = 0

            for row in rows[1:]:
                cols = row.find_all(['td'])
                if len(cols) <= td_idx: continue

                raw_td = cols[td_idx].get_text(separator=' ', strip=True)
                raw_ai = cols[ai_idx].get_text(separator=' ', strip=True) if ai_idx != -1 and len(cols) > ai_idx else ""
                raw_title = cols[title_idx].get_text(separator=' ', strip=True) if title_idx != -1 and len(cols) > title_idx else ""

                # Check if this row is a banner or separator row (TD# is '-' or empty or not a valid TDoc)
                is_tdoc = bool(raw_td and raw_td != '-' and re.search(r'\b[A-Za-z0-9]{1,4}-\d{4,8}', raw_td))

                if not is_tdoc:
                    # Update active Agenda Item if provided
                    if raw_ai and raw_ai != '-' and raw_ai != '0':
                        clean_ai = _clean_text(raw_ai)
                        if clean_ai != current_ai:
                            current_ai = clean_ai
                            # The first row of a new AI is the Agenda Item header itself
                            current_sub_item = ''
                            continue

                    # If this separator row appears under an active AI, it defines a Sub-Item!
                    clean_subj = _clean_text(raw_title)
                    if clean_subj and clean_subj != '-' and not clean_subj.lower().startswith('all agenda items'):
                        current_sub_item = clean_subj
                        sub_item_order_counter += 1
                    continue

                # It's a valid TDoc row!
                tdoc_id = raw_td.strip()
                td_match = re.search(r'\b([A-Za-z0-9]{1,4}-\d{4,8}(?:[rR]\d{1,2}[a-zA-Z]?)?)\b', tdoc_id)
                if td_match:
                    tdoc_id = td_match.group(1).upper()

                order_in_agenda += 1

                # Update current_ai from row's AI column if present
                if raw_ai and raw_ai != '-' and raw_ai != '0':
                    current_ai = _clean_text(raw_ai)

                raw_comments = cols[comments_idx].get_text(separator='\n', strip=True) if comments_idx != -1 and len(cols) > comments_idx else ""
                comments = _clean_text(raw_comments, preserve_newlines=True)
                email_disc = _clean_text(cols[email_idx].get_text(separator='\n', strip=True) if email_idx != -1 and len(cols) > email_idx else "", preserve_newlines=True)
                title = _clean_text(raw_title)
                source = _clean_text(cols[source_idx].get_text(separator='\n', strip=True) if source_idx != -1 and len(cols) > source_idx else "")
                doc_for = _clean_text(cols[for_idx].get_text(separator='\n', strip=True) if for_idx != -1 and len(cols) > for_idx else "")
                result = _clean_text(cols[result_idx].get_text(separator='\n', strip=True) if result_idx != -1 and len(cols) > result_idx else "")
                doc_type = _clean_text(cols[type_idx].get_text(separator='\n', strip=True) if type_idx != -1 and len(cols) > type_idx else "")
                raw_wi = cols[wi_idx].get_text(separator=' ', strip=True) if wi_idx != -1 and len(cols) > wi_idx else ""
                work_item = _clean_text(raw_wi) if raw_wi and raw_wi != '-' else ""
                raw_rel = cols[rel_idx].get_text(separator=' ', strip=True) if rel_idx != -1 and len(cols) > rel_idx else ""
                rel = _clean_text(raw_rel) if raw_rel and raw_rel != '-' else ""

                if ui_logger and (comments or email_disc):
                    logger.log(logging.DEBUG, f"   ➔ Extracted agenda remarks for {tdoc_id}")

                data[tdoc_id] = {
                    'Comments': comments,
                    'e-mail_Discussion': email_disc,
                    'Title': title,
                    'Source': source,
                    'For': doc_for,
                    'Result': result,
                    'Type': doc_type,
                    'Agenda Item': current_ai if current_ai != 'N/A' else '',
                    'Sub-Item': current_sub_item,
                    'Related WIs': work_item,
                    'Release': rel,
                    '_agenda_order': order_in_agenda,
                    '_sub_item_order': sub_item_order_counter
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
            current_sub_item = ''
            sub_item_order_counter = 0
            order_in_agenda = 0

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
                        current_sub_item = ''
                    else:
                        title_candidate = cells[col_map['title']] if 'title' in col_map and col_map['title'] < len(cells) else ''
                        title_candidate = title_candidate.strip()
                        if title_candidate and title_candidate != '-' and not title_candidate.lower().startswith('all agenda'):
                            current_sub_item = title_candidate
                            sub_item_order_counter += 1
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
                        current_sub_item = ''
                    else:
                        title_candidate = cells[col_map['title']] if 'title' in col_map and col_map['title'] < len(cells) else ''
                        title_candidate = title_candidate.strip()
                        if title_candidate and title_candidate != '-' and not title_candidate.lower().startswith('all agenda'):
                            current_sub_item = title_candidate
                            sub_item_order_counter += 1
                    continue

                tdoc_id = td_match.group(1).upper()
                order_in_agenda += 1

                # Track current Agenda Item across multi-row sub-headers
                if 'ai' in col_map and col_map['ai'] < len(cells):
                    ai_val = cells[col_map['ai']].strip()
                    if ai_val and ai_val != '-':
                        if ai_val != current_ai:
                            current_ai = ai_val
                            current_sub_item = ''

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
                    if not current_sub_item and existing.get('Sub-Item'):
                        current_sub_item = existing.get('Sub-Item')

                agenda_data[tdoc_id] = {
                    'Title': title,
                    'Source': source,
                    'Type': doc_type,
                    'For': doc_for,
                    'Comments': comments,
                    'Result': result,
                    'Agenda Item': current_ai,
                    'Sub-Item': current_sub_item,
                    '_agenda_order': order_in_agenda,
                    '_sub_item_order': sub_item_order_counter,
                    'e-mail_Discussion': ''
                }

        # Fallback for Working Groups that format Chairman's Notes as tab-delimited text paragraphs rather than Word tables (e.g. RAN1, RAN2)
        if not agenda_data and doc.paragraphs:
            agenda_data = TDocsParser._parse_tdocs_from_paragraphs(doc)

        logging.info(f"[TDocsParser] Successfully parsed {len(agenda_data)} TDocs from {path.name}")
        return agenda_data

    @staticmethod
    def _parse_tdocs_from_paragraphs(doc: docx.Document) -> dict:
        """
        Parses Chairman's Notes formatted as structured tab-delimited text paragraphs
        (used by RAN1, RAN2, and other Working Groups that do not use Word tables).
        """
        agenda_data = {}
        current_ai = 'N/A'
        current_tdoc = None

        tdoc_line_re = re.compile(
            r'^([A-Z][A-Z0-9]{1,4}-\d{4,8}(?:[rR]\d{1,2}[a-zA-Z]?)?)\s*\t+(.*)$'
        )
        heading_ai_re = re.compile(
            r'^(?:agenda\s*item\s*)?([0-9]+(?:\.[0-9]+)*)\s*(.*)$', re.IGNORECASE
        )

        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue

            style = p.style.name.lower() if p.style else ''

            # Track Agenda Item from headings (e.g. "Heading 1", "Heading 2") or explicit AI lines
            if 'heading' in style or text.lower().startswith(('agenda item', 'ai ')):
                hm = heading_ai_re.match(text)
                if hm:
                    current_ai = hm.group(1).strip()
                current_tdoc = None
                continue

            # Check if line starts with a TDoc ID followed by tab-delimited fields
            tm = tdoc_line_re.match(text)
            if tm:
                tdoc_id = tm.group(1).upper()
                rest = tm.group(2).split('\t')
                parts = [part.strip() for part in rest if part.strip()]
                title = parts[0] if len(parts) > 0 else ''
                source = parts[1] if len(parts) > 1 else ''
                doc_type = parts[2] if len(parts) > 2 else ''

                agenda_data[tdoc_id] = {
                    'Title': title,
                    'Source': source,
                    'Type': doc_type,
                    'For': '',
                    'Comments': '',
                    'Result': '',
                    'Agenda Item': current_ai,
                    'e-mail_Discussion': ''
                }
                current_tdoc = tdoc_id
                continue

            # Capture subsequent paragraphs as Agreement/Status or Comments for the current contribution
            if current_tdoc and current_tdoc in agenda_data:
                lower_text = text.lower()
                is_status = (
                    'agreement' in style or 'decision' in style or 'status' in style
                    or lower_text.startswith((
                        'agreed', 'approved', 'noted', 'revised to', 'revised into',
                        'postponed', 'withdrawn', 'endorsed', 'decision:', 'status:', 'result:'
                    ))
                )
                if is_status:
                    if not agenda_data[current_tdoc]['Result']:
                        agenda_data[current_tdoc]['Result'] = text
                    else:
                        agenda_data[current_tdoc]['Result'] += f" | {text}"
                else:
                    existing_comm = agenda_data[current_tdoc]['Comments']
                    if existing_comm:
                        agenda_data[current_tdoc]['Comments'] = f"{existing_comm}\n{text}"
                    else:
                        agenda_data[current_tdoc]['Comments'] = text

        return agenda_data