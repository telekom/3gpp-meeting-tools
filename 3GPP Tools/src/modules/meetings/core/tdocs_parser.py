# --- File: modules/meetings/core/tdocs_parser.py ---
import io
import re
import openpyxl
import logging
import json
import os


class TDocsParser:
    @staticmethod
    def parse_tdocs_excel(filepath: str) -> list:
        # ---> OPTIMIZATION 1: Lightning-fast JSON Caching
        # If we already parsed this exact Excel file previously, load the JSON instantly.
        json_cache = filepath + ".json"
        try:
            if os.path.exists(json_cache) and os.path.getmtime(json_cache) >= os.path.getmtime(filepath):
                with open(json_cache, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception as e:
            logging.warning(f"Could not read JSON cache: {e}")

        data = []
        try:
            # Load file entirely into RAM to instantly release the OS file lock
            with open(filepath, "rb") as f:
                in_mem_file = io.BytesIO(f.read())

            wb = openpyxl.load_workbook(in_mem_file, data_only=True, read_only=True)
            sheet = wb["TDoc_List"] if "TDoc_List" in wb.sheetnames else wb.worksheets[0]

            headers = []
            header_row_idx = 1

            # FIXED: Bulletproof header hunting.
            # We require at least 3 matching 3GPP columns to prove this is the real header row.
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

                        # Safely isolate the true "Agenda Item" column
                        if ("AGENDA ITEM" in val_up or val_up in ["AI", "AI#", "AI #"]) and "SORT" not in val_up and "DESCRIPTION" not in val_up:
                            val = "Agenda Item"
                        elif val_up in ["TD#", "TDOC#", "TDOC"]:
                            val = "TDoc"

                        headers.append(val)
                    break

            if not headers:
                logging.warning("Could not find a valid header row in the TDocs Excel file.")
                wb.close()
                return []

            # Parse the actual data starting exactly one row beneath the verified headers
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

            # ---> OPTIMIZATION 1: Save the cache for the next time!
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

        if ui_logger: ui_logger.emit("⏳ Parsing TdocsByAgenda HTML (Word Export)...", logging.INFO)
        data = {}

        try:
            with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
                soup = BeautifulSoup(f, 'html.parser')

            tables = soup.find_all('table')
            if not tables:
                if ui_logger: ui_logger.emit("❌ No tables found in HTML.", logging.ERROR)
                return data

            target_table = None
            for table in tables:
                headers = [th.get_text(strip=True).lower() for th in table.find_all(['th', 'td'])[:20]]
                if any('td#' in h or 'td #' in h for h in headers):
                    target_table = table
                    break

            if not target_table:
                if ui_logger: ui_logger.emit("❌ Could not identify the main TDoc table in HTML.", logging.ERROR)
                return data

            rows = target_table.find_all('tr')
            if not rows: return data

            header_row = rows[0].find_all(['th', 'td'])
            headers = [h.get_text(separator=' ', strip=True).lower() for h in header_row]

            td_idx = next((i for i, h in enumerate(headers) if 'td#' in h or 'td #' in h), -1)
            comments_idx = next((i for i, h in enumerate(headers) if 'comments' in h), -1)
            email_idx = next((i for i, h in enumerate(headers) if 'e-mail_discussion' in h), -1)
            title_idx = next((i for i, h in enumerate(headers) if 'title' in h), -1)
            source_idx = next((i for i, h in enumerate(headers) if 'source' in h), -1)

            if td_idx == -1:
                if ui_logger: ui_logger.emit("❌ 'TD#' column missing in HTML table.", logging.ERROR)
                return data

            for row in rows[1:]:
                cols = row.find_all(['td'])
                if len(cols) <= td_idx: continue

                tdoc_id = cols[td_idx].get_text(separator=' ', strip=True)
                if not tdoc_id or not tdoc_id.startswith(('S2-', 'R', 'C', 'S')):
                    continue

                comments = cols[comments_idx].get_text(separator=' ', strip=True) if comments_idx != -1 and len(
                    cols) > comments_idx else ""
                email_disc = cols[email_idx].get_text(separator=' ', strip=True) if email_idx != -1 and len(
                    cols) > email_idx else ""
                title = cols[title_idx].get_text(separator=' ', strip=True) if title_idx != -1 and len(
                    cols) > title_idx else ""
                source = cols[source_idx].get_text(separator=' ', strip=True) if source_idx != -1 and len(
                    cols) > source_idx else ""

                if ui_logger and (comments or email_disc):
                    ui_logger.emit(f"   ➔ Extracted agenda remarks for {tdoc_id}", logging.DEBUG)

                data[tdoc_id] = {
                    'Comments': comments,
                    'e-mail_Discussion': email_disc,
                    'Title': title,
                    'Source': source
                }

            if ui_logger: ui_logger.emit(f"✅ Successfully parsed {len(data)} TDocs from Agenda HTML.", logging.INFO)

        except Exception as e:
            if ui_logger: ui_logger.emit(f"❌ Error parsing HTML: {str(e)}", logging.ERROR)

        return data