# --- File: modules/meetings/core/tdocs_parser.py ---
import io
import re
import openpyxl
import logging


class TDocsParser:
    @staticmethod
    def parse_tdocs_excel(filepath: str) -> list:
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

                        # Safely isolate the true "Agenda Item" column (Ignores 'Agenda item sort order')
                        if ("AGENDA ITEM" in val_up or val_up in ["AI", "AI#",
                                                                  "AI #"]) and "SORT" not in val_up and "DESCRIPTION" not in val_up:
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
            return data

        except Exception as e:
            logging.error(f"Failed to parse Excel file {filepath}: {e}")
            return []