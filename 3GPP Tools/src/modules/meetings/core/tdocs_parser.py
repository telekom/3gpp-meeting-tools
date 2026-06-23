# --- File: modules/meetings/core/tdocs_parser.py ---
import io
import openpyxl
import logging


class TDocsParser:
    @staticmethod
    def parse_tdocs_excel(filepath: str) -> list:
        data = []
        try:
            # FIXED 1: Load file entirely into RAM to instantly release the OS file lock!
            with open(filepath, "rb") as f:
                in_mem_file = io.BytesIO(f.read())

            wb = openpyxl.load_workbook(in_mem_file, data_only=True, read_only=True)
            sheet = wb["TDoc_List"] if "TDoc_List" in wb.sheetnames else wb.worksheets[0]

            headers = []
            for cell in sheet[1]:
                val = str(cell.value).strip() if cell.value else ""
                # FIXED 2: Map the server's "AI" column to our "Agenda Item"
                if val == "AI":
                    val = "Agenda Item"
                headers.append(val)

            for row in sheet.iter_rows(min_row=2, values_only=True):
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