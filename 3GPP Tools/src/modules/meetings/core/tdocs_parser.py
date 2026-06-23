# --- File: modules/meetings/core/tdocs_parser.py ---
import openpyxl
import logging


class TDocsParser:
    @staticmethod
    def parse_tdocs_excel(filepath: str) -> list:
        """Parses the 3GPP TDocs Excel file and returns a list of dictionaries."""
        data = []
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
            # Default to the first sheet if "TDoc_List" isn't explicitly found
            sheet = wb["TDoc_List"] if "TDoc_List" in wb.sheetnames else wb.worksheets[0]

            # Extract headers from the first row
            headers = []
            for cell in sheet[1]:
                headers.append(str(cell.value).strip() if cell.value else "")

            # Map rows to dictionaries
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