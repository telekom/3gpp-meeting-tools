from typing import NamedTuple, Dict

import openpyxl
import openpyxl.utils
from openpyxl.utils.cell import coordinate_from_string


def parse_excel_hyperlinks_by_column_name(file_path, column_name, sheet_name=None):
    """
    Parses hyperlinks from a specific column in an Excel (.xlsx) file,
    identified by its column name (header).

    Args:
        file_path (str): The path to the Excel file.
        column_name (str): The name of the column (header) to parse.
        sheet_name (str, optional): The name of the sheet to parse.
                                     If None, it will parse the active sheet.

    Returns:
        list: A list of tuples, where each tuple contains (hyperlink_text, hyperlink_target).
              Returns an empty list if no hyperlinks are found or an error occurs.
    """
    hyperlinks_info = set()
    column_letter = None

    try:
        workbook = openpyxl.load_workbook(file_path)

        if sheet_name:
            if sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
            else:
                print(f"Error: Sheet '{sheet_name}' not found in the workbook.")
                return hyperlinks_info
        else:
            sheet = workbook.active

        print(f"Parsing column '{column_name}' in sheet: {sheet.title}")

        # 1. Find the column index for the given column_name
        col_index = None
        for col_idx, cell in enumerate(sheet[1]):
            if cell.value == column_name:
                col_index = col_idx
                break

        if col_index is None:
            print(f"Error: Column '{column_name}' not found in the header row.")
            return hyperlinks_info

        try:
            column_letter = openpyxl.utils.get_column_letter(col_index + 1)
        except ValueError:
            print(f"Error: Column '{column_name}' not found in the header row.")
            return hyperlinks_info

        print(f"Column '{column_name}' corresponds to column letter '{column_letter}'.")

        # 2. Iterate directly over the sheet's hyperlinks collection
        for hl in sheet.hyperlinks:
            if not hl.ref:
                continue

            # hl.ref can sometimes be a merged range like 'C5:C10'.
            # Splitting by ':' and taking the first part guarantees we get the top-left cell ('C5')
            top_left_coord = hl.ref.split(':')[0]

            # Robustly separate the letter and the number (e.g., 'C' and 5)
            hl_col, hl_row = coordinate_from_string(top_left_coord)

            # STRICT FILTER: Only process if it matches our target column AND is not the header row
            if hl_col == column_letter and hl_row > 1:
                # We only instantiate the cell object in Python if it passes the filter
                cell_value = sheet[top_left_coord].value

                if hl.target:
                    # 3. Use .add() instead of .append().
                    # If it's a duplicate, the set silently ignores it in O(1) time.
                    hyperlinks_info.add((cell_value, hl.target))

    except FileNotFoundError:
        print(f"Error: File not found at {file_path}")
    except Exception as e:
        print(f"An error occurred: {e}")

    # Remove duplicates by converting to a set and then back to a list
    unique_hyperlinks = list(hyperlinks_info)

    # Sort the list of tuples alphabetically by the hyperlink's text
    unique_hyperlinks.sort(key=lambda x: str(x[0]) if x[0] is not None else "")

    return unique_hyperlinks

class WiData(NamedTuple):
    wi_name:str
    wi_url:str

def parse_tdoc_3gu_list_for_wis(file_path) -> Dict[str,str]:
    hyperlinks_list = parse_excel_hyperlinks_by_column_name(file_path=file_path, column_name='Related WIs')
    return dict(hyperlinks_list)


