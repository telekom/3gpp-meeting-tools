from typing import NamedTuple, Dict

import openpyxl
import openpyxl.utils


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
    hyperlinks_info = []  # Change to a list to store tuples
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

        # 1. Find the column letter based on the column name (header)
        header_row = [cell.value for cell in sheet[1]]  # Assuming headers are in the first row

        try:
            # Get the 0-indexed position of the column name
            col_index = header_row.index(column_name)
            # Convert the 0-indexed position to a 1-indexed column number, then to a letter
            column_letter = openpyxl.utils.get_column_letter(col_index + 1)
        except ValueError:
            print(f"Error: Column '{column_name}' not found in the header row.")
            return hyperlinks_info

        print(f"Column '{column_name}' corresponds to column letter '{column_letter}'.")

        # 2. Iterate through cells in the identified column (starting from the second row to skip header)
        for cell in sheet[column_letter][1:]:  # [1:] slices the tuple to start from the second element
            if cell.hyperlink:
                # Append a tuple: (display_text, hyperlink_target)
                hyperlinks_info.append((cell.value, cell.hyperlink.target))

    except FileNotFoundError:
        print(f"Error: File not found at {file_path}")
    except Exception as e:
        print(f"An error occurred: {e}")

    # Remove duplicates by converting to a set and then back to a list
    unique_hyperlinks = list(set(hyperlinks_info))

    # Sort the list of tuples alphabetically by the hyperlink's text (the first element of the tuple)
    unique_hyperlinks.sort(key=lambda x: x[0])

    return unique_hyperlinks

class WiData(NamedTuple):
    wi_name:str
    wi_url:str

def parse_tdoc_3gu_list_for_wis(file_path) -> Dict[str,str]:
    hyperlinks_list = parse_excel_hyperlinks_by_column_name(file_path=file_path, column_name='Related WIs')
    return dict(hyperlinks_list)


