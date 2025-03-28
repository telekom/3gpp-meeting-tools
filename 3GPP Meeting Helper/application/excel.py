import platform
import traceback
from typing import List
import pandas as pd
import pyperclip

from application.os import startfile

if platform.system() == 'Windows':
    print('Windows System detected. Importing win32.client')
    import win32com.client

# Global Excel instance does not work (removed)
# excel = None
from application import sensitivity_label


def get_excel():
    if platform.system() != 'Windows':
        return None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        try:
            excel.Visible = True
            excel.DisplayAlerts = False
        except Exception as e1:
            print(f'Could not set Excel instance Visible and/or DisplayAlerts property: {e1}')
        return excel
    except Exception as e2:
        print(f'{e2}')
        traceback.print_exc()
        return None


def open_excel_document(filename=None, sheet_name=None):
    """

    Args:
        filename: File to open
        sheet_name: Sheet name in the Workbook

    Returns: A Workbook object. See https://learn.microsoft.com/en-us/office/vba/api/excel.workbook

    """
    if platform.system() != 'Windows':
        if filename is not None and filename != '':
            startfile(filename)
        return None
    if (filename is None) or (filename == ''):
        wb = get_excel().Workbooks.Add()
    else:
        wb = get_excel().Workbooks.Open(filename)

        # Set sensitivity level (if applicable)
        wb = sensitivity_label.set_sensitivity_label(wb)
    if sheet_name is not None:
        select_worksheet(wb, sheet_name)
    return wb


def select_worksheet(wb, name):
    if wb is None:
        return
    wb.Worksheets(name).Activate()


def set_first_row_as_filter(wb, ws_name=None, already_activated=False):
    if wb is None:
        return
    try:
        if not already_activated:
            wb.Activate()
        if ws_name is None:
            ws = wb.ActiveSheet
        else:
            ws = wb.Sheets(ws_name)
            ws.Activate()

        # See https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofilter
        # If you omit all the arguments, this method simply toggles the display of the
        #  AutoFilter drop-down arrows in the specified range.
        ws.Range("1:1").AutoFilter()
        ws.Cells(2, 2).Select()
        get_excel().ActiveWindow.FreezePanes = True
    except Exception as e:
        print(f'Could not set first row as filter: {e}')
        traceback.print_exc()


def set_autofilter_values(
        wb,
        value_list: List[str],
        ws_name=None,
        already_activated=False,
        column_one_indexed=1
):
    if wb is None:
        return
    try:
        if not already_activated:
            wb.Activate()
        if ws_name is None:
            ws = wb.ActiveSheet
        else:
            ws = wb.Sheets(ws_name)
            ws.Activate()

        # https://learn.microsoft.com/en-us/office/vba/api/excel.xlautofilteroperator
        # TDoc
        # xlFilterValues
        ws.Range("1:1").AutoFilter(
            Criteria1=value_list,
            Field=column_one_indexed,
            Operator=7
        )
    except Exception as e:
        print(f'Could not set autofilter: {e}')
        traceback.print_exc()


def close_wb(wb):
    if wb is None:
        return
    wb.Close()


def vertically_center_all_text(wb):
    if wb is None:
        return
    try:
        wb.Activate()
        ws = wb.ActiveSheet
        # Constants do not work well with win32com, so we just use the value directly
        # https://docs.microsoft.com/en-us/office/vba/api/excel.xlvalign
        ws.Range("A:" + last_column).EntireRow.VerticalAlignment = -4108
    except Exception as e:
        print(f'{e}')
        traceback.print_exc()


def rgb_to_hex(rgb):
    # s.Cells(1, i).Interior.color uses bgr in hex
    bgr = (rgb[2], rgb[1], rgb[0])
    strValue = '%02x%02x%02x' % bgr
    # print(strValue)
    iValue = int(strValue, 16)
    return iValue


def hide_columns(wb, columns):
    if wb is None:
        return
    try:
        wb.Activate()
        ws = wb.ActiveSheet

        for column in columns:
            print('Hiding column {0}'.format(column))
            ws.Columns(column).Hidden = True
    except Exception as e:
        print(f'{e}')
        traceback.print_exc()


def save_wb(wb):
    if wb is None:
        return
    try:
        wb.Activate()
        wb.Save()
        print('Workbook saved!')
    except Exception as e:
        print(f'{e}')
        traceback.print_exc()


def set_column_width(column_letter: str, wb, width: int):
    """
    Sets the width of a column in the active WorkSheet
    Args:
        column_letter: The column's letter
        wb: The WorkBook
        width: The width
    """
    if wb is None:
        return
    column_letter = column_letter.upper()
    wb.Activate()
    ws = wb.ActiveSheet
    ws.Range(column_letter + ":" + column_letter).ColumnWidth = width


def hide_column(column_letter: str, wb):
    """
    Hides a column in the active WorkSheet
    Args:
        column_letter: The column's letter
        wb: The WorkBook
    """
    if wb is None:
        return
    column_letter = column_letter.upper()
    wb.Activate()
    ws = wb.ActiveSheet
    ws.Range(column_letter + ":" + column_letter).EntireColumn.Hidden = True


def set_wrap_text(wb):
    """
    Sets Wrap Text for all cells in the active WorkBook
    Args:
        wb: The WorkBook
    """
    if wb is None:
        return
    wb.Activate()
    ws = wb.ActiveSheet
    all_cells = ws.Cells
    all_cells.WrapText = True


def set_row_height(wb):
    """
    Sets Wrap Text for all cells in the active WorkBook
    Args:
        wb: The WorkBook
    """
    if wb is None:
        return
    wb.Activate()
    ws = wb.ActiveSheet
    all_cells = ws.Cells
    all_cells.EntireRow.AutoFit()


def export_columns_to_markdown(wb, columns: List[str], columns_to_scan='A:AE'):
    """
    Exports specific columns to Markdown and puts the content in the clipboard
    Args:
        wb: The WorkBook
        columns: The columns to export, based on the first row's name
        columns_to_scan: The number of columns to scan
    """
    # XlCellType enumeration
    # https://learn.microsoft.com/en-us/office/vba/api/excel.xlcelltype
    xlCellTypeVisible = 12
    titles = []
    row_list = []

    def to_markdown_text(a_cell):
        if len(a_cell.Hyperlinks) > 0:
            return f'[{a_cell.Value}]({a_cell.Hyperlinks[0].Address})'
        return a_cell.Value

    try:
        ws = wb.ActiveSheet
        visible_cells_range = ws.Range(columns_to_scan).SpecialCells(xlCellTypeVisible)
        for row in visible_cells_range.Rows:
            row_content = tuple(to_markdown_text(cell) for cell in row.Cells)
            if len(row_content) == 0 or row_content[0] is None:
                break
            if not titles:
                titles = row_content
            else:
                row_list.append(row_content)

        df = pd.DataFrame(row_list, columns=titles)
        df_to_output = df.loc[:, columns]
        markdown_table = df_to_output.to_markdown(index=False)
        pyperclip.copy(markdown_table)
        print(f'Copied table of length {len(markdown_table)} to clipboard')
    except Exception as e:
        print(f'Could not parse Excel rows: {e}')
        traceback.print_exc()


last_column = 'U'
