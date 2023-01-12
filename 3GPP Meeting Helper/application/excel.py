import traceback

import win32com.client

# Global Excel instance does not work (removed)
# excel = None
from application import sensitivity_label


def get_excel():
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        try:
            excel.Visible = True
            excel.DisplayAlerts = False
        except:
            print('Could not set Excel instance Visible and/or DisplayAlerts property')
        return excel
    except:
        traceback.print_exc()
        return None


def open_excel_document(filename=None, sheet_name=None):
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
    wb.Worksheets(name).Activate()


def set_first_row_as_filter(wb, ws_name=None, already_activated=False):
    try:
        if not already_activated:
            wb.Activate()
        if ws_name is None:
            ws = wb.ActiveSheet
        else:
            ws = wb.Sheets(ws_name)
            ws.Activate()
        ws.Range("1:1").AutoFilter()
        ws.Cells(2, 2).Select()
        get_excel().ActiveWindow.FreezePanes = True
    except:
        traceback.print_exc()


def close_wb(wb):
    wb.Close()


def vertically_center_all_text(wb):
    try:
        wb.Activate()
        ws = wb.ActiveSheet
        # Constants do not work well with win32com, so we just use the value directly
        # https://docs.microsoft.com/en-us/office/vba/api/excel.xlvalign
        ws.Range("A:" + last_column).EntireRow.VerticalAlignment = -4108
    except:
        traceback.print_exc()


def rgb_to_hex(rgb):
    '''
    ws.Cells(1, i).Interior.color uses bgr in hex

    '''
    bgr = (rgb[2], rgb[1], rgb[0])
    strValue = '%02x%02x%02x' % bgr
    # print(strValue)
    iValue = int(strValue, 16)
    return iValue


def hide_columns(wb, columns):
    try:
        wb.Activate()
        ws = wb.ActiveSheet

        for column in columns:
            print('Hiding column {0}'.format(column))
            ws.Columns(column).Hidden = True
    except:
        traceback.print_exc()


def save_wb(wb):
    try:
        wb.Activate()
        wb.Save()
        print('Workbook saved!')
    except:
        traceback.print_exc()


def set_column_width(column_letter: str, wb, width: int):
    """
    Sets the width of a column in the active WorkSheet
    Args:
        column_letter: The column's letter
        wb: The WorkBook
        width: The width
    """
    column_letter = column_letter.upper()
    wb.Activate()
    ws = wb.ActiveSheet
    ws.Range(column_letter+":"+column_letter).ColumnWidth = width


def hide_column(column_letter: str, wb):
    """
    Hides a column in the active WorkSheet
    Args:
        column_letter: The column's letter
        wb: The WorkBook
    """
    column_letter = column_letter.upper()
    wb.Activate()
    ws = wb.ActiveSheet
    ws.Range(column_letter+":"+column_letter).EntireColumn.Hidden = True


def set_wrap_text(wb):
    """
    Sets Wrap Text for all cells in the active WorkBook
    Args:
        wb: The WorkBook
    """
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
    wb.Activate()
    ws = wb.ActiveSheet
    all_cells = ws.Cells
    all_cells.EntireRow.AutoFit()


last_column = 'U'