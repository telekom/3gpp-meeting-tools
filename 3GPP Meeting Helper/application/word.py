import os
import traceback
from enum import Enum
from typing import List

import win32com.client

# Global Word instance does not work (removed)
# word = None

# See https://docs.microsoft.com/en-us/office/vba/api/word.wdexportformat
from application import sensitivity_label

wdExportFormatPDF = 17  # PDF format

# https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveformat
wdFormatHTML = 8
wdFormatFilteredHTML = 10

# See https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdexportcreatebookmarks?view=word-pia
wdExportCreateHeadingBookmarks = 1

# See https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdexportoptimizefor?view=word-pia
wdExportOptimizeForPrint = 0

def get_word(visible=True, display_alerts=False):
    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except:
        try:
            word = win32com.client.Dispatch("Word.Application")
        except:
            word = None
    if word is not None:
        try:
            word.Visible = visible
        except:
            print('Could not set property "Visible" from Word to "True"')
        try:
            word.DisplayAlerts = display_alerts
        except:
            print('Could not set property "DisplayAlerts" from Word to "False"')
    return word


def open_word_document(filename='', set_as_active_document=True, visible=True, ):
    if (filename is None) or (filename == ''):
        doc = get_word(visible=visible).Documents.Add()
    else:
        doc = get_word().Documents.Open(filename)
    if set_as_active_document:
        get_word().Activate()
        doc.Activate()

    # Set sensitivity level (if applicable)
    doc = sensitivity_label.set_sensitivity_label(doc)

    return doc


def open_file(file, go_to_page=1, metadata_function=None):
    if (file is None) or (file == ''):
        return
    metadata = None
    try:
        (head, tail) = os.path.split(file)
        extension = tail.split('.')[-1]
        not_mac_metadata = True
        try:
            filename_start = tail[0:2]
            if filename_start == '._':
                not_mac_metadata = False
        except:
            pass
        if ((extension == 'doc') or (extension == 'docx')) and not_mac_metadata:
            doc = open_word_document(file)
            if metadata_function is not None:
                metadata = metadata_function(doc)
            else:
                metadata = None
            if go_to_page != 1:
                pass
                # Not working :(
                # constants = win32com.client.constants
                # if go_to_page<0:
                #    doc.GoTo(constants.wdGoToPage,constants.wdGoToLast)
                #    print('Moved to last page')
                # doc.GoTo(constants.wdGoToPage,constants.wdGoToRelative, go_to_page)
        else:
            # Basic avoidance of executables, but anyway per se not very safe... :P
            if extension != 'exe' and not_mac_metadata:
                os.startfile(file, 'open')
            else:
                print('Executable file {0} not opened for precaution'.format(file))
        return_value = True
    except:
        print('Caught exception while opening {0}'.format(file))
        traceback.print_exc()
        return_value = False
    if metadata_function is not None:
        return return_value, metadata
    else:
        return return_value


def open_files(files, metadata_function=None, go_to_page=1):
    if files is None:
        return 0
    opened_files = 0
    metadata_list = []
    for file in files:
        print('Opening {0}'.format(file))
        if metadata_function is not None:
            file_opened, metadata = open_file(file, metadata_function=metadata_function, go_to_page=go_to_page)
        else:
            file_opened = open_file(file, metadata_function=metadata_function, go_to_page=go_to_page)
        if file_opened:
            opened_files += 1
            if metadata_function is not None:
                metadata_list.append(metadata)
    if metadata_function is not None:
        return opened_files, metadata_list
    else:
        return opened_files


class ExportType(Enum):
    PDF = 1
    HTML = 2


def export_document(
        word_files: List[str],
        export_format: ExportType = ExportType.PDF,
        exclude_if_includes='_rm.doc') -> List[str]:
    """
    Converts a given set of Word files to PDF/HTML
    Args:
        export_format: The format to which the document should be exported to
        word_files: String list containing local paths to the Word files to convert
        exclude_if_includes: a string suffix to ignore certain files (e.g. files with change marks)
    Returns:
        String list containing local paths to the converted PDF files
    """
    pdf_files = []
    if word_files is None or len(word_files) == 0:
        return pdf_files

    # Filter out some files (e.g. files with change tracking)
    if exclude_if_includes != '' and exclude_if_includes is not None:
        word_files = [ e for e in word_files if exclude_if_includes not in e ]

    print('Converting to PDF: {0}'.format(word_files))

    if export_format == ExportType.HTML:
        extension = '.html'
    else:
        extension = '.pdf'

    try:
        word = None
        for word_file in word_files:
            file, ext = os.path.splitext(word_file)
            if ext == '.doc' or ext == '.docx':
                # See https://stackoverflow.com/questions/6011115/doc-to-pdf-using-python
                out_file = file + extension
                print('Export file path: {0}'.format(out_file))
                if not os.path.exists(out_file):
                    if word is None:
                        word = get_word()
                    print('Converting {0} to {1}'.format(word_file, out_file))
                    doc = word.Documents.Open(word_file)

                    if export_format == ExportType.PDF:
                        # See https://docs.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat
                        doc.ExportAsFixedFormat(
                            OutputFileName=out_file,
                            ExportFormat=wdExportFormatPDF,
                            OpenAfterExport=False,
                            OptimizeFor=wdExportOptimizeForPrint,
                            IncludeDocProps=True,
                            CreateBookmarks=wdExportCreateHeadingBookmarks
                        )
                    else:
                        doc.WebOptions.AllowPNG = True
                        doc.SaveAs2(
                            FileName=out_file,
                            FileFormat=wdFormatFilteredHTML
                        )

                    print('Converted {0} to {1}'.format(word_file, out_file))
                else:
                    print('{0} already exists. No need to convert'.format(out_file))
                pdf_files.append(out_file)
    except:
        print('Could not export Word document')
        traceback.print_exc()
    return pdf_files
