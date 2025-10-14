import os
import re
import shutil
import traceback
import zipfile
from typing import List, Tuple, Any, Callable
from zipfile import ZipFile

import platform

from application.common import ExportType, ActionAfter, DocumentMetadata, get_extension_type
from application.os import startfile
from server.common.Tdoc import Tdoc
from server.common.server_utils import DownloadedTdocDocument
from utils import local_cache
from utils.caching.common import export_subfolder

if platform.system() == 'Windows':
    print('Windows System detected. Importing win32.client')
    import win32com.client

# See https://docs.microsoft.com/en-us/office/vba/api/word.wdexportformat
from application import sensitivity_label
from utils.local_cache import file_exists, get_export_target_path_for_file, create_folder_if_needed

# Global Word instance does not work (removed)
# word = None

wdExportFormatPDF = 17  # PDF format

# https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveformat
wdFormatHTML = 8
wdFormatFilteredHTML = 10

# https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveformat
# Word default document file format. For Word, this is the DOCX format
wdFormatDocumentDefault = 16

# See https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdexportcreatebookmarks?view=word-pia
wdExportCreateHeadingBookmarks = 1

# See https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdexportoptimizefor?view=word-pia
wdExportOptimizeForPrint = 0


def get_word(visible=True, display_alerts=False):
    if platform.system() != 'Windows':
        return None

    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except Exception as e1:
        print(f'Could not GetActiveObject "Word.Application": {e1}')
        try:
            word = win32com.client.Dispatch("Word.Application")
        except Exception as e2:
            print(f'Could not Dispatch "Word.Application": {e2}')
            word = None
    if word is not None:
        try:
            word.Visible = visible
        except Exception as e3:
            print(f'Could not set property "Visible" from Word to "True": {e3}')
        try:
            word.DisplayAlerts = display_alerts
        except Exception as e4:
            print(f'Could not set property "DisplayAlerts" from Word to "False": {e4}')
    return word


def open_word_document(filename='', set_as_active_document=True, visible=True, ) -> Any|None:
    if filename is None or filename == '':
        return None
    if platform.system() != 'Windows':
        print(f'Opening {filename} using COM wrapper only available on Windows')
        startfile(filename)
        return None
    try:
        if (filename is None) or (filename == ''):
            doc = get_word(visible=visible).Documents.Add()
        else:
            doc = get_word().Documents.Open(filename)
    except Exception as e:
        print(f'Could not open Word file {filename}: {e}')
        traceback.print_exc()
        return None

    if set_as_active_document:
        get_word().Activate()
        doc.Activate()

    # Set sensitivity level (if applicable)
    doc = sensitivity_label.set_sensitivity_label(doc)

    return doc


def open_file(
        file,
        go_to_page=1,
        metadata_function: Callable[[Any, Any], Any]|None=None,
        convert_to_before_opening=ExportType.NONE
) -> None | bool | Tuple[bool, Any]:
    if (file is None) or (file == ''):
        return None
    metadata = None
    try:
        (head, tail) = os.path.split(file)
        extension = tail.split('.')[-1]
        not_mac_metadata = True
        try:
            filename_start = tail[0:2]
            if filename_start == '._':
                not_mac_metadata = False
        except Exception as e:
            print(f'Could not split filepath: {e}')
            pass

        if ((extension == 'doc') or (extension == 'docx')) and not_mac_metadata:
            if convert_to_before_opening == ExportType.NONE:
                doc = open_word_document(file)
                if metadata_function is not None and doc is not None:
                    metadata = metadata_function(doc, file)
                else:
                    metadata = None
            elif convert_to_before_opening==ExportType.PDF or convert_to_before_opening==ExportType.HTML:
                [head,_] = os.path.split(file)
                export_folder = os.path.join(head, export_subfolder)
                create_folder_if_needed(export_folder)
                exported_files = export_document(word_files=[file],
                                export_format=convert_to_before_opening,
                                export_folder=export_folder)
                metadata = None
                for file in exported_files:
                    metadata = DocumentMetadata(
                        title=None,
                        source=None,
                        path=file)
                    startfile(metadata.path)
            else:
                print('Conversion not PDF or HTML. Not implemented')
            if go_to_page != 1:
                pass
                # Not working :(
                # constants = win32com.client.constants
                # if go_to_page<0:
                #    doc.GoTo(constants.wdGoToPage,constants.wdGoToLast)
                #    print('Moved to last page')
                # doc.GoTo(constants.wdGoToPage,constants.wdGoToRelative, go_to_page)
        else:
            metadata = DocumentMetadata(
                title=None,
                source=None,
                path=file)

            # Basic avoidance of executables, but anyway per se not very safe... :P
            if extension != 'exe' and not_mac_metadata:
                startfile(file)
            else:
                print('Executable file {0} not opened for precaution'.format(file))
        return_value = True
    except Exception as e:
        print(f'Caught exception while opening {file}: {e}')
        traceback.print_exc()
        return_value = False
    if metadata_function is not None:
        return return_value, metadata
    else:
        return return_value


def open_files(
        files,
        metadata_function: Callable[[Any, str], DocumentMetadata] | None = None,
        go_to_page=1,
        convert_to_before_opening: ExportType = ExportType.NONE) \
        -> int | Tuple[int, List[DocumentMetadata]]:
    if files is None:
        return 0
    opened_files_count = 0
    metadata_list = []
    for file in files:
        print(f'Opening {file}')
        try:
            if metadata_function is not None:
                file_opened, metadata = open_file(
                    file,
                    metadata_function=metadata_function,
                    go_to_page=go_to_page,
                    convert_to_before_opening=convert_to_before_opening
                )
                metadata = DocumentMetadata(
                    title=metadata.title if metadata is not None else None,
                    source=metadata.source if metadata is not None else None,
                    path=file)
            else:
                file_opened = open_file(
                    file,
                    metadata_function=metadata_function,
                    go_to_page=go_to_page,
                    convert_to_before_opening=convert_to_before_opening
                )
                metadata = []
        except Exception as e:
            print(f'Could not open {file}: {e}')
            traceback.print_exc()

            file_opened = None
            metadata = None
        if file_opened:
            opened_files_count += 1
            if metadata_function is not None:
                metadata_list.append(metadata)

    if metadata_function is not None:
        return opened_files_count, metadata_list
    else:
        return opened_files_count


def export_document(
        word_files: List[str],
        export_format: ExportType = ExportType.PDF,
        exclude_if_includes='_rm.doc',
        remove_all_fields=False,
        accept_all_changes=False,
        do_after=ActionAfter.NOTHING,
        export_folder=None
) -> List[str]:
    """
    Converts a given set of Word files to PDF/HTML
    Args:
        do_after: What to do after the conversion
        remove_all_fields: Whether to remove all fields
        accept_all_changes: Whether to accept all changes
        export_format: The format to which the document should be exported to
        word_files: String list containing local paths to the Word files to convert
        exclude_if_includes: a string suffix to ignore certain files (e.g. files with change marks)
        export_folder: If specified, a folder where to place the exported file
    Returns:
        String list containing local paths to the converted PDF files
    """
    if export_format == ExportType.NONE:
        # Nothing to do
        return []

    if platform.system() != 'Windows':
        print('Word export to PDF/HTML only available in Windows')
        return []

    converted_files = []
    if word_files is None or len(word_files) == 0:
        return converted_files

    # Filter out some files (e.g. files with change tracking)
    if exclude_if_includes != '' and exclude_if_includes is not None:
        word_files = [e for e in word_files if exclude_if_includes not in e]

    extension = get_extension_type(export_format)
    print(f'Converting to {export_format.name} ({extension})')

    if export_format != ExportType.NONE:
        word = None
        for word_file in word_files:
            try:
                if export_folder is None:
                    file, ext = os.path.splitext(word_file)
                else:
                    # Split into directory and file
                    folder, filename = os.path.split(word_file)
                    # Split filename into name and extension
                    file, ext = os.path.splitext(filename)
                    file = os.path.join(export_folder, file)

                out_file = file + extension

                if ext == '.doc' or ext == '.docx':
                    # See https://stackoverflow.com/questions/6011115/doc-to-pdf-using-python
                    print('Export file path: {0}'.format(out_file))
                    if not file_exists(out_file):
                        if word is None:
                            word = get_word()
                        print('Converting {0} to {1}'.format(word_file, out_file))
                        doc = open_word_document(word_file, visible=False)

                        if remove_all_fields:
                            print('Removing all Fields')
                            doc.Fields.Unlink()

                        if accept_all_changes:
                            print('Accepting all changes')
                            doc.Revisions.AcceptAll()

                        match export_format:
                            case ExportType.PDF:
                                # .doc files often give problems when exporting to PDF. First convert to .docx
                                if ext == '.doc':
                                    print('A POPUP MAY APPEAR ASKING YOU TO SET A SENSITIVITY LABEL (.doc file extension)')
                                    print('Unfortunately, VBA cannot automate this step. Please set the label manually')
                                    converted_docx_list = export_document(
                                        [word_file],
                                        ExportType.DOCX)
                                    docx_version = converted_docx_list[0]
                                    doc = word.Documents.Open(docx_version)

                                # See https://docs.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat
                                print(f'PDF Conversion started: OutputFileName={out_file}')
                                doc.ExportAsFixedFormat(
                                    OutputFileName=out_file,
                                    ExportFormat=wdExportFormatPDF,
                                    OpenAfterExport=False,
                                    OptimizeFor=wdExportOptimizeForPrint,
                                    IncludeDocProps=True,
                                    CreateBookmarks=wdExportCreateHeadingBookmarks
                                )
                            case ExportType.DOCX:
                                print('DOCX Conversion started')
                                doc.SaveAs2(
                                    FileName=out_file,
                                    FileFormat=wdFormatDocumentDefault
                                )
                            case _:
                                print('HTML Conversion started')
                                doc.WebOptions.AllowPNG = True
                                doc.SaveAs2(
                                    FileName=out_file,
                                    FileFormat=wdFormatFilteredHTML
                                )

                        print('Converted {0} to {1}'.format(word_file, out_file))

                        if do_after == ActionAfter.CLOSE_FILE or do_after == ActionAfter.CLOSE_AND_DELETE_FILE:
                            try:
                                doc.Close()
                                print(f'Closed {word_file}')
                            except Exception as e:
                                print(f'Could not close file {word_file}: {e}')
                        if do_after == ActionAfter.CLOSE_AND_DELETE_FILE:
                            try:
                                os.remove(word_file)
                                print(f'Deleted {word_file}')
                            except Exception as e:
                                print(f'Could not delete file {word_file}: {e}')
                    else:
                        print('{0} already exists. No need to convert'.format(out_file))
                    converted_files.append(out_file)
            except Exception as e:
                print(f'Could not export Word document {word_file}: {e}')
                traceback.print_exc()
    return converted_files


def close_word(force=True):
    """
    Close all Word documents and application
    Args:
        force: Whether to skip saving files
    """
    # See http://www.vbaexpress.com/kb/getarticle.php?kb_id=488
    app = get_word()
    try:
        print('Closing Word instance')
        app.ScreenUpdating = False
        # Loop through open documents
        for x in range(app.Documents.Count):
            # See https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveoptions
            this_doc = app.Documents(1)
            print('Closing {0}'.format(this_doc.FullName))
            this_doc.Close(SaveChanges=0)
        print('Closing Word instance')
        app.Quit(SaveChanges=0)
    except:
        print('Could not close Word documents')


def get_active_document():
    active_word_instance = get_word()
    if active_word_instance is None:
        print("Could not retrieve Word instance")
        return

    try:
        active_document = active_word_instance.ActiveDocument
        active_document_name = active_document.Name
        active_document_folder = active_document.Path
        active_document_path = os.path.join(active_document_folder, active_document_name)
        print("Document: {0}. Located in {1}".format(active_document_name, active_document_path))
        return (active_document, active_document_path, active_document_folder)
    except Exception as e:
        print(f"Could not get retrieve active document: {e}")
        return None


def get_reviews_for_active_document(search_author: str = None, replace_author: str = None):
    """
    Gets all of the review objects in the active Word document and optionally replaces the name of the Author
    Args:
        search_author: If present, a regular expression that will be matched to the author name of revisions
        replace_author: If search_author is present, an author name that will replace the matching author names in search_author

    Returns:

    """

    active_document_info = get_active_document()
    if active_document_info is None:
        return

    active_document = active_document_info[0]
    active_document_path = active_document_info[1]
    active_document_folder = active_document_info[2]

    try:
        print("Retrieving revision marks for {0}".format(active_document.Name))
        document_revisions = list(active_document.Revisions)
        print("Found {0} revisions".format(len(document_revisions)))

        if search_author is None or search_author == "" or replace_author is None or replace_author == "":
            print("Nothing to replace in document")
            return

        author_re = re.compile(search_author)
        matching_document_revisions = [r for r in document_revisions if author_re.match(r.Author) is not None]
        matching_authors = set([r.Author for r in matching_document_revisions])
        print("Found {0} matching revisions to author '{1}'. Matching authors: {2}. Will replace with '{3}'".format(
            len(matching_document_revisions),
            search_author,
            matching_authors,
            replace_author))
        # Closing Word document to edit file and then re-open
        print("Closing {0}".format(active_document_path))
        active_document.Close()
        try:
            zip_folder = os.path.join(active_document_folder, 'zip_tmp')
            print('Unzipping {0}'.format(zip_folder))
            with ZipFile(active_document_path, 'r') as wordfile_as_zip:
                wordfile_as_zip.extractall(path=zip_folder)
            document_xml_path = os.path.join(zip_folder, 'word', 'document.xml')

            # Read in the file
            print('Opening {0}'.format(document_xml_path))
            with open(document_xml_path, 'r', encoding="utf8") as file:
                document_xml_contents = file.read()

            # Change authors
            print('Changing matching authors to {0}'.format(replace_author))
            for matching_author in matching_authors:
                document_xml_contents = document_xml_contents.replace(
                    'w:author="{0}"'.format(matching_author),
                    'w:author="{0}"'.format(replace_author))

                # Write the file out again
                print('Saving file to {0}'.format(document_xml_path))
                with open(document_xml_path, 'w', encoding="utf8") as file:
                    file.write(document_xml_contents)
        except:
            print('Could not extract and edit /word/document.xml from {0}'.format(active_document_path))
            traceback.print_exc()
            return

        try:
            print('Removing original file {0}'.format(active_document_path))
            os.remove(active_document_path)

            # Write the file to the ZIP file again
            # Based on https://stackoverflow.com/questions/58955341/create-zip-from-directory-using-python
            print('Writing file to {0}'.format(active_document_path))
            with ZipFile(active_document_path, 'w', compression=zipfile.ZIP_DEFLATED) as zip_ref:
                for folder_name, subfolders, filenames in os.walk(zip_folder):
                    for filename in filenames:
                        file_path = os.path.join(folder_name, filename)
                        zip_ref.write(file_path, arcname=os.path.relpath(file_path, zip_folder))
            zip_ref.close()

            print('Removing temporarily-extracted files in {0}'.format(active_document_folder))
            shutil.rmtree(os.path.join(zip_folder))
        except:
            print('Could not recreate ZIP file at {0}'.format(active_document_path))
            traceback.print_exc()
            return

        open_word_document(active_document_path)
    except:
        print("Could not get reviews from active document")
        traceback.print_exc()
        return

def convert_tdoc_files_to_format(
        tdoc_documents:list[DownloadedTdocDocument],
        tdoc:Tdoc,
        export_format: ExportType = ExportType.PDF):
    export_folder = tdoc.get_local_export_path
    local_cache.create_folder_if_needed(export_folder)
    total_exported_files = []
    print(f'Exporting {tdoc.tdoc.tdoc} files to {export_format.name}')
    for tdoc_document in tdoc_documents:
        total_exported_files.extend(
            export_document(
                [tdoc_document.path],
                export_format=export_format,
                do_after=ActionAfter.CLOSE_FILE,
                export_folder=export_folder
                ))

    return total_exported_files

