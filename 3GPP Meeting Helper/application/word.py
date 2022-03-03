import os
import traceback

import win32com.client

# Global Word instance does not work (removed)
# word = None

def get_word():
    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except:
        try:
            word = win32com.client.Dispatch("Word.Application")
        except:
            word = None
    if word is not None:
        try:
            word.Visible = True
        except:
            print('Could not set property "Visible" from Word to "True"')
        try:
            word.DisplayAlerts = False
        except:
            print('Could not set property "DisplayAlerts" from Word to "False"')
    return word


def open_word_document(filename='', set_as_active_document=True):
    if (filename is None) or (filename == ''):
        doc = get_word().Documents.Add()
    else:
        doc = get_word().Documents.Open(filename)
    if set_as_active_document:
        get_word().Activate()
        doc.Activate()
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


