import os
import os.path
import re
import parsing.word as word_parser
import traceback
import win32com.client
import collections
from config.tdoc_regex_matching import tdoc_regex

spec_file_regex   = re.compile(r'(?P<series>[\d]{2})(\.)?(?P<number>[\d]{3})(-(?P<version>[\w\d]*))?(\.zip)?')
spec_number_regex = re.compile(r'(?P<series>[\d]{2})\.(?P<number>[\d]{3})')

TS = collections.namedtuple('TS', 'series number version match')

def openfiles(files, return_metadata=False, go_to_page=1):
    if files is None:
        return 0
    opened_files = 0
    metadata_list = []
    for file in files:
        file_opened, metadata = openfile(file, return_metadata=True, go_to_page=go_to_page)
        if file_opened:
            opened_files += 1
            metadata_list.append(metadata)
    if return_metadata:
        return opened_files, metadata_list
    else:
        return opened_files

def openfile(file, return_metadata=False, go_to_page=1):
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
        if ((extension=='doc') or (extension=='docx')) and not_mac_metadata:
            doc = word_parser.open_word_document(file)
            metadata = word_parser.get_metadata_from_doc(doc)
            if go_to_page!=1:
                pass
                # Not working :(
                #constants = win32com.client.constants
                #if go_to_page<0:
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
        traceback.print_exc()
        return_value = False
    if return_metadata:
        return return_value, metadata
    else:
        return return_value

def write_data_and_open_file(data, local_file, open_file=True):
    if data is None:
        return
    with open(local_file,'wb') as output:
        output.write(data)
    if open_file:
        openfile(local_file)

def is_tdoc(tdoc):
    if (tdoc is None) or (tdoc == ''):
        return False
    tdoc = tdoc.strip()
    regex_match = tdoc_regex.match(tdoc)
    if regex_match is None:
        return False
    return regex_match.group(0) == tdoc

def is_ts(tdoc):
    if (tdoc is None) or (tdoc == ''):
        return False
    tdoc = tdoc.strip()
    regex_match = spec_number_regex.match(tdoc)
    if regex_match is None:
        return False
    return regex_match.group(0) == tdoc

def parse_ts_number(ts):
    if ts is None:
        return None
    regex_match = spec_file_regex.match(ts)
    if regex_match is None:
        return None
    grouptdict = regex_match.groupdict()
    full_match = regex_match.group(0)
    if ts != full_match:
        return None
    try:
        series = int(grouptdict['series'])
        number = int(grouptdict['number'])
        version = grouptdict['version']
        if version is None:
            version = ''
    except:
        return None
    return TS(series, number, version, full_match)

def get_tdoc_year(tdoc, include_revision=False):
    if not is_tdoc(tdoc):
        if not include_revision:
            return None, None
        return None, None, None
    regex_match = tdoc_regex.match(tdoc)
    if regex_match is None:
        return None
    match_groups = regex_match.groupdict()
    year = int(match_groups['year']) + 2000
    tdoc_number = int(match_groups['tdoc_number'])

    if not include_revision:
        return year,tdoc_number

    try:
        revision = match_groups['revision']
        return year,tdoc_number,revision
    except:
        return year,tdoc_number,None
