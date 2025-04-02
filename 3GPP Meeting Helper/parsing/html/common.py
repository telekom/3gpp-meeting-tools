import collections
import datetime
import os
import os.path
import re
import traceback
from re import Pattern
from typing import NamedTuple, List, Tuple

from lxml import html as lh
from lxml.etree import tostring

import utils.local_cache


class Meeting(NamedTuple):
    """
    Represents a 3GPP Meeting (text, folder, date) as parsed from the page.
    e.g., https://www.3gpp.org/ftp/tsg_sa/WG2_Arch
    """
    # e.g.: 163, Jeju 2024-05
    text: str
    # e.g.: TSGS2_163_Jeju_2024-05
    folder: str
    # As parsed from the server page. e.g.: 2024/05/26 22:45
    date: datetime.datetime


class TdocBasicInfo(NamedTuple):
    """
    The basic information related to a TDoc (the TDoc ID, title, source, AI, WI)
    """
    tdoc: str
    title: str
    source: str
    ai: str
    work_item: str


class FolderList(NamedTuple):
    """Provides a list of folders containing datetime information"""
    location: str
    folders: List[str]
    files: List[str]
    folders_with_dates: List[Tuple[str, datetime.datetime]]


ftp_list_regex = re.compile(r'(\d?\d\/\d?\d\/\d\d\d\d) *(\d?\d:\d\d) (AM|PM) *(<dir>|\d+) *')
comment_span = re.compile(r'<span title="(.*)">(.*)')
current_cache_version = 1.41

# Control maximum recursion to avoid stack overflow. Some manual errors in the TDocsByAgenda may lead to circular
# references
max_recursion = 10


def get_tdoc_info(tdoc, df):
    if (tdoc is None) or (df is None) or (tdoc == '') or (tdoc not in df.index):
        return None
    return TdocBasicInfo(tdoc, df.at[tdoc, 'Title'], df.at[tdoc, 'Source'], df.at[tdoc, 'AI'], df.at[tdoc, 'Work Item'])


def get_tdoc_infos(tdocs, df):
    if (tdocs is None) or (tdocs == ''):
        return []
    tdocs.replace(' ', '')
    tdoc_list = [get_tdoc_info(tdoc, df) for tdoc in tdocs.split(',') if tdoc != '']
    tdoc_list = [e for e in tdoc_list if e is not None]
    return tdoc_list


def parse_3gpp_http_ftp(html) -> FolderList:
    if html is None or html == '':
        return None

    # If it is a byte array, then decode
    try:
        html = html.decode("utf-8")
    except (UnicodeDecodeError, AttributeError):
        pass

    if '<title>Directory Listing' in html:
        return parse_3gpp_http_ftp_v2(html)
    else:
        # Old version of the HTML FTP file listing
        return parse_3gpp_http_ftp_v1(html)


def parse_3gpp_http_ftp_v1(html):
    parsed = lh.fromstring(html)
    Result = collections.namedtuple('Http_3GPP_folder', ['location', 'folders', 'files', 'folders_with_dates'])
    location = parsed.xpath('head/title')[0].text_content().replace('www.3gpp.org - ', '').strip()
    text = [e.text_content() for e in parsed.xpath('body/pre/a')][1:]
    full_text = parsed.xpath('body/pre')[0].text_content()
    for value in text:
        full_text = full_text.replace(value, '')
    matches = ftp_list_regex.findall(full_text)
    folders = []
    files = []
    folders_with_dates = []
    for match_idx in range(0, len(text)):
        match = matches[match_idx]
        current_text: str = text[match_idx]
        if match[-1] == '<dir>':
            folders.append(current_text)
            the_time: datetime = datetime.datetime.strptime(match[0], '%m/%d/%Y')
            folders_with_dates.append((current_text, the_time))
        else:
            files.append(current_text)

    return FolderList(location, folders, files, folders_with_dates)


def parse_3gpp_http_ftp_v2(html):
    parsed = lh.fromstring(html)

    location = parsed.xpath('head/title')[0].text_content().replace('Directory Listing', '').strip()

    # Examples:
    #   - ['TSGS2_07', '2008/11/04', '20:21']
    #   - ['Ad-hoc_meetings', '2011/05/05', '5:35']
    row_list = [(
        list(filter(
            None,
            e.text_content().replace('\r', '').replace('\n', '').replace('\t', '').strip().split(' '))),
        '/ftp/geticon.axd?file=' in tostring(e, encoding=str)
    )
        for e in parsed.xpath('body/form/table/tbody/tr')]

    folders = []
    files = []
    folders_with_dates = []

    parsed_folder_count = 0
    for row in row_list:
        row_content = row[0]
        row_is_dir = row[1]

        entry_name = row_content[0]
        entry_date = row_content[1]
        entry_time = row_content[2]

        entry_datetime = f'{entry_date} {entry_time}'
        if row_is_dir:
            folders.append(entry_name)
            try:
                folders_with_dates.append((entry_name, datetime.datetime.strptime(entry_datetime, '%Y/%m/%d %H:%M')))
                parsed_folder_count = parsed_folder_count + 1
            except ValueError as e:
                print(f'Could not parse time of row {row}: {e}')
        else:
            files.append(entry_name)

    print(f'Parsed {parsed_folder_count} folders from HTML file')
    return FolderList(location, folders, files, folders_with_dates)


def parse_current_document(html):
    """
    Parses the "current document" HTML file. Do note that this file is not always used (used mainly by Maurice)
    and not in breakout sessions. So it is not always reliable.
    Args:
        html:

    Returns:

    """
    if html is None or html == '':
        return None
    parsed = lh.fromstring(html)
    tdoc_id = parsed.xpath('body/table/tr/td')[2].text_content()
    return tdoc_id


def parse_3gpp_meeting_list(html, ordered=True, remove_old_meetings=False):
    if html is None or html == '':
        return None
    meeting_regex = re.compile(r'TSGS2_.*')
    parsed_html = parse_3gpp_http_ftp(html)
    folder_list = parsed_html.folders_with_dates
    meeting_folders = []
    for folder_data in folder_list:
        folder = folder_data[0]
        folder_date = folder_data[1]
        if not meeting_regex.match(folder):
            continue
        folder_without_prefix = folder.replace('TSGS2_', '')
        meeting_location = folder_without_prefix.split('_')
        if len(meeting_location) > 1:
            meeting_location_joined = ' '.join(meeting_location[1:])
            meeting_folders.append(Meeting(meeting_location[0] + ', ' + meeting_location_joined, folder, folder_date))
        else:
            meeting_folders.append(Meeting(folder_without_prefix, folder, folder_date))

    number_parse_regex = re.compile(r'\d*')
    if remove_old_meetings:
        meeting_folders = [folder for folder in meeting_folders if
                           get_meeting_number(folder.text, number_parse_regex) > 45]
    if ordered:
        meeting_folders = sorted(meeting_folders, reverse=True,
                                 key=lambda folder: get_meeting_number(folder.text, number_parse_regex))
    return meeting_folders


def parse_3gpp_meeting_list_object(html, ordered=True, remove_old_meetings=False):
    meeting_data = parse_3gpp_meeting_list(html, ordered, remove_old_meetings)
    if meeting_data is None:
        return None
    return MeetingData(meeting_data)


def get_meeting_number(number_string: str, regex: Pattern[str]):
    """
    Returns the meeting number as an integer based on the meeting string
    Args:
        number_string: A string containing a meeting number
        regex: Regular expression to use matching the number that will be used with regex.match

    Returns: The number

    """
    key_match = regex.match(number_string)
    if key_match is None:
        return 0
    return int(key_match.group())


def sort_and_remove_duplicates_from_list(a_list):
    if (a_list is None) or (type(a_list) != list):
        return a_list
    return sorted(set([tdoc for tdoc in a_list if tdoc != '']))

    # Join results


def join_results(tdocs_split, df_tdocs, recursive_call, original_index, n_recursion):
    all_results = []
    try:
        for tdoc in tdocs_split:
            if (tdoc is None) or (tdoc == ''):
                continue

            results = recursive_call(tdoc, df_tdocs, original_index, n_recursion + 1)
            if type(results) != list:
                if results is not None:
                    results = results.split(',')
                else:
                    results = []
            all_results.extend(results)

        if len(all_results) == 0:
            return ''

        # Need the sorted() function to make the tests reproducible
        return ','.join(sort_and_remove_duplicates_from_list(all_results))
    except:
        print('Could not join docs')
        traceback.print_exc()
        return ''


def get_cache_filepath(meeting_folder_name, html_hash):
    if meeting_folder_name == '':
        return None
    meeting_local_folder = utils.local_cache.get_local_agenda_folder(meeting_folder_name)
    file_name = 'TDocsByAgenda_{0}_{1}.pickle'.format(meeting_folder_name, html_hash)
    full_path = os.path.join(meeting_local_folder, file_name)
    return full_path


class MeetingData:
    """Allows easy access to the overall meeting data from a list of meetings and provides convenient mapping
    functions"""

    def __init__(self, meeting_data: List[Meeting]):
        self._meeting_data = meeting_data
        self._meeting_names = [t.text for t in self._meeting_data]
        self._meeting_mapping = {t.text: t.folder for t in self._meeting_data}
        self._meeting_number_to_menu_option_mapping = {t.text.split(',')[0].strip(): t.text for t in self._meeting_data}
        self._meeting_text_to_year_mapping = {t.text: t.date.year for t in self._meeting_data}
        self._meeting_number_to_year_mapping = {t.text.split(',')[0].strip(): t.date.year for t in self._meeting_data}

    def get_sa2_meeting_data(self):
        return self._meeting_data

    @property
    def meeting_names(self):
        return self._meeting_names

    def get_server_folder_for_meeting_choice(self, text):
        try:
            return self._meeting_mapping[text]
        except:
            return None

    def get_meeting_text_for_given_meeting_number(self, meeting_number):
        try:
            return self._meeting_number_to_menu_option_mapping[str(meeting_number).strip()]
        except:
            return None

    @property
    def length(self):
        return len(self._meeting_data)

    def get_year_from_meeting_text(self, text: str):
        try:
            return self._meeting_text_to_year_mapping[text]
        except:
            return None

    def get_year_from_meeting_number(self, number):
        try:
            return self._meeting_number_to_year_mapping[str(number)]
        except:
            return None

    def get_meetings_for_given_year(self, year):
        filtered_meeting_data = [meeting for meeting in self._meeting_data if meeting.date.year == year]
        return MeetingData(filtered_meeting_data)
