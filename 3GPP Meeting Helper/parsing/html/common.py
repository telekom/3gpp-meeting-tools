import collections
import datetime
import hashlib
import os
import os.path
import pickle
import re
import traceback
from re import Pattern
from typing import NamedTuple, List, Tuple

import pandas as pd
from lxml import html as lh

import config.contributor_names
import utils.local_cache
from parsing.html.common_tools import parse_tdoc_comments, tdoc_regex_str
from parsing.html.tdocs_by_agenda_v3 import assert_if_tdocs_by_agenda_post_sa2_159, parse_tdocs_by_agenda_v3
from server.common import decode_string
from tdoc.utils import title_cr_regex


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
    folders_with_dates: Tuple[str, datetime.datetime]


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
    rows = [(e.xpath('td/a')[0].attrib['href'],  # URL
             e.xpath('td')[2].text_content().replace('\n', '').replace('\t', '').strip(),  # Date
             e.xpath('td/img')[0].attrib['src']  # Icon
             ) for e in parsed.xpath('body/form/table/tbody/tr')]

    folders = []
    files = []
    folders_with_dates = []
    for row in rows:
        split_url = row[0].split('/')
        row_is_dir = (row[2] == '/ftp/geticon.axd?file=')
        entry_name = split_url[-1]
        entry_time = row[1]
        if row_is_dir:
            folders.append(entry_name)
            folders_with_dates.append((entry_name, datetime.datetime.strptime(entry_time, '%Y/%m/%d %H:%M')))
        else:
            files.append(entry_name)

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


# Storing all of them is easier, and the cache should not grow that big in the end
tdocs_by_document_cache = {}


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


class tdocs_by_agenda(object):
    """Contains the information for the "Tdocs by Agenda" meeting file"""
    revision_of_regex = re.compile(r'.*Revision of (?P<tdoc>[S\d-]*)( from (?P<previous_meeting>S[\d\w#-]*))?')
    revised_to_regex = re.compile(
        r'.*Revised([ ]?(off-line|in parallel session|in drafting session))? to (?P<tdoc>[S\d-]*)')
    merge_of_regex = re.compile(r'.*merging (?P<tdoc>(( and )?(, )?(part of)?( )?[S\d-]*)+)')
    merged_to_regex = re.compile(r'.*Merged (into|with) (?P<tdoc>[S\d-]*)')

    # Strings to match:
    #  - Draft Agenda for SA WG2#128
    #  - Draft Agenda for SA WG2#129#BIS
    #  - SA WG2 #137 Meeting Agenda
    #  - SA2#155 Meeting Agenda
    meeting_number_regex = re.compile(r'((SA WG2[ ]?)|(SA2))#(?P<meeting>[\d]{1,3}[#]?[\w]*)')

    creation_date_regex = re.compile(
        r'>Created: <(B|b)>((?P<year>[\d]{4})-(?P<month>[\d]{2})-(?P<day>[\d]{2}) (?P<hour>[\d]{2}):(?P<minute>[\d]{2}))</(B|b)>&nbsp[;]?&nbsp[;]?&nbsp[;]?&nbsp[;]?')
    creation_date_regex_if_fails = re.compile(
        r'<o:LastSaved>((?P<year>[\d]{4})-(?P<month>[\d]{2})-(?P<day>[\d]{2})T(?P<hour>[\d]{2}):(?P<minute>[\d]{2}))')

    # Hardcoded list of typos to correct
    tdoc_typos = {
        'S2-181812649': 'S2-1812649',
        'S2-19000963': 'S2-1900963'
    }

    def get_tdoc_by_agenda_html(path_or_html, return_raw_html=False):
        try:
            # Initial check added, as it seemed to sometimes break things if the cache directory does not exist in fresh installations
            if (len(path_or_html)) < 1000 and os.path.isfile(path_or_html):
                # If the HTML is input into the open function, it will throw an exception
                print('Reading TDocsByAgenda from {0}'.format(path_or_html))
                with open(path_or_html, 'rb') as f:  # with can auto close the file like f.close() does
                    html = f.read()
            else:
                html = path_or_html
        except:
            html = path_or_html

        if html is None:
            print('No HTML to parse. Maybe a communication error occured?')
        print('TDocsByAgenda HTML: {0} bytes'.format(len(html)))

        if return_raw_html:
            return decode_string(html, 'TDocsByAgenda')
        else:
            try:
                parsed_html = lh.fromstring(html)
            except:
                print('Could not parse TDocs by Agenda HTML')
                parsed_html = None
            return parsed_html

    def get_tdoc_by_agenda_date(path_or_html: str) -> datetime.datetime:
        html = tdocs_by_agenda.get_tdoc_by_agenda_html(path_or_html, return_raw_html=True)
        email_approval_results = False

        if html is None:
            print('Cannot not read TDocs by Agenda file')
            return None

        try:
            search_result = tdocs_by_agenda.creation_date_regex.search(html)
            if search_result is None:
                print('Could not parse date from TDocs by Agenda file, trying with LastSaved')
                search_result = tdocs_by_agenda.creation_date_regex_if_fails.search(html)
                email_approval_results = (html.find('This CR was agreed') != -1)
                if email_approval_results:
                    print(
                        'TDocsByAgenda contains e-mail approval results. Assuming last version and returning now() as time')
                    return datetime.datetime.now()
                if search_result is None:
                    print('Could not parse date from TDocs by Agenda file')
                    return None
            year = int(search_result.group('year'))
            month = int(search_result.group('month'))
            day = int(search_result.group('day'))
            hour = int(search_result.group('hour'))
            minute = int(search_result.group('minute'))
            parsed_date = datetime.datetime(year=year, month=month, day=day, hour=hour, minute=minute)
            return parsed_date
        except:
            print('Error parsing date of TDocs by Agenda file')
            traceback.print_exc()
            return None

    def __init__(self, path_or_html, v=2, html_hash='', meeting_server_folder=''):
        self.tdocs = None

        print('Parsing TDocsByAgenda file: version {0}'.format(v))
        raw_html = tdocs_by_agenda.get_tdoc_by_agenda_html(path_or_html, return_raw_html=True)

        try:
            self.meeting_number = tdocs_by_agenda.get_meeting_number(raw_html)
        except:
            self.meeting_number = 'Unknown'
        print('Parsed meeting number: {0}'.format(self.meeting_number))
        self.meeting_server_folder: str = meeting_server_folder

        dataframe_from_cache = False

        if v == 1:
            # print('XPath fro title: ' + html.xpath('//P/FONT/B').tostring())
            html = tdocs_by_agenda.get_tdoc_by_agenda_html(path_or_html)
            dataframe = tdocs_by_agenda.read_tdocs_by_agenda(html)
        else:
            if meeting_server_folder != '' and html_hash != '':
                cache_file_name = get_cache_filepath(meeting_server_folder, html_hash)
                try:
                    if cache_file_name is not None and os.path.exists(cache_file_name):
                        with open(cache_file_name, 'rb') as f:
                            # Unpickle the 'data' dictionary using the highest protocol available.
                            cache = pickle.load(f)
                            if cache['cache_version'] == current_cache_version:
                                dataframe = cache['tdocs']
                                dataframe_from_cache = True
                                print('Loaded TDocsByAgenda from file cache: {0}'.format(cache_file_name))
                                remove_old_cache = False
                            else:
                                print('Cache version mismatch. Not reading from cache. Removing old cache')
                                remove_old_cache = True

                        if remove_old_cache:
                            print('New cache version. Removing old cache (will get re-saved with new version)')
                            os.remove(cache_file_name)

                except:
                    print(
                        'Could not load file cache for meeting {0}, hash {1}'.format(meeting_server_folder, html_hash))
                    traceback.print_exc()

            if not dataframe_from_cache:
                dataframe = tdocs_by_agenda.read_tdocs_by_agenda_v2(raw_html, force_html=True)

        # Cleanup Unicode characters (see https://stackoverflow.com/questions/42306755/how-to-remove-illegal-characters-so-a-dataframe-can-write-to-excel)
        if not dataframe_from_cache:
            print('Cleaning up Unicode characters so that Excel export does not crash')
            dataframe = dataframe.map(
                lambda x: x.encode('unicode_escape').decode('utf-8') if isinstance(x, str) else x)

        # Cleanup comments. Sometimes we have "span" tags polluting comments
        if not dataframe_from_cache:
            print('Cleaning up comments column')
            try:
                dataframe['Comments'] = dataframe['Comments'].apply(lambda x: tdocs_by_agenda.clean_up_comment(x))
            except:
                print('Could not clean-up comments')
                traceback.print_exc()

        # Other cleanups that happened over the time
        dataframe['Title'] = dataframe['Title'].apply(lambda x: tdocs_by_agenda.clean_up_title(x))

        # Assign dataframe
        self.tdocs: pd.DataFrame = dataframe

        if not dataframe_from_cache:
            tdocs_by_agenda.get_original_and_final_tdocs(self.tdocs)
            self.others_cosigners, self.tdocs = config.contributor_names.add_contributor_columns_to_tdoc_list(
                self.tdocs, self.meeting_server_folder)
            self.contributor_columns = config.contributor_names.get_contributor_columns()
        else:
            # get_original_and_final_tdocs should already be in the cache
            self.others_cosigners = cache['others_cosigners']
            self.contributor_columns = cache['contributor_columns']
        config.contributor_names.reset_others()

    def clean_up_comment(comment_str):
        comment_match = comment_span.match(comment_str)
        if comment_match is None:
            return comment_str
        fixed_comment = '{0}. {1}'.format(comment_match.group(1), comment_match.group(2))
        return fixed_comment

    def clean_up_title(title_str: str) -> str:
        if title_str is None:
            return title_str

        return title_str.replace(r'&apos;', "'").replace(r'&amp;', "&").replace(r'&#39;', "'")

    def get_meeting_number(tdocs_by_agenda_html) -> str:
        """
        Returns: The meeting number based on the HTML of TDocsByAgenda
        """
        print('Parsing TDocsByAgenda meeting number')
        meeting_number_match = tdocs_by_agenda.meeting_number_regex.search(tdocs_by_agenda_html)
        if meeting_number_match is None:
            print('Could not parse meeting number from TDocsByAgenda HTML')
            return 'Unknown'
        meeting_number = meeting_number_match.groupdict()['meeting']
        meeting_number = meeting_number.replace('#', '')
        print('Meeting number: {0}'.format(meeting_number))
        return meeting_number.upper()

    def read_tdocs_by_agenda_v2(path_or_html, force_html=False) -> pd.DataFrame:
        # New HTML-parsing method. Regex-based. It works better with malformed and broken HTML files, which appear to happen quite often
        if force_html:
            html = path_or_html
        else:
            html = tdocs_by_agenda.get_tdoc_by_agenda_html(path_or_html, return_raw_html=True)
        print('TDocsByAgenda: HTML file length: {0}'.format(len(html)))

        if assert_if_tdocs_by_agenda_post_sa2_159(html):
            print("TDocsByAgenda is newer than SA2#159")
            df_tdocs = parse_tdocs_by_agenda_v3(html)
            # Post-processing
            df_tdocs = tdocs_by_agenda.post_process_df_tdocs(df_tdocs)
            return df_tdocs
        else:
            print("TDocsByAgenda is prior to SA2#159")

        # Remove beginning
        pre_cleaning_html_length = len(html)
        start_of_tdocs = html.find('List for Meeting:')
        if start_of_tdocs > -1:
            html = html[start_of_tdocs:-1]
            print('Skipped first {0} characters until "List for Meeting:"'.format(start_of_tdocs))

        uncapitalized_separator = '<tr valign='
        capitalized_separator = '<TR VALIGN='
        if html.find(uncapitalized_separator) != -1:
            splitter = uncapitalized_separator
        else:
            splitter = capitalized_separator

        separators = ['(td|TD)', '(bordercolor|BORDERCOLOR)', '(bgcolor|BGCOLOR)', '(font|FONT)',
                      '(font-size|FONT-SIZE)', '(face|FACE)', '(color|COLOR)', '(th|TH)']

        # Strings that will be substituted
        s = {
            'TD': separators[0],
            'TH': separators[7],
            'BORDERCOLOR': separators[1],
            'BGCOLOR': separators[2],
            'FONT': separators[3],
            'FONT-STYLE': '[\\\"]?FONT-SIZE:[\\d]+pt[\\\"]?',
            'FACE': separators[5],
            'COLOR': separators[6],
            'B': 'b',  # Bold tag is always lower-case
            'HEX': '([\\\"]?(#[0-9a-fA-F]+)?[\\\"]?)?',
            '=': '[ ]*=[ ]*',
            'ARIAL': '[\\\"]?Arial[\\\"]?',
            'CRLF': '[\r\n]?',
            'TDOC': tdoc_regex_str,
            'INTEGER': '[\\\"]?[\\d]*[\\\"]?',
        }

        # Generate big regex that will substitute all of the unwanted strings for easier parsing
        strings_to_sub = [
            ' </{B}></{FONT}>'.format(**s),
            '<{FONT} style{=}{FONT-STYLE} {FACE}{=}({ARIAL})?[ \"]{COLOR}{=}{HEX}>{CRLF}'.format(**s),
            '</{FONT}>'.format(**s),
            '<{TD} {BORDERCOLOR}{=}{HEX} {BGCOLOR}{=}{HEX}>'.format(**s),
            '<{TD} [wW][iI][dD][tT][hH]={INTEGER} {BORDERCOLOR}{=}{HEX} {BGCOLOR}{=}{HEX}>'.format(**s),
            '<{TH} (width={INTEGER} )?({TH}[ ]?(={INTEGER})?)?[ ]?{BGCOLOR}{=}{HEX} {BORDERCOLOR}{=}{HEX}[ ]?>'.format(
                **s),
            '<[/]?[bBiI]>',
            '<TD  WIDTH="75" BORDERCOLOR=#000000 BGCOLOR =#C0C0C0>',
            '<[/]?{TH}>'.format(**s)
        ]
        full_sub = '|'.join(strings_to_sub)

        # Cleanup HTML tags
        html = re.sub(full_sub, '', html)

        post_cleaning_html_length = len(html)
        gain = post_cleaning_html_length / pre_cleaning_html_length
        print('TDocsByAgenda: HTML file length after cleaning: {0} ({1:.0%})'.format(len(html), gain))

        split_html = html.split(splitter)
        print('TDocsByAgenda: split in {0} segments. Parsing segments for TDoc info'.format(len(split_html)))

        tdoc_columns = None
        for element in split_html:
            if 'TD#' in element:
                try:
                    element_clean = re.sub('</{TD}>([\r\n])*'.format(**s), '\n', element)
                    tdoc_columns = [e.strip() for e in element_clean.split('\n')[1:-4]]
                    print('Columns found: {0}'.format(tdoc_columns))
                    break
                except:
                    pass

        # Build regex to parse columns
        tdocs_by_agenda_regex = '.*)</{TD}>{CRLF}'.format(**s)
        tdocs_by_agenda_email_disc_regex = '.*({CRLF}=== Close of meeting ===)?)</{TD}>{CRLF}'.format(**s)
        tdocs_by_agenda_tdoc_regex = '[ ]*' + '.*(?P<TD>{TDOC}).*</{TD}>{CRLF}'.format(**s)
        tdoc_by_agenda_simple_tdoc = re.compile(r'(?P<TD>{TDOC})'.format(**s))

        # Take default columns if none are found
        if tdoc_columns is None:
            tdocs_by_agenda_regex_full = '(?P<AI>{0}{1}(?P<TYPE>{0}(?P<Doc For>{0}(?P<Subject>{0}(?P<Source>{0}(?P<Rel>{0}(?P<Work Item>{0}(?P<Comments>{0}(?P<Result>{0}'.format(
                tdocs_by_agenda_regex, tdocs_by_agenda_tdoc_regex)
            column_indices = {
                'AI': 0,
                'TD#': 1,
                'Type': 2,
                'Doc For': 3,
                'Subject': 4,
                'Source': 5,
                'Rel': 6,
                'Work Item': 7,
                'Comments': 8,
                'Result': 9,
                'Subject': None,
                'e-mail_Discussion': None
            }
            total_columns = 10
        # Detect columns in file
        else:
            # Column is sometimes called "Title" sometimes "Subject"
            column_indices = {
                'AI': None,
                'TD#': None,
                'Type': None,
                'Doc For': None,
                'Subject': None,
                'Source': None,
                'Rel': None,
                'Work Item': None,
                'Comments': None,
                'Result': None,
                'Title': None,
                'e-mail_Discussion': None
            }
            for e in ['AI', 'TD#', 'Type', 'Doc For', 'Subject', 'Source', 'Title', 'Rel', 'Work Item', 'Comments',
                      'Result', 'e-mail_Discussion']:
                if e in tdoc_columns:
                    found_index = tdoc_columns.index(e)
                    column_indices[e] = found_index

            # Sort parameter list according to their column index to generate search regex
            column_indices_list = [e for e in column_indices.items()]
            column_indices_list.sort(key=lambda x: x[1] if x[1] is not None else -1)

            # Generate Regex
            tdocs_by_agenda_regex_full = ''
            next_expected_index = None
            total_columns = 0
            for column_name, column_index in column_indices_list:
                if column_index is not None and next_expected_index is None:
                    if 'TD#' in column_name:
                        tdocs_by_agenda_regex_full = '{0}{1}'.format(tdocs_by_agenda_regex_full,
                                                                     tdocs_by_agenda_tdoc_regex)
                        total_columns += 1
                    elif 'e-mail_Discussion' in column_name:
                        tdocs_by_agenda_regex_full = '{0}(?P<{1}>{2}'.format(tdocs_by_agenda_regex_full,
                                                                             column_name.replace(' ', '').replace('-',
                                                                                                                  ''),
                                                                             tdocs_by_agenda_email_disc_regex)
                        total_columns += 1
                    else:
                        tdocs_by_agenda_regex_full = '{0}(?P<{1}>{2}'.format(tdocs_by_agenda_regex_full,
                                                                             column_name.replace(' ', '').replace('-',
                                                                                                                  ''),
                                                                             tdocs_by_agenda_regex)
                        total_columns += 1

        print('TDocsByAgenda: Matching regular expressions')

        TdocMatch = collections.namedtuple('TdocMatch',
                                           'AI TD Type DocFor Title Source Rel WorkItem Comments Result ParsedComments')
        line_separator = re.compile(r'</[tT][dD]>')
        if column_indices['Subject'] is not None:
            title_name = 'Subject'
        else:
            title_name = 'Title'

        def convert_split_to_tdoc_info(info, td_column_idx):
            info_split = line_separator.split(info)
            if len(info_split) != total_columns + 1:
                return None

            tdoc_info = info_split[0:-1]
            found_tdoc = tdoc_by_agenda_simple_tdoc.search(tdoc_info[td_column_idx])
            if found_tdoc is None:
                return None
            tdoc_id = found_tdoc.group('TD')
            tdoc_info[td_column_idx] = tdoc_id
            tdoc_info = [e.strip() for e in tdoc_info]

            try:
                def get_str(column_name):
                    idx = column_indices[column_name]
                    if idx is None:
                        return ''
                    return tdoc_info[idx]

                ai = get_str('AI').replace('"', '').replace('TOP>', '').replace('\n', '').strip()
                comments = get_str('Comments').replace('<span title="">', '').replace('</span>', '').replace(
                    '<span title=" ">', '').strip()
                parsed_tdoc_info = TdocMatch(
                    ai,
                    tdoc_id,
                    get_str('Type'),
                    get_str('Doc For'),
                    get_str(title_name),
                    get_str('Source'),
                    get_str('Rel'),
                    get_str('Work Item'),
                    comments,
                    get_str('Result'),
                    parse_tdoc_comments(comments))
                return parsed_tdoc_info
            except:
                return None

        td_column_idx = column_indices['TD#']
        tdoc_matches = [convert_split_to_tdoc_info(e, td_column_idx) for e in split_html]
        tdoc_matches = [e for e in tdoc_matches if e is not None]

        print('TDocsByAgenda: Organizing TDoc matches')
        print('TDocsByAgenda: found {0} TDocs'.format(len(tdoc_matches)))
        column_for_dictionary = list(zip(*tdoc_matches))

        print('TDocsByAgenda: Generating TDocs DataFrame')
        d = {
            'AI': column_for_dictionary[0],
            'TD#': column_for_dictionary[1],
            'Type': column_for_dictionary[2],
            'Doc For': column_for_dictionary[3],
            'Title': column_for_dictionary[4],
            'Source': column_for_dictionary[5],
            'Rel': column_for_dictionary[6],
            'Work Item': column_for_dictionary[7],
            'Comments': column_for_dictionary[8],
            'Result': column_for_dictionary[9],
            'Revision of': [e.revision_of for e in column_for_dictionary[10]],
            'Revised to': [e.revised_to for e in column_for_dictionary[10]],
            'Merge of': [e.merge_of for e in column_for_dictionary[10]],
            'Merged to': [e.merged_to for e in column_for_dictionary[10]],
            '#': range(len(tdoc_matches))
        }

        print('TDocsByAgenda: Done')

        # Some cleaning for some special characters
        d['Source'] = [at_n_t.replace('&amp;', '&') for at_n_t in d['Source']]

        df_tdocs = pd.DataFrame(data=d)
        df_tdocs = df_tdocs.set_index('TD#')
        df_tdocs.index.name = 'TD#'
        print('TDocsByAgenda: {0} TDocs entries parsed'.format(len(df_tdocs)))
        email_approval_tdocs = df_tdocs[(df_tdocs['Result'] == 'For e-mail approval')]
        n_email_approval = len(email_approval_tdocs)
        print('TDocsByAgenda: {0} TDocs marked as "For e-mail approval"'.format(n_email_approval))

        # Post-processing
        df_tdocs = tdocs_by_agenda.post_process_df_tdocs(df_tdocs)

        return df_tdocs

    def read_tdocs_by_agenda(doc):
        # Legacy method for parsing TdocsByAgenda file. HTML-based parsing assumes a correct HTML file
        word_exported = False
        tdoc_table = doc.xpath('//table')[-1]
        tdoc_header = [element.text_content().strip() for element in tdoc_table.xpath('thead')[0].xpath('tr/th')]
        if len(tdoc_header) == 0:
            tdoc_header = [element.text_content().strip() for element in tdoc_table.xpath('thead')[0].xpath('tr/td')]
            tdoc_header = [e for e in tdoc_header if (e is not None) and (e != '')]
            if len(tdoc_header) > 0:
                word_exported = True
        tdoc_header.extend(['Revision of', 'Revised to', 'Merge of', 'Merged to', '#'])

        if not word_exported:
            tdocs = tdoc_table.xpath('tbody/tr')
        else:
            tdocs = tdoc_table.xpath('tr')

        rows = []
        row_idx = 0
        for tdoc in tdocs:
            tdoc_cols = tdoc.xpath('td')
            cols_as_array = [element.text_content().strip() for element in tdoc_cols]
            if word_exported:
                cols_as_array = cols_as_array[1:]
            new_cols = ['', '', '', '', row_idx]
            cols_as_array.extend(new_cols)

            # Sort out non-TDoc rows (e.g. separator rows)
            if (cols_as_array[1] == '-') or (cols_as_array[1] == ''):
                continue

            # The row index is the last column
            cols_as_array[-1] = row_idx

            added_columns = len(new_cols)

            # Extract chairman comments (auto-generated text from Word macro)
            comments = cols_as_array[-added_columns - 2]
            # print(cols_as_array)
            # print('Comments (row ' + str(-len(new_cols)-2) + '): ' + comments)

            # revision_of_match = tdocs_by_agenda.revision_of_regex.match(comments)
            # revised_to_match = tdocs_by_agenda.revised_to_regex.match(comments)
            # merge_of_match = tdocs_by_agenda.merge_of_regex.match(comments)
            # merged_to_match = tdocs_by_agenda.merged_to_regex.match(comments)

            parsed_comments = parse_tdoc_comments(comments)
            cols_as_array[-(added_columns - 2)] = parsed_comments.merge_of
            cols_as_array[-(added_columns - 3)] = parsed_comments.merged_to
            cols_as_array[-(len(new_cols))] = parsed_comments.revision_of
            cols_as_array[-(added_columns - 1)] = parsed_comments.revised_to

            rows.append(cols_as_array)
            row_idx += 1

        df_tdocs = pd.DataFrame(rows, columns=tdoc_header)
        df_tdocs = df_tdocs.set_index('TD#')
        print('{0} TDocs entries parsed'.format(len(df_tdocs)))

        # Post-processing
        df_tdocs = tdocs_by_agenda.post_process_df_tdocs(df_tdocs)

        return df_tdocs

    def post_process_df_tdocs(df_tdocs):
        # Remove duplicates. It was seen in S2-134 that sometimes duplicate #TD can be present
        # See https://www.dataquest.io/blog/settingwithcopywarning/ as o why the "copy()"
        df_tdocs = df_tdocs.loc[~df_tdocs.index.duplicated(keep='last')].copy()
        print('TDocsByAgenda: {0} TDocs entries parsed after de-duplication'.format(len(df_tdocs)))

        # Fix LS OUTs that have a wrong source
        print('Fixing wrong LS OUT sources')
        ls_with_wrong_source_idx = (df_tdocs['Type'] == 'LS OUT') & (df_tdocs['Revision of'] != '') & (
                df_tdocs['Source'] == 'SA WG2')
        ls_outs_with_wrong_source = df_tdocs.loc[ls_with_wrong_source_idx, :]
        for idx, row in ls_outs_with_wrong_source.iterrows():
            try:
                revision_idx = row['Revision of']
                old_source = df_tdocs.at[revision_idx, 'Source']
                df_tdocs.at[idx, 'Source'] = old_source
                print('{0}: {1}->{2}'.format(idx, row['Source'], df_tdocs.at[idx, 'Source']))
            except:
                pass
        print('Fixed {0} LS OUT sources'.format(len(ls_outs_with_wrong_source)))

        # Extract CR info
        print('TDocsByAgenda: Parsing TS and CR numbers')
        df_tdocs.loc[:, 'TS'] = ''
        df_tdocs.loc[:, 'CR'] = ''

        matches = [(idx, title_cr_regex.match(row['Title'])) for idx, row in df_tdocs.iterrows()]
        matches = [match for match in matches if match[1] is not None]
        for match in matches:
            df_tdocs.at[match[0], 'TS'] = match[1].group(1)
            df_tdocs.at[match[0], 'CR'] = match[1].group(2)
        print('TDocsByAgenda: Parsing TS and CR numbers Done')

        email_approval_tdocs = df_tdocs[(df_tdocs['Result'] == 'For e-mail approval')]
        n_email_approval = len(email_approval_tdocs)
        print('TDocsByAgenda: {0} TDocs marked as "For e-mail approval" after de-duplication'.format(n_email_approval))

        return df_tdocs

    def get_original_and_final_tdocs(df_tdocs):
        print('TDocsByAgenda: Tracking original/final tdocs')

        # Final pass to write the original predecessors and final children
        # Added .astype(str) to avoid some exceptions
        df_tdocs['Original TDocs'] = df_tdocs['Revision of'].astype(str) + ',' + df_tdocs['Merge of'].astype(str)
        df_tdocs['Final TDocs'] = df_tdocs['Revised to'].astype(str) + ',' + df_tdocs['Merged to'].astype(str)

        for index in df_tdocs.index:
            original_tdocs = df_tdocs.at[index, 'Original TDocs']
            final_tdocs = df_tdocs.at[index, 'Final TDocs']

            if original_tdocs != '':
                if original_tdocs == ',':
                    df_tdocs.at[index, 'Original TDocs'] = index
                else:
                    df_tdocs.at[index, 'Original TDocs'] = tdocs_by_agenda.get_original_tdocs(
                        original_tdocs,
                        df_tdocs,
                        index, 0).replace(',', ', ')
            if final_tdocs != '':
                if final_tdocs == ',':
                    df_tdocs.at[index, 'Final TDocs'] = index
                else:
                    df_tdocs.at[index, 'Final TDocs'] = tdocs_by_agenda.get_final_tdocs(
                        final_tdocs,
                        df_tdocs,
                        index,
                        0).replace(',', ', ')

        print('TDocsByAgenda: Finished tracking original/final tdocs')

    # Given a TDoc, returns the TDoc or TDocs that originated this TDoc
    def get_original_tdocs(tdocs, df_tdocs, original_index, n_recursion):
        tdocs_split = tdocs.split(',')
        if len(tdocs_split) > 1:
            tdocs_split = [e for e in tdocs_split if (e != '') and (e is not None)]
        if len(tdocs_split) > 1:
            return join_results(tdocs_split, df_tdocs, tdocs_by_agenda.get_original_tdocs, original_index, n_recursion)

        # We know that length is 1
        tdoc = tdocs_split[0].strip()

        # Fix for 137E final TDocsByAgenda
        if n_recursion > max_recursion:
            print('Maximum recursion reached ({0}) for {1}. Stopping search.'.format(max_recursion, original_index))
            return tdoc

        if tdoc not in df_tdocs.index:
            return tdoc

        try:
            revision_of = df_tdocs.at[tdoc, 'Revision of']
            merge_of = df_tdocs.at[tdoc, 'Merge of']
        except:
            # Case when the tdoc is not found
            print("Original TDoc: '{0}' not found. Stopping recursive search".format(tdoc))
            return tdoc

        # Check for circular reference (found one in the final TDocsByAgenda of 137E)
        if revision_of == original_index:
            revision_of = ''
        if merge_of == original_index:
            merge_of = ''

        if revision_of == '' and merge_of == '':
            return tdoc

        # Merge a list of "Revision of and Merge of" and execute recursively
        if revision_of == '' or merge_of == '':
            all_parents = revision_of + merge_of
        else:
            all_parents = ', '.join([revision_of, merge_of])

        return sort_and_remove_duplicates_from_list(
            tdocs_by_agenda.get_original_tdocs(all_parents, df_tdocs, original_index, n_recursion + 1))

    # Given a TDoc, returns the TDoc or TDocs that ultimately originate from this TDoc
    def get_final_tdocs(tdocs, df_tdocs, original_index, n_recursion):
        tdocs_split = tdocs.split(',')
        if len(tdocs_split) > 1:
            tdocs_split = [e for e in tdocs_split if (e != '') and (e is not None)]
        if len(tdocs_split) > 1:
            return join_results(tdocs_split, df_tdocs, tdocs_by_agenda.get_final_tdocs, original_index, n_recursion)

        # We know that length is 1
        tdoc = tdocs_split[0].strip()

        # Fix for 137E final TDocsByAgenda
        if n_recursion > max_recursion:
            print('Maximum recursion reached ({0}) for {1}. Stopping search.'.format(max_recursion, original_index))
            return tdoc

        if tdoc not in df_tdocs.index:
            tdoc = tdocs_by_agenda.try_to_correct_tdoc_typo(tdoc)

        try:
            revisions = df_tdocs.at[tdoc, 'Revised to']
            merges = df_tdocs.at[tdoc, 'Merged to']
        except:
            # Case when the tdoc is not found
            print("Final TDoc: '{0}' not found. Stopping recursive search".format(tdoc))
            return tdoc

        # This means that the TDoc has no more children
        if revisions == '' and merges == '':
            return tdoc

        # Merge a list of "Revised to and Merged to" and execute recursively
        if revisions == '' or merges == '':
            all_children = revisions + merges
        else:
            all_children = ', '.join([revisions, merges])

        return sort_and_remove_duplicates_from_list(
            tdocs_by_agenda.get_final_tdocs(all_children, df_tdocs, original_index, n_recursion + 1))

    def try_to_correct_tdoc_typo(tdoc):
        if tdoc not in tdocs_by_agenda.tdoc_typos.keys():
            return tdoc
        return tdocs_by_agenda.tdoc_typos[tdoc]


def get_tdocs_by_agenda_with_cache(path_or_html, meeting_server_folder='') -> tdocs_by_agenda:
    if (path_or_html is None) or (path_or_html == ''):
        print('Parse TDocsByAgenda skipped. path_or_html={0}, meeting_server_folder={1}'.format(path_or_html,
                                                                                                meeting_server_folder))
        return None

    global tdocs_by_document_cache

    print('Retrieving TDocsByAgenda (cache is enabled)')

    # If this is an HTML
    if len(path_or_html) > 1000:
        print('TDocsByAgenda retrieval based on HTML content')

        # Changed to hashlib as it is reinitialized beween sessions.
        # See https://stackoverflow.com/questions/27522626/hash-function-in-python-3-3-returns-different-results-between-sessions
        m = hashlib.md5()
        m.update(path_or_html)
        html_hash = m.hexdigest()

        # Retrieve
        if html_hash in tdocs_by_document_cache:
            print('Retrieving TdocsByAgenda from parsed document cache: {0}'.format(html_hash))
            last_tdocs_by_agenda = tdocs_by_document_cache[html_hash]
        else:
            print('TdocsByAgenda {0} not in cache'.format(html_hash))
            last_tdocs_by_agenda = tdocs_by_agenda(
                path_or_html,
                html_hash=html_hash,
                meeting_server_folder=meeting_server_folder)

            # I found out tht this was not a good idea for the inbox. File cache should be enough
            # print('Storing TdocsByAgenda with hash {0} in memory cache'.format(html_hash))
            # tdocs_by_document_cache[html_hash] = last_tdocs_by_agenda

            # Save TDocsByAgenda data in a pickle file so that we can plot graphs later on
            try:
                data_to_save = {
                    'contributor_columns': last_tdocs_by_agenda.contributor_columns,
                    'others_cosigners': last_tdocs_by_agenda.others_cosigners,
                    'tdocs': last_tdocs_by_agenda.tdocs,
                    'cache_version': current_cache_version
                }

                cache_file_name = get_cache_filepath(meeting_server_folder, html_hash)
                if cache_file_name is not None and not os.path.exists(cache_file_name):
                    with open(cache_file_name, 'wb') as f:
                        # Pickle the 'data' dictionary using the highest protocol available.
                        pickle.dump(data_to_save, f, pickle.HIGHEST_PROTOCOL)
                        print('Saved TDocsByAgenda cache to file {0}'.format(cache_file_name))
            except:
                print('Could not cache TDocsByAgenda for meeting {0}'.format(meeting_server_folder))
                print('Object to serialize:')
                print(data_to_save)
                traceback.print_exc()
    else:
        # Path-based fetching uses no hash
        print('TDocsByAgenda retrieval based on path')
        the_tdocs_by_agenda = tdocs_by_agenda(path_or_html)
        last_tdocs_by_agenda = the_tdocs_by_agenda

    return last_tdocs_by_agenda