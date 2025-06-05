import concurrent.futures
import datetime
import os.path
import re
import time
from tkinter import BooleanVar
from typing import List, Tuple, Dict

import parsing.word.pywin32
import tdoc.utils
from application.os import startfile
from application.zip_files import unzip_files_in_zip_file
from config.meetings import MeetingConfig
from server.common import (download_file_to_location, FileToDownload, batch_download_file_to_location, \
                           meeting_pages_per_group,
                           meeting_ftp_pages_per_group, MeetingEntry, DownloadedTdocDocument, DownloadedData)
from utils.local_cache import get_meeting_list_folder, convert_html_file_to_markup, file_exists

# If more than this number of files are included in a zip file, the folder is opened instead.
# Some TDocs, especially in plenary, could contain many, many TDocs, e.g. SP-230457 (22 documents)
maximum_number_of_files_to_open = 5

initialized = False
local_cache_folder = ''

# Group name as key
html_cache_files: Dict[str, str] = {}
html_cache_files_ftp: Dict[str, str] = {}
markup_cache_files: Dict[str, str] = {}
markup_cache_files_ftp: Dict[str, str] = {}


def initialize():
    print(f'Starting meeting list one-time initialization')
    start = time.time()
    global initialized, local_cache_folder, html_cache_files, markup_cache_files, \
        html_cache_files_ftp, markup_cache_files_ftp
    local_cache_folder = get_meeting_list_folder()

    # 3GPP meeting pages
    html_cache_files = {k: os.path.join(local_cache_folder, k + '.htm') for k, v in
                        meeting_pages_per_group.items()}
    markup_cache_files = {k: os.path.join(local_cache_folder, k + '.md') for k, v in
                          meeting_pages_per_group.items()}

    # 3GPP FTP server pages
    html_cache_files_ftp = {k: os.path.join(local_cache_folder, k + '_ftp.htm') for k, v in
                            meeting_pages_per_group.items()}
    markup_cache_files_ftp = {k: os.path.join(local_cache_folder, k + '_ftp.md') for k, v in
                              meeting_pages_per_group.items()}
    end = time.time()
    print(f'Finished meeting list one-time initialization ({end-start:0.2f}s)')
    initialized = True


# Example parsing of:
#   - [SP-102](https://portal.3gpp.org/Home.aspx#/meeting?MtgId=60012) | 3GPPSA#102| [Edinburgh](/../../../\\ftp\\TSG_SA\\TSG_SA\\TSGS_102_Edinburgh_2023-12\\Invitation/)| [2023-12-11](/../../../\\ftp\\TSG_SA\\TSG_SA\\TSGS_102_Edinburgh_2023-12\\Agenda/)| [2023-12-15](/../../../\\ftp\\TSG_SA\\TSG_SA\\TSGS_102_Edinburgh_2023-12\\Report/)| [SP-231205 - SP-231807](/../../../\\ftp\\TSG_SA\\TSG_SA\\TSGS_102_Edinburgh_2023-12\\\\docs\\)[full document list](https://portal.3gpp.org/ngppapp/TdocList.aspx?meetingId=60012) | - | [Participants](https://webapp.etsi.org/3GPPRegistration/fViewPart.asp?mid=60012)| [Files](/../../../\\ftp\\TSG_SA\\TSG_SA\\TSGS_102_Edinburgh_2023-12\\) | - | -
# Filters out following parameters:
#   - meeting_group: SP
#   - meeting_number: 102
#   - meeting_url_3gu: https://portal.3gpp.org/Home.aspx#/meeting?MtgId=60012
#   - meeting_name: 3GPPSA#102
#   - meeting_location: Edinburgh
#   - meeting_url_invitation: /../../../\\\\ftp\\\\TSG_SA\\\\TSG_SA\\\\TSGS_102_Edinburgh_2023-12\\\\Invitation/
#   - start_year: 2023
#   - start_month: 12
#   - start_day: 11
#   - meeting_url_agenda: /../../../\\\\ftp\\\\TSG_SA\\\\TSG_SA\\\\TSGS_102_Edinburgh_2023-12\\\\Agenda/
#   - end_year: 2023
#   - end_month: 12
#   - end_day: 15
#   - meeting_url_report: /../../../\\\\ftp\\\\TSG_SA\\\\TSG_SA\\\\TSGS_102_Edinburgh_2023-12\\\\Report/
#   - tdoc_start: SP-231205
#   - tdoc_end: SP-231807
#   - meeting_url_docs: /../../../\\\\ftp\\\\TSG_SA\\\\TSG_SA\\\\TSGS_102_Edinburgh_2023-12\\\\\\\\docs\\\\
meeting_regex = re.compile(
    r'\[(?P<meeting_group>[a-zA-Z][\d\w]+)\-(?P<meeting_number>[\d\-\w ]+)\]\((?P<meeting_url_3gu>[^ ]+)\)[ ]?\|[ ]?(?P<meeting_name>[^ ]+)[ ]?\|[ ]?\[(?P<meeting_location>[^\]]+)\]\((?P<meeting_url_invitation>[^ ]+)\)[ ]?\|[ ]?\[(?P<start_year>[\d]+)\-(?P<start_month>[\d]+)\-(?P<start_day>[\d]+)\]\((?P<meeting_url_agenda>[^ ]+)\)[ ]?\|[ ]?\[(?P<end_year>[\d]+)\-(?P<end_month>[\d]+)\-(?P<end_day>[\d]+)\]\((?P<meeting_url_report>[^ ]+)\)[ ]?\|[ ]?\[(?P<tdoc_start>[\w\-\d]+)[ -]+(?P<tdoc_end>[\w\-\d]+)\]\((?P<meeting_url_docs>[^ ]+)\).*\[(Files)\]\((?P<files_url>[^ ]+)\)')

# Meetings such as this one:
# [SP-103](https://portal.3gpp.org/Home.aspx#/meeting?MtgId=60295) | 3GPPSA#103 | [Maastricht](/../../../\\ftp\\TSG_SA\\TSG_SA\\TSGS_103_Maastricht_2024-03\\Invitation/) | [2024-03-19](/../../../\\ftp\\TSG_SA\\TSG_SA\\TSGS_103_Maastricht_2024-03\\Agenda/) | 2024-03-22 | [SP-240001 - SP-240285](/../../../\\ftp\\TSG_SA\\TSG_SA\\TSGS_103_Maastricht_2024-03\\\\docs\\)[ full document list](https://portal.3gpp.org/ngppapp/TdocList.aspx?meetingId=60295) | [Register](https://webapp.etsi.org/3GPPRegistration/fMain.asp?mid=60295) | [Participants](https://webapp.etsi.org/3GPPRegistration/fViewPart.asp?mid=60295) | [Files](/../../../\\ftp\\TSG_SA\\TSG_SA\\TSGS_103_Maastricht_2024-03\\) | [ICS](https://portal.3gpp.org/webapp/meetingCalendar/ical.asp?qMTG_ID=60295) | -
meeting_without_report_regex = re.compile(
    r"\[(?P<meeting_group>[a-zA-Z][\d\w]+)\-(?P<meeting_number>[\d\-\w ]+)\]\((?P<meeting_url_3gu>[^ ]+)\)[ ]?\|[ ]?(?P<meeting_name>[^ ]+)[ ]?\|[ ]?\[(?P<meeting_location>[^\]]+)\]\((?P<meeting_url_invitation>[^ ]+)\)[ ]?\|[ ]?\[(?P<start_year>[\d]+)\-(?P<start_month>[\d]+)\-(?P<start_day>[\d]+)\]\((?P<meeting_url_agenda>[^ ]+)\)[ ]?\|[ ]?(?P<end_year>[\d]+)\-(?P<end_month>[\d]+)\-(?P<end_day>[\d]+)[ ]?\|[ ]?\[(?P<tdoc_start>[\w\-\d]+)[ -]+(?P<tdoc_end>[\w\-\d]+)\]\((?P<meeting_url_docs>[^ ]+)\).*\[(Files)\]\((?P<files_url>[^ ]+)\)")

# Meetings such as this one:
# [S2-169](https://portal.3gpp.org/Home.aspx#/meeting?MtgId=60632) | 3GPPSA2#169 | [Japan](/../../..//Specification-Groups/) | 2025-05-19 | 2025-05-23 | \- | [Register](https://webapp.etsi.org/3GPPRegistration/fMain.asp?mid=60632) | [Participants](https://webapp.etsi.org/3GPPRegistration/fViewPart.asp?mid=60632) | - | [ICS](https://portal.3gpp.org/webapp/meetingCalendar/ical.asp?qMTG_ID=60632) | -
meeting_without_invitation_regex = re.compile(
    r"\[(?P<meeting_group>[a-zA-Z][\d\w]+)\-(?P<meeting_number>[\d\-\w ]+)\]\((?P<meeting_url_3gu>[^ ]+)\)[ ]?\|[ ]?(?P<meeting_name>[^ ]+)[ ]?\|[ ]?\[?(?P<meeting_location>[^\]]+)\]\((?P<meeting_url_invitation>[^ ]+)\) \| (?P<start_year>[\d]+)\-(?P<start_month>[\d]+)\-(?P<start_day>[\d]+) \| (?P<end_year>[\d]+)\-(?P<end_month>[\d]+)\-(?P<end_day>[\d]+) [\\\-\| ]+\[(Register|Participants)\]")

# Meetings such as this one:
# [S6-62-AdHoc-e](https://portal.3gpp.org/Home.aspx#/meeting?MtgId=60688) | 3GPPSA6#62-AdHoc-e | [Online](/../../../\\ftp\\tsg_sa\\WG6_MissionCritical\\Ad-hoc_meetings\\2024-07-10_adhoc\\Invitation/) | [2024-07-10](/../../../\\ftp\\tsg_sa\\WG6_MissionCritical\\Ad-hoc_meetings\\2024-07-10_adhoc\\Agenda/) | 2024-07-18 | - | [Register](https://webapp.etsi.org/3GPPRegistration/fMain.asp?mid=60688) | [Participants](https://webapp.etsi.org/3GPPRegistration/fViewPart.asp?mid=60688) | [Files](/../../../\\ftp\\tsg_sa\\WG6_MissionCritical\\Ad-hoc_meetings\\2024-07-10_adhoc\\) | [ICS](https://portal.3gpp.org/webapp/meetingCalendar/ical.asp?qMTG_ID=60688) | -
meeting_sa6_adhocs = re.compile(
    r"\[(?P<meeting_group>[a-zA-Z][\d\w]+)\-(?P<meeting_number>[\d\-\w ]+)\]\((?P<meeting_url_3gu>[^ ]+)\)[ ]?\|[ ]?(?P<meeting_name>[^ ]+)[ ]?\|[ ]?\[?(?P<meeting_location>[^\]]+)\]\((?P<meeting_url_invitation>[^ ]+)\) \| \[(?P<start_year>[\d]+)\-(?P<start_month>[\d]+)\-(?P<start_day>[\d]+)\]\((?P<meeting_url_agenda>[^ ]+)\)[ ]?\|[ ]?(?P<end_year>[\d]+)\-(?P<end_month>[\d]+)\-(?P<end_day>[\d]+) \| - \| \[Register\]"
)

# Used to split the generated Markup text
meeting_split_regex = re.compile(r'(\[[a-zA-Z][\d\w]+\-[\d\-\w ]+\]\([^ ]+\))')

# Loaded meeting entries
loaded_meeting_entries: List[MeetingEntry] = []


def get_meeting_groups() -> List[str]:
    """
    The possible 3GPP groups that can be queried. Based on a hard-coded list of 3GPP Working Groups (WGs)
    Returns: A list of strings
    """
    return [k for k, v in meeting_pages_per_group.items()]

def get_meeting_years() -> List[int]:
    """
        The years of the start date of the loaded meetings
        Returns: A list of years
        """
    start_years = list(set(
        [m.start_date.year for m in loaded_meeting_entries
         if m is not None and m.start_date is not None]))
    start_years.sort(reverse=True)
    return start_years


def update_local_html_cache(redownload_if_exists=False) -> List[str]:
    """
    Download the meeting files to the cache

    Args:
        redownload_if_exists: Whether to force a download of the file(s) if they exist
    Returns: The groups that were downloaded
    """
    if not initialized:
        initialize()
    print('Updating local cache')
    files_to_download: List[FileToDownload] = []
    downloaded_group_meetings: List[str] = []

    for k, v in meeting_pages_per_group.items():
        # Download meeting page
        local_file = html_cache_files[k]
        if redownload_if_exists or not os.path.exists(local_file):
            files_to_download.append(FileToDownload(
                remote_url=v,
                local_filepath=local_file,
                force_download=True
            ))
            downloaded_group_meetings.append(k)
        else:
            print(f'Skipping download of {k} group meeting page to {local_file}')

    for k, v in meeting_ftp_pages_per_group.items():
        # Download FTP server page
        local_file = html_cache_files_ftp[k]
        if redownload_if_exists or not os.path.exists(local_file):
            files_to_download.append(FileToDownload(
                remote_url=v,
                local_filepath=local_file,
                force_download=True
            ))
            downloaded_group_meetings.append(k)
        else:
            print(f'Skipping download of {k} group FTP page to {local_file}')

    batch_download_file_to_location(files_to_download, cache=True)
    return downloaded_group_meetings


def filter_markdown_text(markdown_text: str) -> str:
    """
    Further filtering (TDoc-specific) for the markdown text
    Args:
        markdown_text: The input markdown text

    Returns: Filtered and clean-up text
    """
    text_lines = meeting_split_regex.split(markdown_text)
    text_lines = [line.replace('\n', '').replace('\r', '').replace("""‑""", '-') for line in text_lines]
    full_text = '\n'.join(text_lines)

    full_text = full_text.replace(""")
 | """, ") | ")

    full_text = full_text.replace('  ', ' ').replace(""")
[ full""", ')[ full').replace("""| 
[Sophia""", '| [Sophia').replace(""")
 [ full document list]""", ')[ full document list]').replace("""\\Report/) | 
[""", """\\Report/) | [""").replace("""| 
[Stor-Göteborg]""", "| [Göteborg]")

    # Catches when the report is not yet ready
    full_text = re.sub(r"(\d\d\d\d-\d\d-\d\d) \| [\r\n]{1,}\[", r"\1 | [", full_text, flags=re.M)
    full_text = re.sub(r"[Aa]d [Hh]oc", r"AdHoc", full_text)

    return full_text


markdown_url_regex = r'\[(.*?)\]\((.*?)\)'


def filter_markdown_urls(markdown_text: str) -> str:
    # Regex pattern to match Markdown links: [text](url)
    links = re.finditer(markdown_url_regex, markdown_text)
    links_tuple = [(e.group(1), e.group(2)) for e in links]

    def remove_duplicates(tuples_list):
        return list({item[0]: item for item in tuples_list}.values())

    links_tuple = remove_duplicates(links_tuple)

    to_ignore = ["ETSI logo", "www.3gpp.org", "ftp", "upload", "sort by name", "desc", "sort by date", "sort by size",
                 "icon", "Draft_Specs", "Latest_Templates"]
    links_tuple = [e for e in links_tuple if e[0] not in to_ignore]

    out_list = [f"[{e[0]}]({e[1]})" for e in links_tuple]
    out_str = '\n'.join(out_list)
    return out_str


def convert_local_cache_to_markdown(downloaded_groups: List[str], force_conversion=False):
    """
        Convert local cache to markdown
    """
    if not initialized:
        initialize()
    start = time.time()

    print(f'Converting local cache to markdown (meeting pages)')
    for k, v in html_cache_files.items():
        if os.path.exists(v):
            if force_conversion or (not os.path.exists(markup_cache_files[k]) or k in downloaded_groups):
                print(f'Markup conversion for {k} group')
                convert_html_file_to_markup(
                    v,
                    output_path=markup_cache_files[k],
                    ignore_links=False,
                    filter_text_function=filter_markdown_text
                )
            else:
                print(f'Skipped Markup conversion for {k} group')

    print(f'Converting local cache to markdown (FTP server pages)')
    for k, v in html_cache_files_ftp.items():
        if os.path.exists(v):
            if force_conversion or (not os.path.exists(markup_cache_files_ftp[k]) or k in downloaded_groups):
                print(f'Markup conversion for {k} group')
                convert_html_file_to_markup(
                    v,
                    output_path=markup_cache_files_ftp[k],
                    ignore_links=False,
                    filter_text_function=lambda x: filter_markdown_urls(filter_markdown_text(x))
                )
            else:
                print(f'Skipped Markup conversion for {k} group')
    end = time.time()
    print(f'Finished converting local cache to markdown ({end - start:.2f}s)')


def load_markdown_cache_to_memory(groups: List[str] = None):
    """
    Parses the markdown cache files and returns the parsed 3GPP meeting list.
    Returns: 3GPP meeting list

    """
    if not initialized:
        initialize()

    start = time.time()
    print(f'Loading markdown meeting cache')
    global loaded_meeting_entries
    loaded_meeting_entries = []
    loaded_meeting_entries.extend(MeetingConfig.additional_meetings)

    groups_to_load = markup_cache_files.items()
    if groups is None or len(groups) == 0:
        # Load all
        pass
    else:
        groups_to_load = [kvp for kvp in groups_to_load if kvp[0] in groups]

    groups_to_load_str = ', '.join([k for k, v in groups_to_load])
    end = time.time()

    # Enrich information from meeting site with information from the FTP page
    meeting_info_ftp_parsed = dict()
    for group_to_check, v in groups_to_load:
        print(f'Loading FTP info for meetings for group {group_to_check}')
        file_ftp_markdown = markup_cache_files_ftp[group_to_check]

        if not os.path.exists(file_ftp_markdown):
            # Not al groups may exist
            continue

        with open(file_ftp_markdown, 'r', encoding='utf-8') as file:
            markup_file_content = file.read()
        links = re.finditer(markdown_url_regex, markup_file_content)

        def cleanup_match(match_in: re.Match) -> str | None:
            if match_in is None:
                return None
            return match_in.group(1)

        links_tuple = [(
            cleanup_match(re.search(r"_(\d+[-eEbisBISaAhH]*)", e.group(1))),
            e.group(1),
            e.group(2)) for e in links]

        meeting_info_ftp_parsed[group_to_check] = links_tuple

    print(f'Loading meeting entries from meeting list: {groups_to_load_str} ({end - start:.2f})')

    def server_url_replace(a_url: str | None) -> str | None:
        """
        Cleans up the URL and returns an absolute URL pointing to the 3GPP HTTP(s) file server
        Args:
            a_url: A relative URL as parsed form the markdown cache,
            e.g. /../../../\\ftp\\tsg_sa\\TSG_SA\\TSGS_60\\Invitation/

        Returns: An absolute URL

        """
        if a_url is None:
            return a_url
        return (a_url
                .replace('ftpTSG_SA',
                         'ftp/TSG_SA')  # See https://www.3gpp.org/dynareport?code=Meetings-S5.htm for SA5#154
                .replace(r'\\', '/')
                .replace('/../../..//ftp/', 'https://www.3gpp.org/ftp/')
                .replace('//', '/')
                .replace(':/', '://'))

    regex_list: List[re.Pattern] = [
        meeting_regex,
        meeting_without_report_regex,
        meeting_without_invitation_regex,
        meeting_sa6_adhocs]

    def parse_match_to_meeting_entry(a_match: re.Match,
                                     group_meetings_ftp: List[Tuple[str | None, str, str]]) -> MeetingEntry:
        m_matches: Dict[str, str] = a_match.groupdict()
        meeting_group = None
        meeting_number = None
        meeting_url_3gu = None
        meeting_name = None
        meeting_location = None
        meeting_url_invitation = None
        start_date = None
        meeting_url_agenda = None
        end_date = None
        meeting_url_report = None
        tdoc_start = None
        tdoc_end = None
        meeting_url_docs = None
        meeting_folder_url = None

        for k, v in m_matches.items():
            match k:
                case 'meeting_group':
                    meeting_group = v
                case 'meeting_number':
                    meeting_number = v
                case 'meeting_url_3gu':
                    meeting_url_3gu = server_url_replace(v)
                case 'meeting_name':
                    meeting_name = v
                case 'meeting_location':
                    meeting_location = v
                case 'meeting_url_invitation':
                    meeting_url_invitation = server_url_replace(v)
                case 'start_year':
                    start_date = datetime.datetime(
                        year=int(v),
                        month=int(m_matches['start_month']),
                        day=int(m_matches['start_day']))
                case 'meeting_url_agenda':
                    meeting_url_agenda = server_url_replace(v)
                case 'end_year':
                    end_date = datetime.datetime(
                        year=int(v),
                        month=int(m_matches['end_month']),
                        day=int(m_matches['end_day']))
                case 'meeting_url_report':
                    meeting_url_report = server_url_replace(v)
                case 'tdoc_start':
                    tdoc_start = tdoc.utils.is_generic_tdoc(v)
                case 'tdoc_end':
                    tdoc_end = tdoc.utils.is_generic_tdoc(v)
                case 'meeting_url_docs':
                    meeting_url_docs = server_url_replace(v)
                case 'files_url':
                    meeting_folder_url = server_url_replace(v)

        if meeting_folder_url is None and group_meetings_ftp is not None:
            # Try to add information parsed from FTP server
            try:
                matching_meeting = [m for m in group_meetings_ftp if m[0] == meeting_number][0]
                print(f"Match in FTP server for {meeting_name}: {matching_meeting}")
                meeting_folder_url = matching_meeting[2]
                meeting_url_docs = server_url_replace(f"{meeting_folder_url}{'/Docs/'}")
            except Exception as e:
                pass

        return MeetingEntry(
            meeting_group=meeting_group,
            meeting_number=meeting_number,
            meeting_url_3gu=meeting_url_3gu,
            meeting_name=meeting_name,
            meeting_location=meeting_location,
            meeting_url_invitation=meeting_url_invitation,
            start_date=start_date,
            meeting_url_agenda=meeting_url_agenda,
            end_date=end_date,
            meeting_url_report=meeting_url_report,
            tdoc_start=tdoc_start,
            tdoc_end=tdoc_end,
            meeting_url_docs=meeting_url_docs,
            meeting_folder_url=meeting_folder_url
        )

    for k, v in groups_to_load:
        if os.path.exists(v):
            print(f'Loading meetings for group {k}')
            with open(v, 'r', encoding='utf-8') as file:
                markup_file_content = file.read()

            try:
                group_meetings_ftp = meeting_info_ftp_parsed[k]
            except KeyError:
                group_meetings_ftp = []

            # Check different regex patterns
            parsed_meetings_for_k: List[MeetingEntry] = []
            for regex_to_check in regex_list:
                meeting_matches = regex_to_check.finditer(markup_file_content)
                already_parsed_meetings = [m.meeting_number for m in parsed_meetings_for_k]
                matches_to_process = [m for m in meeting_matches
                                      if m is not None and m.group('meeting_number') not in already_parsed_meetings]

                meetings_to_add = [parse_match_to_meeting_entry(m, group_meetings_ftp) for m in matches_to_process]
                loaded_meeting_entries.extend(meetings_to_add)
        else:
            print(f'Not found: {v}')

    end = time.time()
    print(f'Finished loading meetings ({end - start:.2f}s)')


def group_is_li(group_name: str) -> bool:
    return group_name.lower() == 's3i'


def search_meeting_for_tdoc(
        tdoc_str: str,
        return_last_meeting_if_tdoc_is_new: bool = False
) -> MeetingEntry | None:
    """
    Searches for a specific TDoc in the loaded meeting list
    Args:
        return_last_meeting_if_tdoc_is_new: Allows you to specify that the last meeting with allocated TDocs should be
        returned if no meeting is found. e.g., RP-241153 is actually TSGR_104, although the meeting officially only got
        allocated from RP-240861 to RP-240906 (2025.06.11)
        tdoc_str: A TDoc ID

    Returns: A meeting containing this TDoc. None if none found or if the input TDoc is not a TDoc.

    """
    parsed_tdoc = tdoc.utils.is_generic_tdoc(tdoc_str)
    if parsed_tdoc is None:
        return None

    # Whether a SA3 meeting is LI is encoded in the meeting number
    if group_is_li(parsed_tdoc.group):
        group_to_search = 'S3'
        sa3_li_tdoc = True
    else:
        group_to_search = parsed_tdoc.group
        sa3_li_tdoc = False

    print(f'Searching for group {parsed_tdoc.group}, tdoc {parsed_tdoc.number}. LI WG: {sa3_li_tdoc}')

    def group_match(m: MeetingEntry, group_str: str):

        if sa3_li_tdoc:
            return (group_str == m.meeting_group) and m.is_li

        return (group_str == m.meeting_group) and not m.is_li

    group_meetings = [m for m in loaded_meeting_entries if group_match(m, group_to_search)]
    print(f'{len(group_meetings)} Group meetings for group {group_to_search}. LI: {sa3_li_tdoc}')
    matching_meetings = [m for m in group_meetings if m.tdoc_start is not None and m.tdoc_end is not None and
                         m.tdoc_start.number <= parsed_tdoc.number <= m.tdoc_end.number]

    if len(matching_meetings) > 0:
        matching_meeting = matching_meetings[0]
        print(f'Matching meeting found for TDoc {tdoc_str}: {matching_meeting.meeting_name}')
    else:
        if return_last_meeting_if_tdoc_is_new:
            matching_meetings = [m for m in group_meetings if m.tdoc_start is not None and m.tdoc_end is not None and
                                 m.tdoc_start.number <= parsed_tdoc.number and m.tdoc_end.number <= parsed_tdoc.number]
        if len(matching_meetings) > 0:
            matching_meetings.sort(key=lambda x: x.end_date)
            matching_meeting = matching_meetings[-1]
            print(f'Set meeting for TDoc {tdoc_str} as last meeting with available documents: '
                  f'{matching_meeting.meeting_name}')
        else:
            matching_meeting = None
            print(f'Matching meeting NOT found for TDoc {tdoc_str}')

    return matching_meeting


def fully_update_cache(redownload_if_exists=False):
    """
    Fully updates the meeting list, which includes downloading from the 3GPP server the meetings for all WGs.

    Args:
        redownload_if_exists: Whether to re-download the files even if they exist

    """
    print('Triggering update of local cache')
    downloaded_groups = update_local_html_cache(redownload_if_exists=redownload_if_exists)
    convert_local_cache_to_markdown(downloaded_groups, force_conversion=redownload_if_exists)
    load_markdown_cache_to_memory()
    print('Finished update of local cache')


def search_download_and_open_tdoc(
        tdoc_str: str,
        skip_open=False,
        tkvar_3gpp_wifi_available: BooleanVar|None=None,
        tdoc_meeting: MeetingEntry =None,
) -> DownloadedData:
    """
    Searches for a given TDoc. If the zip file contains many files (e.g. typical for plenary CR packs), it will only
    open the folder.
    Args:
        tdoc_meeting: The meeting for this TDoc if known (so we do not have to search for it)
        tkvar_3gpp_wifi_available: Whether we should use a private server if available
        skip_open: Whether to skip opening the files
        tdoc_str: The TDoc ID

    Returns: The files that were opened

    """
    if tdoc_str is None or tdoc_str == '':
        return DownloadedData(None, None)

    if tdoc_meeting is None:
        print(f'Searching for TDoc {tdoc_str}. Unknown meeting (will search for meeting)')
    else:
        print(f'Searching for TDoc {tdoc_str}. Meeting {tdoc_meeting.meeting_name}')

    # Cleanup
    tdoc_str = tdoc.utils.cleanup_tdoc(tdoc_str)

    # Load data if needed
    if len(loaded_meeting_entries) == 0:
        fully_update_cache()

    if tdoc_meeting is None:
        tdoc_meeting = search_meeting_for_tdoc(tdoc_str, return_last_meeting_if_tdoc_is_new=True)
    if tdoc_meeting is None:
        return DownloadedData(None, None)

    in_3gpp_wifi = False
    if tkvar_3gpp_wifi_available is not None and tkvar_3gpp_wifi_available.get():
        in_3gpp_wifi = True
    if (not tdoc_meeting.meeting_is_now or
            not in_3gpp_wifi):
        print(f'Opening {tdoc_str} from remote server. '
              f'Meeting is now: {tdoc_meeting.meeting_is_now}, In 3GPP Wifi: {in_3gpp_wifi}')
        use_private_server = False
    else:
        print(f'Opening {tdoc_str} from local server')
        use_private_server = True

    if not use_private_server:
        tdoc_urls = [tdoc_meeting.get_tdoc_url(tdoc_str), tdoc_meeting.get_tdoc_inbox_url(tdoc_str)]
    else:
        tdoc_urls = tdoc_meeting.get_tdoc_3gpp_wifi_url(tdoc_str)
    local_target = tdoc_meeting.get_tdoc_local_path(tdoc_str)

    print(f'Will download: {tdoc_urls}')

    # Only download file if needed
    downloaded_tdoc_url = ''
    if not file_exists(local_target):
        for tdoc_url in tdoc_urls:
            print(f'Downloading {tdoc_url} to {local_target}')
            if download_file_to_location(tdoc_url, local_target):
                print('File successfully downloaded')
                downloaded_tdoc_url = tdoc_url
                break
    else:
        print(f'Using local cache for {tdoc_urls} in {local_target}')

    if not file_exists(local_target):
        print(f'No file to open in {local_target}')
        return DownloadedData(None, None)

    files_in_zip = unzip_files_in_zip_file(local_target)
    if (len(files_in_zip) <= maximum_number_of_files_to_open) and (not skip_open):
        folder_to_open, first_file = os.path.split(files_in_zip[0])
        opened_files, metadata_list = parsing.word.pywin32.open_files(files_in_zip, return_metadata=True)
        metadata_list = [DownloadedTdocDocument(
            title=m.title,
            source=m.source,
            url=downloaded_tdoc_url,
            tdoc_id=tdoc_str,
            path=m.path)
            for m in metadata_list if m is not None]
    else:
        folder_to_open, first_file = os.path.split(files_in_zip[0])
        if not skip_open:
            print(
                f'More than {maximum_number_of_files_to_open} contained within {tdoc_str}. Opening folder instead of files')
            startfile(folder_to_open)
        metadata_list = [DownloadedTdocDocument(
            title=None,
            source=None,
            url=None,
            tdoc_id=tdoc_str,
            path=m)
            for m in files_in_zip if m is not None if (m is not None)]
    return DownloadedData(folder_to_open, metadata_list)


def batch_search_and_download_tdocs(
        tdoc_list: List[str],
        tkvar_3gpp_wifi_available=None,
        tdoc_meeting=None
) -> List[DownloadedData]:
    """
    Parallel download of a list of TDocs, e.g. for caching purposes
    Args:
        tdoc_meeting: If available, the 3GPP meeting for the requested files
        tkvar_3gpp_wifi_available: Whether to use the 10.10.10.10 server
        tdoc_list: A list of TDoc IDs
    """
    if tdoc_list is None or not isinstance(tdoc_list, list) or len(tdoc_list) < 1:
        return

    # See https://docs.python.org/3/library/concurrent.futures.html
    all_downloads = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        future_to_dl = {
            executor.submit(
                search_download_and_open_tdoc,
                tdoc_str,
                True,
                tkvar_3gpp_wifi_available,
                tdoc_meeting
            ): tdoc_str for tdoc_str in tdoc_list
        }
        for future in concurrent.futures.as_completed(future_to_dl):
            tdoc_to_download = future_to_dl[future]
            try:
                downloaded_files = future.result()
                if not downloaded_files:
                    print(f'Could not download {tdoc_to_download}')
                else:
                    all_downloads.append(downloaded_files)
            except Exception as exc:
                print('%r generated an exception: %s' % (tdoc_to_download, exc))

    # Return all downloaded files
    return all_downloads


def compare_two_tdocs(
        original_tdoc: str,
        new_tdoc: str,
        tkvar_3gpp_wifi_available: BooleanVar|None=None):
    print(f'Comparing {new_tdoc}  (original) vs. {original_tdoc}')
    opened_docs1_folder, metadata1 = search_download_and_open_tdoc(
        original_tdoc,
        skip_open=False,
        tkvar_3gpp_wifi_available=tkvar_3gpp_wifi_available
    )
    opened_docs2_folder, metadata2 = search_download_and_open_tdoc(
        new_tdoc,
        skip_open=False,
        tkvar_3gpp_wifi_available=tkvar_3gpp_wifi_available
    )
    doc_1 = metadata1[0].path
    doc_2 = metadata2[0].path
    print(f'Comparing {doc_2} vs. {doc_1}')
    parsing.word.pywin32.compare_documents(doc_2, doc_1)
