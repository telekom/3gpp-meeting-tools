import datetime
import os.path
import re
from typing import NamedTuple, List, Tuple, Any, Dict
import parsing.word.pywin32

import tdoc.utils
from application.zip_files import unzip_files_in_zip_file
from server.common import download_file_to_location
import utils
from utils.local_cache import get_meeting_list_folder, convert_html_file_to_markup, get_cache_folder, \
    create_folder_if_needed

# If more than this number of files are included in a zip file, the folder is opened instead.
# Some TDocs, especially in plenary, could contain many, many TDocs, e.g. SP-230457 (22 documents)
maximum_number_of_files_to_open = 5

meeting_pages_per_group: dict[str, str] = {
    'SP': 'https://www.3gpp.org/dynareport?code=Meetings-SP.htm',
    'S1': 'https://www.3gpp.org/dynareport?code=Meetings-S1.htm',
    'S2': 'https://www.3gpp.org/dynareport?code=Meetings-S2.htm',
    'S3': 'https://www.3gpp.org/dynareport?code=Meetings-S3.htm',
    'S4': 'https://www.3gpp.org/dynareport?code=Meetings-S4.htm',
    'S5': 'https://www.3gpp.org/dynareport?code=Meetings-S5.htm',
    'S6': 'https://www.3gpp.org/dynareport?code=Meetings-S6.htm',
    'CP': 'https://www.3gpp.org/dynareport?code=Meetings-CP.htm',
    'C1': 'https://www.3gpp.org/dynareport?code=Meetings-C1.htm',
    'C3': 'https://www.3gpp.org/dynareport?code=Meetings-C3.htm',
    'C6': 'https://www.3gpp.org/dynareport?code=Meetings-C6.htm',
    'RP': 'https://www.3gpp.org/dynareport?code=Meetings-RP.htm',
    'R1': 'https://www.3gpp.org/dynareport?code=Meetings-R1.htm',
    'R2': 'https://www.3gpp.org/dynareport?code=Meetings-R2.htm',
    'R3': 'https://www.3gpp.org/dynareport?code=Meetings-R3.htm',
    'R4': 'https://www.3gpp.org/dynareport?code=Meetings-R4.htm',
    'R5': 'https://www.3gpp.org/dynareport?code=Meetings-R5.htm',
}

local_cache_folder = get_meeting_list_folder()
html_cache: Dict[str, str] = {k: os.path.join(local_cache_folder, k + '.htm') for k, v in
                              meeting_pages_per_group.items()}
markup_cache: Dict[str, str] = {k: os.path.join(local_cache_folder, k + '.md') for k, v in
                                meeting_pages_per_group.items()}
pickle_cache = os.path.join(local_cache_folder, '3gpp_meeting_list.pickle')

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
    r"\[(?P<meeting_group>[a-zA-Z][\d\w]+)\-(?P<meeting_number>[\d\-\w ]+)\]\((?P<meeting_url_3gu>[^ ]+)\)[ ]?\|[ ]?(?P<meeting_name>[^ ]+)[ ]?\|[ ]?\[?(?P<meeting_location>[^\]]+)\]\((?P<meeting_url_invitation>[^ ]+)\) \| (?P<start_year>[\d]+)\-(?P<start_month>[\d]+)\-(?P<start_day>[\d]+) \| (?P<end_year>[\d]+)\-(?P<end_month>[\d]+)\-(?P<end_day>[\d]+) \| \\\- \| \[Register\]")

# Used to split the generated Markup text
meeting_split_regex = re.compile(r'(\[[a-zA-Z][\d\w]+\-[\d\-\w ]+\]\([^ ]+\))')


class MeetingEntry(NamedTuple):
    meeting_group: str
    meeting_number: str
    meeting_url_3gu: str
    meeting_name: str
    meeting_location: str
    meeting_url_invitation: str
    start_date: datetime.datetime
    meeting_url_agenda: str
    end_date: datetime.datetime
    meeting_url_report: str
    tdoc_start: tdoc.utils.GenericTdoc | None
    tdoc_end: tdoc.utils.GenericTdoc | None
    meeting_url_docs: str
    meeting_folder_url: str

    @property
    def meeting_folder(self) -> str:
        """
        The remote meeting folder name in the 3GPP server's group directory based on the meeting_folder URL
        Returns: The remote folder of the meeting in the 3GPP server

        """
        folder_url = self.meeting_folder_url
        if folder_url is None or folder_url == '':
            return folder_url
        split_folder_url = [f for f in folder_url.split('/') if f != '']
        return split_folder_url[-1]

    def get_tdoc_url(self, tdoc_to_get: tdoc.utils.GenericTdoc):
        """
        For a string containingg a potential TDoc, returns a URL concatenating the Docs folder and the input TDoc and
        adds a .'zip' extension.
        Args:
            tdoc_to_get: A TDoc ID. Note that the input is NOT checked!

        Returns: A URL

        """
        tdoc_file = tdoc_to_get.__str__() + '.zip'
        return self.meeting_url_docs + tdoc_file

    def get_local_folder_for_meeting(self, create_dir=False) -> str:
        """
        For a given meeting, returns the cache folder and creates it if it does not exist
        Args:
            create_dir: Whether to create the directory if it does not exist

        Returns:

        """
        folder_name = self.meeting_folder
        full_path = os.path.join(utils.local_cache.get_cache_folder(), folder_name)
        utils.local_cache.create_folder_if_needed(full_path, create_dir=create_dir)
        return full_path


# Loaded meeting entries
loaded_meeting_entries: List[MeetingEntry] = []


def get_meeting_groups() -> List[str]:
    """
    The possible 3GPP groups that can be queried. Based on a hard-coded list of 3GPP Working Groups (WGs)
    Returns: A list of strings
    """
    return [k for k, v in meeting_pages_per_group.items()]


def update_local_html_cache(redownload_if_exists=False):
    """
    Download the meeting files to the cache

    Args:
        redownload_if_exists: Whether to force a download of the file(s) if they exist
    """
    print('Updating local cache')
    for k, v in meeting_pages_per_group.items():
        local_file = html_cache[k]
        if redownload_if_exists or not os.path.exists(local_file):
            download_file_to_location(v, local_file)
        else:
            print(f'Skipping download of {v} to {local_file}')


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
[""", """\\Report/) | [""")

    # Catches when the report is not yet ready
    full_text = re.sub(r"(\d\d\d\d-\d\d-\d\d) \| [\r\n]{1,}\[", r"\1 | [", full_text, flags=re.M)
    full_text = re.sub(r"[Aa]d [Hh]oc", r"AdHoc", full_text)

    return full_text


def convert_local_cache_to_markdown():
    """
        Convert local cache to markdown
    """
    for k, v in html_cache.items():
        if os.path.exists(v):
            convert_html_file_to_markup(
                v,
                output_path=markup_cache[k],
                ignore_links=False,
                filter_text_function=filter_markdown_text
            )


def load_markdown_cache_to_memory(groups: List[str] = None):
    """
    Parses the markdown cache files and returns the parsed 3GPP meeting list.
    Returns: 3GPP meeting list

    """
    global loaded_meeting_entries
    loaded_meeting_entries = []

    items_to_load = markup_cache.items()
    if groups is None or len(groups) == 0:
        # Load all
        pass
    else:
        items_to_load = [kvp for kvp in items_to_load if kvp[0] in groups]

    groups_to_load_str = ', '.join([k for k, v in items_to_load])

    print(f'Loading meeting entries from meeting list: {groups_to_load_str}')

    def server_url_replace(meeting_url: str | None) -> str | None:
        """
        Cleans up the URL and returns an absolute URL pointing to the 3GPP HTTP(s) file server
        Args:
            meeting_url: A relative URL as parsed form the markdown cache,
            e.g. /../../../\\ftp\\tsg_sa\\TSG_SA\\TSGS_60\\Invitation/

        Returns: An absolute URL

        """
        if meeting_url is None:
            return meeting_url
        return (meeting_url
                .replace(r'\\', '/')
                .replace('/../../..//ftp/', 'https://www.3gpp.org/ftp/')
                .replace('//', '/')
                .replace(':/', '://'))

    for k, v in items_to_load:
        if os.path.exists(v):
            print(f'Loading meetings for group {k}')
            with open(v, 'r', encoding='utf-8') as file:
                markup_file_content = file.read()

            # Finished meetings
            meeting_matches = meeting_regex.finditer(markup_file_content)
            meeting_matches_parsed = [
                MeetingEntry(
                    meeting_group=m.group('meeting_group'),
                    meeting_number=m.group('meeting_number'),
                    meeting_url_3gu=server_url_replace(m.group('meeting_url_3gu')),
                    meeting_name=m.group('meeting_name'),
                    meeting_location=m.group('meeting_location'),
                    meeting_url_invitation=server_url_replace(m.group('meeting_url_invitation')),
                    start_date=datetime.datetime(
                        year=int(m.group('start_year')),
                        month=int(m.group('start_month')),
                        day=int(m.group('start_day'))),
                    meeting_url_agenda=server_url_replace(m.group('meeting_url_agenda')),
                    end_date=datetime.datetime(
                        year=int(m.group('end_year')),
                        month=int(m.group('end_month')),
                        day=int(m.group('end_day'))),
                    meeting_url_report=server_url_replace(m.group('meeting_url_report')),
                    tdoc_start=tdoc.utils.is_generic_tdoc(m.group('tdoc_start')),
                    tdoc_end=tdoc.utils.is_generic_tdoc(m.group('tdoc_end')),
                    meeting_url_docs=server_url_replace(m.group('meeting_url_docs')),
                    meeting_folder_url=server_url_replace(m.group('files_url'))
                )
                for m in meeting_matches if m is not None
            ]
            loaded_meeting_entries.extend(meeting_matches_parsed)
            print(f'Added {len(meeting_matches_parsed)} finished meetings to group {k}')

            # Meetings for which a report is not yet ready
            meeting_matches = meeting_without_report_regex.finditer(markup_file_content)
            meeting_matches_parsed = [
                MeetingEntry(
                    meeting_group=m.group('meeting_group'),
                    meeting_number=m.group('meeting_number'),
                    meeting_url_3gu=server_url_replace(m.group('meeting_url_3gu')),
                    meeting_name=m.group('meeting_name'),
                    meeting_location=m.group('meeting_location'),
                    meeting_url_invitation=server_url_replace(m.group('meeting_url_invitation')),
                    start_date=datetime.datetime(
                        year=int(m.group('start_year')),
                        month=int(m.group('start_month')),
                        day=int(m.group('start_day'))),
                    meeting_url_agenda=server_url_replace(m.group('meeting_url_agenda')),
                    end_date=datetime.datetime(
                        year=int(m.group('end_year')),
                        month=int(m.group('end_month')),
                        day=int(m.group('end_day'))),
                    meeting_url_report='',
                    tdoc_start=tdoc.utils.is_generic_tdoc(m.group('tdoc_start')),
                    tdoc_end=tdoc.utils.is_generic_tdoc(m.group('tdoc_end')),
                    meeting_url_docs=server_url_replace(m.group('meeting_url_docs')),
                    meeting_folder_url=server_url_replace(m.group('files_url'))
                )
                for m in meeting_matches if m is not None
            ]
            loaded_meeting_entries.extend(meeting_matches_parsed)
            print(f'Added {len(meeting_matches_parsed)} started but unfinished meetings to group {k}')

            # Meetings for which an invitation is not yet ready
            meeting_matches = meeting_without_invitation_regex.finditer(markup_file_content)
            meeting_matches_parsed = [
                MeetingEntry(
                    meeting_group=m.group('meeting_group'),
                    meeting_number=m.group('meeting_number'),
                    meeting_url_3gu=server_url_replace(m.group('meeting_url_3gu')),
                    meeting_name=m.group('meeting_name'),
                    meeting_location=m.group('meeting_location'),
                    meeting_url_invitation='',
                    start_date=datetime.datetime(
                        year=int(m.group('start_year')),
                        month=int(m.group('start_month')),
                        day=int(m.group('start_day'))),
                    meeting_url_agenda='',
                    end_date=datetime.datetime(
                        year=int(m.group('end_year')),
                        month=int(m.group('end_month')),
                        day=int(m.group('end_day'))),
                    meeting_url_report='',
                    tdoc_start=None,
                    tdoc_end=None,
                    meeting_url_docs='',
                    meeting_folder_url=''
                )
                for m in meeting_matches if
                m is not None and
                m.group('start_year') is not None and
                m.group('end_year') is not None
            ]
            loaded_meeting_entries.extend(meeting_matches_parsed)
            print(f'Added {len(meeting_matches_parsed)} meetings without invitation to group {k}')
        else:
            print(f'Not found: {v}')
    # print(meeting_entries)


def search_meeting_for_tdoc(tdoc_str: str) -> MeetingEntry:
    """
    Searches for a specific TDoc in the loaded meeting list
    Args:
        tdoc_str: A TDoc ID

    Returns: A meeting containing this TDoc. None if none found

    """
    parsed_tdoc = tdoc.utils.is_generic_tdoc(tdoc_str)
    if parsed_tdoc is None:
        return None
    print(f'Searching for group {parsed_tdoc.group}, tdoc {parsed_tdoc.number}')
    group_meetings = [m for m in loaded_meeting_entries if parsed_tdoc.group == m.meeting_group]
    print(f'{len(group_meetings)} Group meetings for group {parsed_tdoc.group}')
    matching_meetings = [m for m in group_meetings if m.tdoc_start is not None and m.tdoc_end is not None and
                         m.tdoc_start.number <= parsed_tdoc.number <= m.tdoc_end.number]

    if len(matching_meetings) > 0:
        matching_meeting = matching_meetings[0]
        print(f'Matching meeting found for TDoc {tdoc_str}: {matching_meeting.meeting_name}')
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
    update_local_html_cache(redownload_if_exists=redownload_if_exists)
    convert_local_cache_to_markdown()
    load_markdown_cache_to_memory()


def search_download_and_open_tdoc(tdoc_str: str) -> Tuple[Any, Any]:
    """
    Searches for a given TDoc. If the zip file contains many files (e.g. typical for plenary CR packs), it will only
    open the folder.
    Args:
        tdoc_str: The TDoc ID

    Returns: The files that were opened

    """
    if tdoc_str is None or tdoc_str == '':
        return None, None

    # Load data if needed
    if len(loaded_meeting_entries) == 0:
        print('Triggering update of local cache')
        update_local_html_cache(redownload_if_exists=False)
        convert_local_cache_to_markdown()
        load_markdown_cache_to_memory()

    tdoc_meeting = search_meeting_for_tdoc(tdoc_str)
    if tdoc_meeting is None:
        return None, None

    tdoc_url = tdoc_meeting.get_tdoc_url(tdoc_str)
    local_folder = os.path.join(tdoc_meeting.get_local_folder_for_meeting(), tdoc_str)
    create_folder_if_needed(local_folder, create_dir=True)
    local_target = os.path.join(local_folder, f'{tdoc_str}.zip')
    print(f'Downloading {tdoc_url} to {local_target}')
    download_file_to_location(tdoc_url, local_target)
    files_in_zip = unzip_files_in_zip_file(local_target)
    if len(files_in_zip) <= maximum_number_of_files_to_open:
        opened_files, metadata_list = parsing.word.pywin32.open_files(files_in_zip, return_metadata=True)
    else:
        print(
            f'More than {maximum_number_of_files_to_open} contained within {tdoc_str}. Opening folder instead of files')
        folder_to_open, first_file = os.path.split(files_in_zip[0])
        os.startfile(folder_to_open)
        opened_files = folder_to_open
        metadata_list = None
    return opened_files, metadata_list
