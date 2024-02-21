import datetime
import os.path
import re
from typing import NamedTuple, List

import tdoc.utils
from server.common import download_file_to_location
import utils
from utils.local_cache import get_meeting_list_folder, convert_html_file_to_markup, get_cache_folder

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
html_cache = {k: os.path.join(local_cache_folder, k + '.htm') for k, v in meeting_pages_per_group.items()}
markup_cache = {k: os.path.join(local_cache_folder, k + '.md') for k, v in meeting_pages_per_group.items()}
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
    tdoc_start: tdoc.utils.GenericTdoc
    tdoc_end: tdoc.utils.GenericTdoc
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
meeting_entries: List[MeetingEntry] = []


def update_local_cache(redownload_if_exists=False):
    """
    Download the meeting files to the cache

    Args:
        redownload_if_exists: Whether to force a download of the file(s) if they exist
    """
    for k, v in meeting_pages_per_group.items():
        local_file = html_cache[k]
        if redownload_if_exists or not os.path.exists(local_file):
            download_file_to_location(v, local_file)


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
 |""", ") |").replace(""")
|""", ') |').replace('  ', ' ').replace(""")
[ full""", ')[ full]').replace("""| 
[Sophia""", '| [Sophia')
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


def load_markdown_cache_to_memory() -> List[MeetingEntry]:
    """
    Parses the markdown cache files and returns the parsed 3GPP meeting list.
    Returns: 3GPP meeting list

    """
    global meeting_entries
    meeting_entries = []

    def server_url_replace(meeting_url: str) -> str:
        """
        Cleans up the URL and returns an absolute URL pointing to the 3GPP HTTP(s) file server
        Args:
            meeting_url: A relative URL as parsed form the markdown cache,
            e.g. /../../../\\ftp\\tsg_sa\\TSG_SA\\TSGS_60\\Invitation/

        Returns: An absolute URL

        """
        if meeting_url is None:
            return meeting_url
        return meeting_url.replace(r'\\', '/').replace('/../../..//ftp/', 'https://www.3gpp.org/ftp/').replace('//',
                                                                                                               '/')

    for k, v in markup_cache.items():
        if os.path.exists(v):
            print(f'Loading meetings for group {k}')
            with open(v, 'r', encoding='utf-8') as file:
                markup_file_content = file.read()
            meeting_matches = meeting_regex.finditer(markup_file_content)
            if meeting_matches is None:
                return meeting_matches
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
            meeting_entries.extend(meeting_matches_parsed)
    # print(meeting_entries)


def search_meeting_for_tdoc(tdoc_str: str) -> List[MeetingEntry]:
    """
    Searches for a specific TDoc in the loaded meeting list
    Args:
        tdoc_str: A TDoc ID

    Returns: A list of meetings (ideally containing only one element) containing this TDoc

    """
    parsed_tdoc = tdoc.utils.is_generic_tdoc(tdoc_str)
    if parsed_tdoc is None:
        return []
    print(f'Searching for group {parsed_tdoc.group}, tdoc {parsed_tdoc.number}')
    group_meetings = [m for m in meeting_entries if parsed_tdoc.group == m.meeting_group]
    print(f'{len(group_meetings)} Group meetings for group {parsed_tdoc.group}')
    matching_meetings = [m for m in group_meetings if m.tdoc_start is not None and m.tdoc_end is not None and
                         m.tdoc_start.number <= parsed_tdoc.number <= m.tdoc_end.number]
    print(
        f'{len(matching_meetings)} matching meeting(s) found: {", ".join([m.meeting_name for m in matching_meetings])}')
    return matching_meetings
