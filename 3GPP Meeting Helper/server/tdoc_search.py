import datetime
import os.path
import re
from typing import NamedTuple, List

from server.common import download_file_to_location
from utils.local_cache import get_meeting_list_folder, convert_html_file_to_markup

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
    r'\[(?P<meeting_group>[\d\w]+)\-(?P<meeting_number>[\d\-\w ]+)\]\((?P<meeting_url_3gu>[^ ]+)\)[ ]?\|[ ]?(?P<meeting_name>[^ ]+)[ ]?\|[ ]?\[(?P<meeting_location>[^\]]+)\]\((?P<meeting_url_invitation>[^ ]+)\)[ ]?\|[ ]?\[(?P<start_year>[\d]+)\-(?P<start_month>[\d]+)\-(?P<start_day>[\d]+)\]\((?P<meeting_url_agenda>[^ ]+)\)\|[ ]?\[(?P<end_year>[\d]+)\-(?P<end_month>[\d]+)\-(?P<end_day>[\d]+)\]\((?P<meeting_url_report>[^ ]+)\)\|[ ]?\[(?P<tdoc_start>[\w\-\d]+)[ -]+(?P<tdoc_end>[\w\-\d]+)\]\((?P<meeting_url_docs>[^ ]+)\)')


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
    tdoc_start: int
    tdoc_end: int
    meeting_url_docs: str


def update_local_cache():
    """
    Download the meeting files to the cache
    """
    for k, v in meeting_pages_per_group.items():
        download_file_to_location(v, html_cache[k])


def convert_local_cache_to_markdown():
    """
        Convert local cache to markdown
    """
    str_replace_list = [
        ("""
| """, '| '),
        ("""|
[Register](""", '| [Register]('),
        ("""|
[Participants](""", '| [Participants]('),
        ("""
| [Participants""", '| [Participants]('),
        ("""|
[ICS](""", '| [ICS]('),
        ("""|
\-  |""", '| \- |'),
        ("""‑""", '-'),
        ("""|
[""", '| ['),
        (""")
[  
full document
list]""", ')[full document list]'),
        ("""|
-""", '| - '),
        ("""-
""", '- '),
        ("""|
""", '| '),
        ("""
CANCELLED""", ' CANCELLED'),
        ("""[  
full document
list]""", '[full document list]'),
        ("""Jeju
""", 'Jeju '),
        (""",
""", ', '),
        ("""New
""", '''New '''),
        ("""Puerto
""", 'Puerto '),
        ("""Beach
""", 'Beach '),
        ("""
| [""", """| [""")
    ]
    for k, v in html_cache.items():
        if os.path.exists(v):
            convert_html_file_to_markup(
                v,
                output_path=markup_cache[k],
                ignore_links=False,
                str_replace_list=str_replace_list)


def load_markdown_cache_to_memory() -> List[MeetingEntry]:
    """
    Parses the markdown cache files and returns the parsed 3GPP meeting list.
    Returns: 3GPP meeting list

    """
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
            return None
        return meeting_url.replace(r'\\', '/').replace('/../../..//ftp/', 'https://www.3gpp.org/ftp/').replace('//', '/')

    for k, v in markup_cache.items():
        if os.path.exists(v):
            with open(v, 'r', encoding='utf-8') as file:
                markup_file_content = file.read()
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
                    tdoc_start=m.group('tdoc_start'),
                    tdoc_end=m.group('tdoc_end'),
                    meeting_url_docs=server_url_replace(m.group('meeting_url_docs')))
                for m in meeting_matches
            ]
            meeting_entries.extend(meeting_matches_parsed)
    print(meeting_entries)
    return meeting_entries
