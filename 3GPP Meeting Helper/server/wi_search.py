import os.path
import re
from typing import NamedTuple, List

from server.common import download_file_to_location
from server.connection import get_html
from utils.local_cache import file_exists, convert_html_file_to_markup, \
    get_work_items_cache_folder

sid_page = 'https://www.3gpp.org/dynareport?code=WI-List.htm'
wgs_list = [ 'SP', 'S1', 'S2', 'S3', 'S4', 'S5', 'S6', 'CP', 'C1', 'C3', 'C6', 'RP', 'R1', 'R2', 'R3', 'R4', 'R5' ]
local_cache_folder = get_work_items_cache_folder()
local_cache_file = os.path.join(local_cache_folder, 'wi_list.htm')
local_cache_file_md = os.path.join(local_cache_folder, 'wi_list.md')

wi_parse_regex = re.compile(
    r'(?P<uid>[\d]+)\|(?P<code>[\w\-_\.]+)\|(?P<title>[\.\w \d\-‑;\&\(\)/:,<>=\+]+)\|(?P<release>[\w \d‑-]+)\|(?P<lead_body>[\w \d‑\-,]+)')


class WiEntry(NamedTuple):
    uid: str
    code: str
    title: str
    release: str
    lead_body: str

    @property
    def cr_list_url(self) -> str:
        cr_list_url = f'https://portal.3gpp.org/ChangeRequests.aspx?q=1&specnumber=&release=all&workitem={self.uid}'
        return cr_list_url

    @property
    def spec_list_url(self) -> str:
        spec_list_url = f'https://portal.3gpp.org/Specifications.aspx?q=1&WiUid={self.uid}'
        return spec_list_url

    @property
    def wid_page_url(self) -> str:
        sid_page_url = f'https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId={self.uid}'
        return sid_page_url

    def retrieve_last_wid_tdoc_id_from_server(self):
        url_to_download = self.wid_page_url
        wi_html = file = get_html(url_to_download, cache=False)


loaded_wi_entries: List[WiEntry] = []


def download_wi_list(re_download_if_exists=False):
    if re_download_if_exists or not file_exists(local_cache_file):
        download_file_to_location(sid_page, local_cache_file)

    if re_download_if_exists or not file_exists(local_cache_file_md):
        convert_html_file_to_markup(
            local_cache_file,
            output_path=local_cache_file_md,
            ignore_links=True,
            filter_text_function=filter_markdown_text
        )


def filter_markdown_text(markdown_text: str) -> str:
    markdown_text = markdown_text.replace(' | ', '|').replace('| ', '|')
    markdown_text = re.sub(r"\.[\.]{2,}[ ]", '', markdown_text, flags=re.M)
    return markdown_text


def load_wi_entries(re_download_if_exists=False):
    global loaded_wi_entries

    if not file_exists(local_cache_file_md) or re_download_if_exists:
        download_wi_list(re_download_if_exists=re_download_if_exists)

    with open(local_cache_file_md, 'r', encoding='utf-8') as file:
        markup_file_content = file.read()

    wi_matches = wi_parse_regex.finditer(markup_file_content)
    wi_entries = [
        WiEntry(
            uid=m.group('uid').strip(),
            code=m.group('code').strip(),
            title=m.group('title').strip(),
            release=m.group('release').strip(),
            lead_body=m.group('lead_body').strip()
        )
        for m in wi_matches if m is not None
    ]
    loaded_wi_entries = wi_entries
    print(f'Added {len(wi_entries)} Work Item entries')
