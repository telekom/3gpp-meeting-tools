import os.path
import re
from typing import List

from server import tdoc_search
from server.common.server_utils import download_file_to_location, WiEntry
from utils.local_cache import file_exists, convert_html_file_to_markup, \
    get_work_items_cache_folder

sid_page = 'https://www.3gpp.org/dynareport?code=WI-List.htm'
wgs_list = ['SP', 'S1', 'S2', 'S3', 'S4', 'S5', 'S6', 'CP', 'C1', 'C3', 'C6', 'RP', 'R1', 'R2', 'R3', 'R4', 'R5']

initialized = False
local_cache_folder = ''
local_cache_file = ''
local_cache_file_md = ''


def initialize():
    global initialized, local_cache_folder, local_cache_file, local_cache_file_md
    local_cache_folder = get_work_items_cache_folder()
    local_cache_file = os.path.join(local_cache_folder, 'wi_list.htm')
    local_cache_file_md = os.path.join(local_cache_folder, 'wi_list.md')
    initialized = True


wi_parse_regex = re.compile(
    r'(?P<uid>\d{4,}) *\| *(?P<code>[\w, -_]+) *\| *(?P<name>[\d\w, -()-‑–™”]+) *\| *(?P<release>[\w\d, -‑]+) *\| *(?P<groups>[\w, ]+)')

loaded_wi_entries: List[WiEntry] = []

# Update the meeting list such that we can use tdoc_search.loaded_meeting_entries
tdoc_search.fully_update_cache(redownload_if_exists=False)

def download_wi_list(re_download_if_exists=False):
    if not initialized:
        initialize()
    if re_download_if_exists or not file_exists(local_cache_file):
        download_file_to_location(
            sid_page,
            local_cache_file,
            force_download=True
        )

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
    if not initialized:
        initialize()
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
            title=m.group('name').strip(),
            release=m.group('release').strip(),
            lead_body=m.group('groups').strip()
        )
        for m in wi_matches if m is not None
    ]
    loaded_wi_entries = wi_entries
    print(f'Added {len(wi_entries)} Work Item entries')
