import re
from collections import defaultdict
from typing import List, NamedTuple

from parsing.html.common import parse_3gpp_http_ftp_v2


class ChairNotesFile(NamedTuple):
    file: str
    authors: List[str]
    is_combined: bool


def get_latest_chairnotes_files(file_list: List[str]) -> List[ChairNotesFile]:
    pattern = re.compile(r"ChairNotes_(.*?)_(\d{2}-\d{2}-\d{4})")
    combined_pattern = re.compile(r"Combined_ChairNotes_(\d{2}-\d{2}-\d{4})")

    file_dict = defaultdict(list)
    combined_latest = ("", "")

    for file in file_list:
        if match := combined_pattern.search(file):
            timestamp = match.group(1).replace('-', '')
            combined_latest = max(combined_latest, (timestamp, file))
        elif match := pattern.search(file):
            author, timestamp = match.groups()
            timestamp = timestamp.replace('-', '')
            file_dict[author] = max(file_dict.get(author, ("", [])), (timestamp, [file]))

    result_dict = defaultdict(set)
    combined_timestamp = combined_latest[0]
    combined_file = combined_latest[1]

    for author, (timestamp, files) in file_dict.items():
        if combined_file and combined_timestamp > timestamp:
            result_dict[combined_file].add(author)
        else:
            result_dict[files[0]].add(author)

    if combined_file and combined_file not in result_dict:
        result_dict[combined_file].add("All")

    result = [ChairNotesFile(file=file, authors=list(authors), is_combined=file == combined_file) for file, authors in
              result_dict.items()]

    return result


def chairnotes_file_to_dataframe(chairnotes_file: str):
    """

    Args:
        chairnotes_file: Path of the HTML file listing the contents of the Chair_Notes folder,
         e.g. https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_153E_Electronic_2022-10/INBOX/Chair_Notes

    Returns:
        DataFrame: Containing three columns: chair (Chair's name), datetime, file with the last file of each chair.
    """
    try:
        with open(chairnotes_file, 'r') as file:
            chairnotes_html = file.read()

        folder_list = parse_3gpp_http_ftp_v2(chairnotes_html)
        latest_chairnotes = get_latest_chairnotes_files(folder_list.files)
        return latest_chairnotes
    except Exception as e:
        print(f"Could not parse Chairman's Notes from {chairnotes_file}: {e}")
        return None
