from typing import List, Any

import html2text
import pandas as pd
import re
import traceback
from datetime import datetime as datetime

def extract_chairnotes_from_html(html_content: str):
    """
    Reads the Chair_Notes HTML page and extracts the parsed names of the Chairman's Notes
    Args:
        html_content: The HTML content
    Returns:

    """
    h = html2text.HTML2Text()
    # Ignore converting links from HTML
    h.ignore_links = True
    chairnotes_list_text = h.handle(html_content)
    chairnotes_list = re.finditer(r'ChairNotes_(?P<chair>.*)_(?P<month>[\d]+)-(?P<day>[\d]+)-(?P<time_hour>[\d]{2})(?P<time_minute>[\d]{2}).doc', chairnotes_list_text)
    chairnotes_list: List[re.Match] = list(set(chairnotes_list))

    year: str = re.search(r'([\d]{4})/([\d]{2})/([\d]{2}) ([\d]{2}):([\d]{2})', chairnotes_list_text).group(1)

    return chairnotes_list, year


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

        chairnotes_list, year = extract_chairnotes_from_html(chairnotes_html)

        files = [(e.groupdict(), e.group(0)) for e in chairnotes_list if e[0] is not None]
        files = [{'chair': e[0]['chair'],
                  'datetime': datetime(year=int(year), month=int(e[0]['month']), day=int(e[0]['day']),
                                       hour=int(e[0]['time_hour']), minute=int(e[0]['time_minute'])), 'file': e[1]} for
                 e in files]
        files_df = pd.DataFrame.from_dict(files)
        most_current_notes_per_chair = files_df.sort_values('datetime').groupby('chair').tail(1)
        return most_current_notes_per_chair
    except:
        print("Could not parse Chairman's Notes from {0}".format(chairnotes_file))
        traceback.print_exc()
        return None