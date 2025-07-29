from typing import List

from  server.common.MeetingEntry import MeetingEntry
from server.common.server_utils import download_file_to_location, FileToDownload, batch_download_file_to_location


def download_meeting_tdocs_excel(meeting:MeetingEntry, redownload_if_exists=False):
    if meeting.tdoc_excel_exists_in_local_folder and not redownload_if_exists:
        return True

    url_to_download = meeting.meeting_tdoc_list_excel_url
    local_path = meeting.tdoc_excel_local_path
    print(f'Downloading TDoc list for {meeting.meeting_name} from {url_to_download} to {local_path}')
    return download_file_to_location(url_to_download, local_path, force_download=True)

def batch_download_meeting_tdocs_excel(meetings:List[MeetingEntry], redownload_if_exists=False):
    if not redownload_if_exists:
        meetings = [m for m in meetings if not m.tdoc_excel_exists_in_local_folder]

    if len(meetings)==0:
        print(f'No meeting TDoc Excel files to download')
        return

    meeting_names_str = ', '.join([m.meeting_name for m in meetings])
    print(f'Will download {len(meetings)} meeting TDoc Excel files: {meeting_names_str}')

    files_to_download = [ FileToDownload(remote_url=m.meeting_tdoc_list_excel_url,
                                         local_filepath=m.tdoc_excel_local_path,
                                         force_download=redownload_if_exists) for m in meetings]
    batch_download_file_to_location(files_to_download, cache = False)