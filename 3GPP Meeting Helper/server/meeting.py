from  server.common.MeetingEntry import MeetingEntry
from server.common.server_utils import download_file_to_location


def download_meeting_tdocs_excel(meeting:MeetingEntry, redownload_if_exists=False):
    if meeting.tdoc_excel_exists_in_local_folder and not redownload_if_exists:
        return True

    url_to_download = meeting.meeting_tdoc_list_excel_url
    local_path = meeting.tdoc_excel_local_path
    print(f'Downloading TDoc list for {meeting.meeting_name} from {url_to_download} to {local_path}')
    return download_file_to_location(url_to_download, local_path, force_download=True)