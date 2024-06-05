from typing import NamedTuple

import server.tdoc


class TdocsByAgendaDownloadResults(NamedTuple):
    tdocs_by_agenda_html_bytes: bytes | None
    revisions_file_path: str | None
    drafts_file_path: str | None


def get_tdocs_by_agenda_for_specific_meeting(
        meeting_folder: str,
        use_private_server: bool,
        get_revisions_file=True,
        get_drafts_file=False,
        open_tdocs_by_agenda_in_browser=False,
) -> TdocsByAgendaDownloadResults:
    """
    Retrieves and saves to local disc the TDocsByAgenda, as well as (optionally) the drafts and inbox folders
    Args:
        meeting_folder: The meetings we refer to (server folder)
        use_private_server: Whether we are in the 3GPP Wi-Fi or not
        get_revisions_file: Whether to return the revisions file
        get_drafts_file: Whether to return the drafts file
        open_tdocs_by_agenda_in_browser: Whether to open the file in the browser

    Returns:

    """
    return_data = server.tdoc.get_tdocs_by_agenda_for_selected_meeting(
        meeting_folder=meeting_folder,
        use_private_server=use_private_server,
        open_tdocs_by_agenda_in_browser=open_tdocs_by_agenda_in_browser)

    # Optional download of revisions
    revisions_file = None
    if get_revisions_file:
        try:
            revisions_file, revisions_folder_url = server.tdoc.download_revisions_file(meeting_folder)
        except Exception as e:
            # Not all meetings have revisions
            print(f'Exception downloading revisions file for {meeting_folder}: {e}')

    # Optional download of drafts
    drafts_file = None
    if get_drafts_file:
        try:
            drafts_file = server.tdoc.download_drafts_file(meeting_folder)
        except Exception as e:
            # Not all meetings have drafts
            print(f'Could not download drafts folder for {meeting_folder}: {e}')

    return TdocsByAgendaDownloadResults(
        tdocs_by_agenda_html_bytes=return_data,
        revisions_file_path=revisions_file,
        drafts_file_path=drafts_file
    )
