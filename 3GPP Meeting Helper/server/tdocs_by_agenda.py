from typing import NamedTuple

import server.connection
import server.tdoc
import utils.local_cache
from application.os import startfile
from server.common.server_utils import get_inbox_root, get_document_or_folder_url
from server.common.server_utils import ServerType, DocumentType
from server.connection import get_remote_file
from server.tdoc import get_inbox_tdocs_list_cache_local_cache


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
    return_data = get_tdocs_by_agenda_for_a_given_meeting(
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


def get_sa2_inbox_tdoc_list(
        open_tdocs_by_agenda_in_browser=False,
        use_cached_file_if_available=False):
    url = get_inbox_root(searching_for_a_file=True) + 'TdocsByAgenda.htm'
    print(
        f'Retrieving TdocsByAgenda from Inbox ({url}): '
        f'open={open_tdocs_by_agenda_in_browser}, '
        f'use cache={use_cached_file_if_available}')
    if open_tdocs_by_agenda_in_browser:
        startfile(url)
    # Return back cached HTML if there is an error retrieving the remote HTML
    fallback_cache = get_inbox_tdocs_list_cache_local_cache()
    online_html = get_remote_file(
        url,
        cached_file_to_return_if_error_or_cache=fallback_cache,
        use_cached_file_if_available=use_cached_file_if_available
    )
    return online_html


def get_tdocs_by_agenda_for_a_given_meeting(
        meeting_folder: str,
        use_private_server=False,
        open_tdocs_by_agenda_in_browser=False) -> bytes | None:
    """
    Returns the HTML of a TdocsByAgenda file for a given meeting
    Args:
        meeting_folder: The meeting folder as named in the 3GPP server
        use_private_server: Whether the private server (10.10.10.10) is to be used
        open_tdocs_by_agenda_in_browser: Whether to open the file in the browser

    Returns: The HTML contents (bytes) or None if it could not be retrieved
    """
    print(f'Retrieving TDocsByAgenda for meeting {meeting_folder}')
    tdocs_by_agenda_server_folder = get_document_or_folder_url(
        server_type=ServerType.PRIVATE if use_private_server else ServerType.PUBLIC,
        document_type=DocumentType.TDOCS_BY_AGENDA,
        meeting_folder_in_server=meeting_folder,
        tdoc_type=None)
    if len(tdocs_by_agenda_server_folder) == 0:
        print(f'Could not retrieve TDocs by Agenda for meeting {meeting_folder}. No target folders for URL retrieval')
        return
    target_url = tdocs_by_agenda_server_folder[0] + 'TdocsByAgenda.htm'
    local_file = utils.local_cache.get_tdocs_by_agenda_filename(meeting_folder_name=meeting_folder)
    tdocs_by_agenda_html = server.connection.get_remote_file(
        target_url,
        cache=True,
        cached_file_to_return_if_error_or_cache=local_file)
    if open_tdocs_by_agenda_in_browser:
        print(f'Opening local TDocsByAgenda file {local_file}')
        startfile(local_file)

    return tdocs_by_agenda_html
