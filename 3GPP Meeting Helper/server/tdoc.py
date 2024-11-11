import concurrent.futures
import os
import os.path
import re
import traceback
from typing import List, Tuple

import server.common
import tdoc.utils
import tdoc.utils
import utils.local_cache
from application.zip_files import unzip_files_in_zip_file
from server.common import get_remote_meeting_folder, get_inbox_root, ServerType, TdocType, DocumentType, \
    get_document_or_folder_url
from server.connection import get_remote_file
from utils.local_cache import get_cache_folder, get_local_revisions_filename, get_local_drafts_filename, \
    get_meeting_folder

agenda_regex = re.compile(r'.*(?P<type>(Agenda|Session( |%20)Plan)).*[-_]([ ]|%20)*([vr])?(?P<version>\d*).*\..*')
agenda_docx_regex = re.compile(
    r'.*(?P<type>(Agenda|Session Plan)).*[-_]([ ]|%20)*([vr])?(?P<version>\d*).*\.(docx|doc|zip)')
agenda_version_regex = re.compile(r'.*(?P<type>(Agenda|Session Plan)).*[-_]?([ ]|%20)*([vr])(?P<version>\d*).*\..*')
agenda_draft_docx_regex = re.compile(
    r'.*(?P<type>(Agenda|Session Plan)).*[-_]([ ]|%20)*([vr])?(?P<version>\d*).*\.(docx|doc|zip)')
folder_ftp_names_regex = re.compile(r'[\d-]+[ ]+.*[ ]+<DIR>[ ]+(.*[uU][pP][dD][aA][tT][eE].*)')


# tdoc_url = 'https://portal.3gpp.org/ngppapp/DownloadTDoc.aspx?contributionUid=S2-2202451'
# Then, search for javascript: window.location.href='https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_150E_Electronic_2022-04/Docs/S2-2202451.zip';//]]> -> extract


def get_sa2_inbox_current_tdoc(searching_for_a_file=False):
    url = get_inbox_root(searching_for_a_file) + 'CurDoc.htm'
    return get_remote_file(url)


def get_sa2_docs_tdoc_list(meeting_folder, save_file_to=None):
    remote_folder = get_remote_meeting_folder(meeting_folder)
    url = remote_folder + 'Docs'
    returned_html = get_remote_file(url, cached_file_to_return_if_error_or_cache=save_file_to)

    return returned_html


def get_sa2_drafts_tdoc_list(meeting_folder):
    remote_folder = get_remote_meeting_folder(meeting_folder)
    url = remote_folder + 'INBOX/DRAFTS'
    returned_html = get_remote_file(url)

    # In this case, we also need to retrieve all sub-pages
    # TO-DO!!!!

    return returned_html


def get_tdoc(
        meeting_folder_name,
        tdoc_id,
        server_type: server.common.ServerType = server.common.ServerType.PUBLIC
):
    """
    Retrieves a TDoc
    Args:
        server_type: The type of server we are using (public/www.3gpp.org or private/10.10.10.10)
        meeting_folder_name: The folder name as in the 3GPP server
        tdoc_id: A TDoc ID, e.g.: S2-240001

    Returns:

    """
    if server_type == server.common.ServerType.PRIVATE:
        use_private_server = True
    else:
        use_private_server = False

    if '*' in tdoc_id:
        is_draft = True
        tdoc_id = tdoc_id.replace('*', '')
    else:
        is_draft = False

    if not tdoc.utils.is_sa2_tdoc(tdoc_id):
        return None, None

    tdoc_local_filename = get_local_filename_for_tdoc(meeting_folder_name, tdoc_id, is_draft=is_draft)
    zip_file_list = get_remote_filename_candidates_for_tdoc(
        meeting_folder_name=meeting_folder_name,
        tdoc_id=tdoc_id,
        use_private_server=use_private_server,
        is_draft=is_draft)
    zip_file_url = None

    if not os.path.exists(tdoc_local_filename):
        # Try all the candidates until we find a working one (e.g. in /Docs and /Inbox)
        print(f'Downloading from: {zip_file_list}')
        tdoc_file = None

        for zip_file_url in zip_file_list:
            tdoc_file = get_remote_file(zip_file_url, cache=False)
            if tdoc_file is not None:
                break
        if tdoc_file is None:
            # No need to retry. Additional download folders are now implemented outside of this fuction
            return_value = None
            return return_value, zip_file_url
        # Drive zip file to disk
        with open(tdoc_local_filename, 'wb') as output:
            output.write(tdoc_file)

    # If the file does not now exist, there was an error (e.g. not found)
    if not os.path.exists(tdoc_local_filename):
        return None, None

    return unzip_files_in_zip_file(tdoc_local_filename), zip_file_url


def cache_tdocs(tdoc_list, download_from_private_server: bool, meeting_folder_name: str):
    if tdoc_list is None:
        return

    # See https://docs.python.org/3/library/concurrent.futures.html
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        future_to_url = {executor.submit(
            lambda tdoc_to_download_lambda: server.tdoc.get_tdoc(
                meeting_folder_name=meeting_folder_name,
                tdoc_id=tdoc_to_download_lambda,
                server_type=server.common.ServerType.PRIVATE if download_from_private_server else server.common.ServerType.PUBLIC),
            tdoc_to_download_lambda): tdoc_to_download_lambda for tdoc_to_download_lambda in tdoc_list}
        for future in concurrent.futures.as_completed(future_to_url):
            file_to_download = future_to_url[future]
            try:
                retrieved_files, tdoc_url = future.result()
            except Exception as exc:
                print('%r generated an exception: %s' % (file_to_download, exc))


def get_inbox_tdocs_list_cache_local_cache(create_dir=True):
    cache_folder = get_cache_folder(create_dir)
    inbox_cache = os.path.join(cache_folder, 'InboxCache.html')
    return inbox_cache


def get_local_folder_for_tdoc(
        meeting_folder_name,
        tdoc_id,
        create_dir=True,
        email_approval=False,
        is_draft=False):
    meeting_folder = get_meeting_folder(meeting_folder_name)

    year, tdoc_number, revision = tdoc.utils.get_tdoc_year(tdoc_id, include_revision=True)
    if revision is not None:
        # Remove 'rXX' from the name for folder generation if found
        tdoc_id = tdoc_id[:-3]

    if not is_draft:
        folder_name = os.path.join(meeting_folder, tdoc_id)
    else:
        folder_name = os.path.join(meeting_folder, tdoc_id, 'Drafts')
    if email_approval:
        folder_name = os.path.join(folder_name, 'email approval')
    if create_dir and (not os.path.exists(folder_name)):
        try:
            os.makedirs(folder_name, exist_ok=True)
        except FileExistsError:
            print("Could not create directory. File already: {0}".format(folder_name))
    return folder_name


def get_local_filename_for_tdoc(
        meeting_folder_name: str,
        tdoc_id: str,
        create_dir=True,
        is_draft=False):
    if not is_draft:
        # TDoc or revision
        folder_name = get_local_folder_for_tdoc(meeting_folder_name, tdoc_id, create_dir)
    else:
        # Draft! We cannot have a '*' in the path. Replace just in case it was not replaced
        folder_name = get_local_folder_for_tdoc(
            meeting_folder_name,
            tdoc_id.replace('*', ''),
            create_dir,
            is_draft=is_draft)
    return os.path.join(folder_name, tdoc_id + '.zip')


def get_remote_filename_candidates_for_tdoc(
        meeting_folder_name,
        tdoc_id: str,
        use_private_server=False,
        is_draft=False,
        override_folder_path: str = None
) -> List[str]:
    """
    Returns all possible locations of a TDoc
    Args:
        meeting_folder_name: The meeting folder name in the 3GPP server
        tdoc_id: The TDoc ID
        use_private_server: What type of server we are using
        is_draft: Whether this is a draft document
        override_folder_path: Folder override

    Returns: A list of candidate URLs

    """
    # Check if this is a TDoc revision. If yes, change the folder to the revisions' folder.
    year, tdoc_number, revision = tdoc.utils.get_tdoc_year(tdoc_id, include_revision=True)
    server_type = ServerType.PRIVATE if use_private_server else ServerType.PUBLIC
    tdoc_type = TdocType.DRAFT if is_draft else (TdocType.REVISION if revision else TdocType.NORMAL)

    # Instead of using get_remote_meeting_folder() (old function)
    candidate_folders = get_document_or_folder_url(
        server_type=server_type,
        document_type=DocumentType.TDOC,
        meeting_folder_in_server=meeting_folder_name,
        tdoc_type=tdoc_type,
        override_folder_path=override_folder_path
    )

    candidate_urls = [candidate_folder + tdoc_id + '.zip' for candidate_folder in candidate_folders]
    return candidate_urls


def get_remote_filename_for_tdoc(
        meeting_folder_name,
        tdoc_id: str,
        use_private_server=False,
        is_draft=False,
        override_folder_path: str = None
) -> str | None:
    """
    Returns the first possible locations of a TDoc
    Args:
        meeting_folder_name: The meeting folder name in the 3GPP server
        tdoc_id: The TDoc ID
        use_private_server: What type of server we are using
        is_draft: Whether this is a draft document
        override_folder_path: Folder override

    Returns: One possible candidate URL

    """
    candidate_folders = get_remote_filename_candidates_for_tdoc(
        meeting_folder_name=meeting_folder_name,
        tdoc_id=tdoc_id,
        use_private_server=use_private_server,
        is_draft=is_draft,
        override_folder_path=override_folder_path
    )

    if len(candidate_folders) == 0:
        return None

    return candidate_folders[0] + tdoc_id + '.zip'


ai_names_cache = {}


def download_docs_file(meeting) -> str | None:
    """
    Downloads the docs list for a given meeting,
    e.g. https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_156E_Electronic_2023-04/Docs
    Args:
        meeting: The folder name in the 3GPP server, e.g. TSGS2_156E_Electronic_2023-04

    Returns: The file (path) where the HTML was saved to

    """
    try:
        meeting_server_folder = meeting  # e.g. TSGS2_144E_Electronic
        print('Retrieving docs for {0} meeting'.format(meeting))
        local_file = utils.local_cache.get_local_docs_filename(meeting_server_folder)
        html = get_sa2_docs_tdoc_list(meeting_server_folder, save_file_to=local_file)
        if html is None:
            print('Docs file for {0} not found'.format(meeting))
            return None
        utils.local_cache.write_data_and_open_file(html, local_file)
        return local_file
    except Exception as e:
        print(f'Could get not docs file for {meeting}: {e}')
        traceback.print_exc()
        return None


def download_revisions_file(meeting) -> Tuple[str, str] | Tuple[None, None]:
    """
    Downloads the revisions list for a given meeting,
    e.g. https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_156E_Electronic_2023-04/INBOX/Revisions
    Args:
        meeting: The folder name in the 3GPP server, e.g. TSGS2_156E_Electronic_2023-04

    Returns: The file (path) where the HTML was saved to

    """
    try:
        meeting_server_folder = meeting  # e.g. TSGS2_144E_Electronic
        print('Retrieving revisions for {0} meeting'.format(meeting))
        local_file = get_local_revisions_filename(meeting_server_folder)
        folder_candidates_for_revisions = get_document_or_folder_url(
            server_type=ServerType.PUBLIC,
            document_type=DocumentType.TDOC,
            meeting_folder_in_server=meeting,
            tdoc_type=TdocType.REVISION
        )
        html = None
        folder_url = None
        for folder_candidate_for_revisions in folder_candidates_for_revisions:
            html = get_remote_file(
                url=folder_candidate_for_revisions,
                cached_file_to_return_if_error_or_cache=local_file)
            if html is not None:
                folder_url = folder_candidate_for_revisions
                break
        if html is None:
            print('Revisions file for {0} not found'.format(meeting))
            return None, None
        utils.local_cache.write_data_and_open_file(html, local_file)
        print(f'Found revisions file under {folder_url}. Saved to {local_file}')
        return local_file, folder_url
    except Exception as e:
        print(f'Could get not revisions file for {meeting}: {e}')
        traceback.print_exc()
        return None, None


def download_drafts_file(meeting) -> str | None:
    """
    Downloads the drafts list for a given meeting,
    e.g. https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_156E_Electronic_2023-04/INBOX/DRAFTS
    Args:
        meeting: The folder name in the 3GPP server, e.g. TSGS2_156E_Electronic_2023-04

    Returns: The file (path) where the HTML was saved to

    """
    try:
        meeting_server_folder = meeting  # e.g. TSGS2_144E_Electronic
        print('Retrieving drafts for {0} meeting'.format(meeting))
        local_file = get_local_drafts_filename(meeting_server_folder)
        html = get_sa2_drafts_tdoc_list(meeting_server_folder)
        if html is None:
            print('Drafts file for {0} not found'.format(meeting))
            return None
        utils.local_cache.write_data_and_open_file(html, local_file)
        return local_file
    except Exception as e:
        print(f'Could not get drafts file for {meeting}: {e}')
        traceback.print_exc()
        return None
