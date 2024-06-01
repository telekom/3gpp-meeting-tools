import concurrent.futures
import datetime
import os
import os.path
import re
import traceback
from enum import Enum
from typing import NamedTuple, List, Tuple
from urllib.parse import urljoin

import application.meeting_helper
import application.word
import parsing.html.common
import parsing.html.common as html_parser
import parsing.word.docx
import server.common
import tdoc.utils
import tdoc.utils
import utils.local_cache
from application.zip_files import unzip_files_in_zip_file
from server.common import get_remote_meeting_folder, get_inbox_root, ServerType, TdocType, DocumentType, \
    get_document_or_folder_url
from server.connection import get_remote_file
from utils.local_cache import get_cache_folder, get_local_revisions_filename, get_local_drafts_filename, \
    get_meeting_folder, get_local_agenda_folder

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


def get_sa2_inbox_tdoc_list(
        open_tdocs_by_agenda_in_browser=False,
        use_cached_file_if_available=False):
    url = get_inbox_root(searching_for_a_file=True) + 'TdocsByAgenda.htm'
    print(
        f'Retrieving TdocsByAgenda from Inbox ({url}): open={open_tdocs_by_agenda_in_browser}, use cache={use_cached_file_if_available}')
    if open_tdocs_by_agenda_in_browser:
        os.startfile(url)
    # Return back cached HTML if there is an error retrieving the remote HTML
    fallback_cache = get_inbox_tdocs_list_cache_local_cache()
    online_html = get_remote_file(
        url,
        cached_file_to_return_if_error_or_cache=fallback_cache,
        use_cached_file_if_available=use_cached_file_if_available
    )
    return online_html


def get_sa2_meeting_tdoc_list(meeting_folder, save_file_to=None, open_tdocs_by_agenda_in_browser=False):
    remote_folder = get_remote_meeting_folder(meeting_folder)
    url = remote_folder + 'TdocsByAgenda.htm'
    returned_html = get_remote_file(url, cached_file_to_return_if_error_or_cache=save_file_to)

    if open_tdocs_by_agenda_in_browser:
        os.startfile(url)

    # Normal case
    if returned_html is not None:
        return returned_html

    # In some cases, the original TDocsByAgenda was removed (e.g. 136AH meeting). In this case, we have to look for a substitute
    folder_contents = get_remote_file(remote_folder)
    parsed_folder = html_parser.parse_3gpp_http_ftp(folder_contents)
    tdocs_by_agenda_files = [file for file in parsed_folder.files if
                             ('TdocsByAgenda' in file) and (('.htm' in file) or ('.html' in file))]
    if len(tdocs_by_agenda_files) > 0:
        file_to_get = tdocs_by_agenda_files[0]
        url = remote_folder + file_to_get
        new_html = get_remote_file(url)
        return new_html
    else:
        print('Returned TdocsByAgenda as NONE. Something went wrong when retrieving TDocsByAgenda.htm...')
        return None


def get_sa2_docs_tdoc_list(meeting_folder, save_file_to=None):
    remote_folder = get_remote_meeting_folder(meeting_folder)
    url = remote_folder + 'Docs'
    returned_html = get_remote_file(url, cached_file_to_return_if_error_or_cache=save_file_to)

    return returned_html


def get_sa2_revisions_tdoc_list(meeting_folder, save_file_to=None):
    remote_folder = get_remote_meeting_folder(meeting_folder)
    url = remote_folder + 'INBOX/Revisions'
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
        server_type: server.common.ServerType = server.common.ServerType.PUBLIC,
        return_url=False,
        use_email_approval_inbox=False,
        additional_folders: List[str] | None = None
):
    """
    Retrieves a TDoc
    Args:
        server_type: The type of server we are using (public/www.3gpp.org or private/10.10.10.10)
        additional_folders: A list of additional folder to search in the server, e.g. ['ftp/SA/SA2/Inbox/']
        meeting_folder_name: The folder name as in the 3GPP server
        tdoc_id: A TDoc ID, e.g.: S2-240001
        return_url: The returned URL
        use_email_approval_inbox: Whether to use the email approval inbox

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
        if not return_url:
            return None
        else:
            return None, None

    tdoc_local_filename = get_local_filename(meeting_folder_name, tdoc_id, is_draft=is_draft)
    zip_file_list: List[str] = []
    zip_file_url = get_remote_filename_for_tdoc(
        meeting_folder_name=meeting_folder_name,
        tdoc_id=tdoc_id,
        use_private_server=use_private_server,
        is_draft=is_draft)
    zip_file_list.append(zip_file_url)
    if additional_folders is not None:
        print(f'Searching additional folders in {additional_folders}')
        for additional_folder in additional_folders:
            additional_zip_file_url = get_remote_filename_for_tdoc(
                meeting_folder_name=meeting_folder_name,
                tdoc_id=tdoc_id,
                use_private_server=use_private_server,
                is_draft=is_draft,
                override_folder_path=additional_folder)
            zip_file_list.append(additional_zip_file_url)
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
            if not return_url:
                return return_value
            else:
                return return_value, zip_file_url
        # Drive zip file to disk
        with open(tdoc_local_filename, 'wb') as output:
            output.write(tdoc_file)

    # If the file does not now exist, there was an error (e.g. not found)
    if not os.path.exists(tdoc_local_filename):
        if not return_url:
            return None
        else:
            return None, None

    if not return_url:
        return unzip_files_in_zip_file(tdoc_local_filename)
    else:
        return unzip_files_in_zip_file(tdoc_local_filename), zip_file_url


def cache_tdocs(tdoc_list, download_from_inbox: bool, meeting_folder_name: str):
    if tdoc_list is None:
        return

    # See https://docs.python.org/3/library/concurrent.futures.html
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        future_to_url = {executor.submit(
            lambda tdoc_to_download_lambda: server.tdoc.get_tdoc(
                meeting_folder_name=meeting_folder_name,
                tdoc_id=tdoc_to_download_lambda,
                server_type=server.common.ServerType.PRIVATE if download_from_inbox else server.common.ServerType.PUBLIC,
                return_url=True),
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


def get_private_server_tdocs_by_agenda_local_cache(create_dir=True):
    cache_folder = get_cache_folder(create_dir)
    cache_file = os.path.join(cache_folder, '3gpp_server_TdocsByAgenda.html')
    return cache_file


def get_local_folder(
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


def get_local_filename(meeting_folder_name, tdoc_id, create_dir=True, is_draft=False):
    if not is_draft:
        # TDoc or revision
        folder_name = get_local_folder(meeting_folder_name, tdoc_id, create_dir)
    else:
        # Draft! We cannot have a '*' in the path. Replace just in case it was not replaced
        folder_name = get_local_folder(meeting_folder_name, tdoc_id.replace('*', ''), create_dir, is_draft=is_draft)
    return os.path.join(folder_name, tdoc_id + '.zip')


def get_local_tdocs_by_agenda_filename(meeting_folder_name):
    folder = get_local_agenda_folder(meeting_folder_name, create_dir=True)
    return os.path.join(folder, 'TdocsByAgenda.htm')


def get_remote_filename_for_tdoc(
        meeting_folder_name,
        tdoc_id: str,
        use_private_server=False,
        is_draft=False,
        override_folder_path: str = None
) -> str | None:
    # Check if this is a TDoc revision. If yes, change the folder to the revisions' folder.
    year, tdoc_number, revision = tdoc.utils.get_tdoc_year(tdoc_id, include_revision=True)
    server_type = ServerType.PRIVATE if use_private_server else ServerType.PUBLIC
    tdoc_type = TdocType.DRAFT if is_draft else (TdocType.REVISION if revision else TdocType.NORMAL)

    # Instead of using get_remote_meeting_folder() (old function)
    folder = get_document_or_folder_url(
        server_type=server_type,
        document_type=DocumentType.TDOC,
        meeting_folder_in_server=meeting_folder_name,
        tdoc_type=tdoc_type,
        override_folder_path=override_folder_path
    )

    if len(folder) == 0:
        return None

    return folder[0] + tdoc_id + '.zip'


def get_agenda_files(
        meeting_folder_name,
        server_type: ServerType) -> None:
    """
    Retrieves all the agenda (and session plan) files from the agenda folders (both in drafts and non-drafts)
    Args:
        meeting_folder_name: The server folder name for this meeting
        server_type: Whether we are querying the private (3GPP WiFi) or public (www.3gpp.org) server
    """
    server_folders = get_document_or_folder_url(
        server_type=server_type,
        document_type=DocumentType.AGENDA,
        meeting_folder_in_server=meeting_folder_name)

    print(f'Will download all agenda files in {server_folders}')
    agenda_folders = [retrieved_data
                      for retrieved_data
                      in [(server_folder, get_remote_file(server_folder, cache=False)) for server_folder in server_folders]
                      if retrieved_data[1] is not None]

    agenda_files: List[Tuple[str, str]] = []
    agenda_local_folder = get_local_agenda_folder(meeting_folder_name)
    for agenda_folder in agenda_folders:
        folder_url = agenda_folder[0]
        html = agenda_folder[1]
        parsed_folder = html_parser.parse_3gpp_http_ftp(html)
        agenda_files_in_folder = [file for file in parsed_folder.files if agenda_regex.match(file)]
        agenda_files_url_local_file_in_folder = [
            (urljoin(folder_url, f), os.path.join(agenda_local_folder, f))
            for f in agenda_files_in_folder]
        agenda_files.extend(agenda_files_url_local_file_in_folder)
    print(f'Will download agenda files: ')
    for agenda_file in agenda_files:
        print(f'  {agenda_file[0]}')

    for agenda_file in agenda_files:
        local_file = agenda_file[1]
        server.common.download_file_to_location(
            url=agenda_file[0],
            local_location=local_file,
            cache=True)
        filename, file_extension = os.path.splitext(local_file)

        if file_extension == '.zip':
            unzip_files_in_zip_file(local_file)


class AgendaType(Enum):
    AGENDA = 1
    SESSION_PLAN = 2


def get_latest_agenda_file(agenda_files, agenda_type: AgendaType = AgendaType.AGENDA) -> str | None:
    if agenda_files is None:
        return None

    # Remove case when Word temporary documents are stored
    agenda_files = [file for file in agenda_files if agenda_docx_regex.match(file) and ('~$' not in file)]

    # Filter by Agenda or Session plan
    if agenda_type == AgendaType.AGENDA:
        agenda_files = [agenda_file for agenda_file in agenda_files if "Agenda" in agenda_file]
        search_type = "agenda"
    else:
        agenda_files = [agenda_file for agenda_file in agenda_files if "Session" in agenda_file]
        search_type = "session plan"

    if len(agenda_files) == 0:
        return None

    # Parse file data
    draft_agenda_files = [file for file in agenda_files if agenda_draft_docx_regex.match(file)]
    if (len(draft_agenda_files) > 0) and (len(draft_agenda_files) != len(agenda_files)):
        # Non-draft agendas have priority over draft agenda files
        agenda_files = [agenda_file for agenda_file in agenda_files if agenda_file not in draft_agenda_files]

    last_agenda_or_session_plan = max(agenda_files, key=get_agenda_file_version_number)
    print(f'Most recent {search_type} file seems to be {last_agenda_or_session_plan}')
    return last_agenda_or_session_plan


# Parses the order (including TDoc number) of an Agenda file.
# This new format was introduced in SA2#158
agenda_tdoc_regex = re.compile('(Draft_)?' + tdoc.utils.tdoc_regex_str)


def get_agenda_file_version_number(x):
    if len(x) == 0 or x[0] == '~':
        # Sanity check for empty strings and/or temporary files
        return -1
    tdoc_match = agenda_tdoc_regex.match(x)
    if tdoc_match is not None:
        tdoc_year = float(tdoc_match.groupdict()['year'])
        tdoc_id = float(tdoc_match.groupdict()['tdoc_number'])
        tdoc_number = tdoc_year * 100000 + tdoc_id
    else:
        print('Could not parse TDoc number from agenda file {0}'.format(x))
        tdoc_number = 0
    x_without_dashes = x.replace('-', '')
    agenda_match = agenda_version_regex.match(x_without_dashes)
    try:
        agenda_version = agenda_match.groupdict()['version']
    except:
        print('Could not parse Agenda version number for file {0}'.format(x))
        return -1
    print('{0} is agenda/session plan version {1} for tdoc ID {2:.0f}'.format(x, agenda_version, tdoc_number))

    # Support up to 100 agenda versions. It should be OK...
    agenda_version = tdoc_number + float(agenda_version) / 100
    return agenda_version


ai_names_cache = {}


class AgendaInfo(NamedTuple):
    agenda_path: str | None
    agenda_version_int: int | None
    session_plan_path: str | None
    session_plan_version_int: int | None


def get_last_agenda(meeting_folder):
    agenda_folder = get_local_agenda_folder(meeting_folder)
    agenda_files = [file for file in os.listdir(agenda_folder)]

    last_agenda = get_latest_agenda_file(agenda_files, AgendaType.AGENDA)
    last_session_plan = get_latest_agenda_file(agenda_files, AgendaType.SESSION_PLAN)

    if last_agenda is None:
        return AgendaInfo(
            agenda_path=None,
            agenda_version_int=None,
            session_plan_path=None,
            session_plan_version_int=None)

    agenda_path = os.path.join(agenda_folder, last_agenda)
    try:
        agenda_item_descriptions = parsing.word.docx.import_agenda(agenda_path)
    except:
        agenda_item_descriptions = {}
    ai_names_cache[meeting_folder] = agenda_item_descriptions

    # Convert agenda version
    agenda_version_str = ''
    agenda_version_int = -1
    try:
        last_agenda_match = agenda_docx_regex.match(last_agenda)
        agenda_version_str = last_agenda_match.groupdict()['version']
        agenda_version_int = int(agenda_version_str)
    except ValueError as e:
        print(f"Could not parse agenda version: {agenda_version_str}. Agenda: {last_agenda}")
        traceback.print_exc()

    session_plan_path = None
    session_plan_version_int = None
    if last_session_plan is not None:
        session_plan_path = os.path.join(agenda_folder, last_session_plan)
        session_plan_version_str = ''
        try:
            session_plan_match = agenda_docx_regex.match(last_agenda)
            session_plan_version_str = session_plan_match.groupdict()['version']
            session_plan_version_int = int(agenda_version_str)
        except ValueError as e:
            print(
                f"Could not parse session plan version: {session_plan_version_str}. Session plan: {last_session_plan}")
            traceback.print_exc()

    return AgendaInfo(
        agenda_path=agenda_path,
        agenda_version_int=agenda_version_int,
        session_plan_path=session_plan_path,
        session_plan_version_int=session_plan_version_int)


def get_tdocs_by_agenda_for_selected_meeting(
        meeting_folder: str,
        use_private_server=False,
        save_file_to=None,
        open_tdocs_by_agenda_in_browser=False):
    """
    Returns the HTML of a TdocsByAgenda file for a given meeting
    Args:
        meeting_folder: The meeting folder as named in the 3GPP server
        use_private_server: Whether the private server (10.10.10.10) is to be used
        save_file_to: Where to save the file to
        open_tdocs_by_agenda_in_browser: Whether to open the file in the browser

    Returns: The HTML contents (bytes)
    """
    # If the inbox is active, we need to download both and return the newest one
    html_inbox = None

    datetime_inbox = datetime.datetime.min

    if use_private_server:
        print('Getting TDocs by agenda from inbox')
        html_inbox = get_sa2_inbox_tdoc_list(
            open_tdocs_by_agenda_in_browser=open_tdocs_by_agenda_in_browser,
            use_cached_file_if_available=True
        )
        # Avoid opening the file twice
        open_tdocs_by_agenda_in_browser = False
        datetime_inbox = parsing.html.common.TdocsByAgendaData.get_tdoc_by_agenda_date(html_inbox)

    print('Getting TDocs by agenda from server')
    # print(inspect.stack())
    html_3gpp = get_sa2_meeting_tdoc_list(meeting_folder, save_file_to=save_file_to,
                                          open_tdocs_by_agenda_in_browser=open_tdocs_by_agenda_in_browser)
    datetime_3gpp = parsing.html.common.TdocsByAgendaData.get_tdoc_by_agenda_date(html_3gpp)

    if datetime_3gpp is None:
        datetime_3gpp = datetime.datetime.min

    if datetime_inbox is None:
        datetime_inbox = datetime.datetime.min

    if use_private_server:
        if datetime_3gpp > datetime_inbox:
            html = html_3gpp
            print('3GPP server TDocs by agenda are more recent')
        else:
            html = html_inbox
            print('Inbox TDocs by agenda are more recent')
        print('3GPP server: {0}, Inbox: {1}'.format(str(datetime_3gpp), str(datetime_inbox)))
    else:
        html = html_3gpp
        print('TDocs by agenda are from 3GPP server (not inbox)')
    return html


def download_agenda_file(meeting, inbox_active=False, open_tdocs_by_agenda_in_browser=False):
    try:
        print('Downloading Agenda File')
        meeting_server_folder = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(
            meeting)
        local_file = get_local_tdocs_by_agenda_filename(meeting_server_folder)
        html = get_tdocs_by_agenda_for_selected_meeting(meeting_server_folder, inbox_active,
                                                        open_tdocs_by_agenda_in_browser=open_tdocs_by_agenda_in_browser)
        if html is None:
            print('Agenda file for {0} not found'.format(meeting))
            return None
        utils.local_cache.write_data_and_open_file(html, local_file)
        return local_file
    except Exception as e:
        print(f'Could not download agenda file for {meeting}: {e}')
        traceback.print_exc()
        return None


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


def download_revisions_file(meeting) -> str | None:
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
        html = get_sa2_revisions_tdoc_list(meeting_server_folder, save_file_to=local_file)
        if html is None:
            print('Revisions file for {0} not found'.format(meeting))
            return None
        utils.local_cache.write_data_and_open_file(html, local_file)
        return local_file
    except Exception as e:
        print(f'Could get not revisions file for {meeting}: {e}')
        traceback.print_exc()
        return None


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
