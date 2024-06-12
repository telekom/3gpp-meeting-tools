import concurrent.futures
import os.path
import re
import socket
import traceback
from enum import Enum
from typing import NamedTuple, List

from server.connection import get_remote_file
from utils.local_cache import get_sa2_root_folder_local_cache, create_folder_if_needed

"""Retrieves data from the 3GPP web server"""
default_http_proxy = 'http://lanbctest:8080'
private_server = '10.10.10.10'
public_server = 'www.3gpp.org'
wg_folder_public_server = 'ftp/tsg_sa/WG2_Arch/'
wg_folder_private_server = 'ftp/SA/SA2/'

sync_folder = 'ftp/Meetings_3GPP_SYNC/SA2/'

host_public_server = 'https://' + public_server
host_private_server = 'http://' + private_server
sa2_url = host_public_server + '/' + wg_folder_public_server
sa2_url_sync = host_public_server + '/' + sync_folder
sa2_url_private_server = host_private_server + '/' + wg_folder_private_server

tdocs_by_agenda_for_checking_meeting_number_in_meeting = 'http://10.10.10.10/ftp/SA/SA2/TdocsByAgenda.htm'


class ServerType(Enum):
    PUBLIC = 1
    PRIVATE = 2


class DocumentType(Enum):
    TDOCS_BY_AGENDA = 1
    AGENDA = 2
    TDOC = 3
    MEETING_ROOT = 4
    CHAIR_NOTES = 5


class TdocType(Enum):
    NORMAL = 1
    REVISION = 2
    DRAFT = 3


class WorkingGroup(Enum):
    SP = 1
    S1 = 2
    S2 = 3
    S3 = 4
    S3LI = 5
    S4 = 6
    S5 = 7
    S6 = 8
    CP = 9
    C1 = 10
    C3 = 11
    C4 = 12
    C6 = 13
    RP = 14
    R1 = 15
    R2 = 16
    R3 = 17
    R4 = 18
    R5 = 19

    def get_wg_folder_name(self, server_type: ServerType) -> str:
        match server_type:
            case ServerType.PRIVATE:
                prefix = 'ftp'
                match self:
                    case WorkingGroup.SP:
                        return f'{prefix}/SA'
                    case WorkingGroup.S1:
                        return f'{prefix}/SA/SA1'
                    case WorkingGroup.S2:
                        return f'{prefix}/SA/SA2'
                    case WorkingGroup.S3:
                        return f'{prefix}/SA/SA3'
                    case WorkingGroup.S3LI:
                        return f'{prefix}/SA/SA3LI'
                    case WorkingGroup.S4:
                        return f'{prefix}/SA/SA4'
                    case WorkingGroup.S5:
                        return f'{prefix}/SA/SA5'
                    case WorkingGroup.S6:
                        return f'{prefix}/SA/SA6'
                    case WorkingGroup.CP:
                        return f'{prefix}/CT'
                    case WorkingGroup.C1:
                        return f'{prefix}/CT1'
                    case WorkingGroup.C3:
                        return f'{prefix}/CT3'
                    case WorkingGroup.C4:
                        return f'{prefix}/CT4'
                    case WorkingGroup.C6:
                        return f'{prefix}/CT6'
                    case WorkingGroup.RP:
                        return f'{prefix}/RAN'
                    case WorkingGroup.R1:
                        return f'{prefix}/RAN1'
                    case WorkingGroup.R2:
                        return f'{prefix}/RAN2'
                    case WorkingGroup.R3:
                        return f'{prefix}/RAN3'
                    case WorkingGroup.R4:
                        return f'{prefix}/RAN4'
                    case WorkingGroup.R5:
                        return f'{prefix}/RAN5'
            case _:
                prefix = 'ftp'
                match self:
                    case WorkingGroup.SP:
                        return f'{prefix}/tsg_sa/TSG_SA'
                    case WorkingGroup.S1:
                        return f'{prefix}/tsg_sa/WG1_Serv'
                    case WorkingGroup.S2:
                        return f'{prefix}/tsg_sa/WG2_Arch'
                    case WorkingGroup.S3:
                        return f'{prefix}/tsg_sa/WG3_Security'
                    case WorkingGroup.S3LI:
                        return f'{prefix}/tsg_sa/WG3_Security/TSGS3_LI'
                    case WorkingGroup.S4:
                        return f'{prefix}/tsg_sa/WG4_CODEC'
                    case WorkingGroup.S5:
                        return f'{prefix}/tsg_sa/WG5_TM'
                    case WorkingGroup.S6:
                        return f'{prefix}/tsg_sa/WG6_MissionCritical'
                    case WorkingGroup.CP:
                        return f'{prefix}/tsg_ct/TSG_CT'
                    case WorkingGroup.C1:
                        return f'{prefix}/tsg_ct/WG1_mm-cc-sm_ex-CN1'
                    case WorkingGroup.C3:
                        return f'{prefix}/tsg_ct/WG3_interworking_ex-CN3'
                    case WorkingGroup.C4:
                        return f'{prefix}/tsg_ct/WG4_protocollars_ex-CN4'
                    case WorkingGroup.C6:
                        return f'{prefix}/tsg_ct/WG6_Smartcard_Ex-T3'
                    case WorkingGroup.RP:
                        return f'{prefix}/tsg_ran/TSG_RAN'
                    case WorkingGroup.R1:
                        return f'{prefix}/tsg_ran/WG1_RL1'
                    case WorkingGroup.R2:
                        return f'{prefix}/tsg_ran/WG2_RL2'
                    case WorkingGroup.R3:
                        return f'{prefix}/tsg_ran/WG3_Iu'
                    case WorkingGroup.R4:
                        return f'{prefix}/tsg_ran/WG4_Radio'
                    case WorkingGroup.R5:
                        return f'{prefix}/tsg_ran/WG5_Test_ex-T1'


def get_document_or_folder_url(
        server_type: ServerType,
        document_type: DocumentType,
        meeting_folder_in_server: str,
        tdoc_type: TdocType | None = None,
        override_folder_path: str | None = None,
        working_group: WorkingGroup = WorkingGroup.S2
) -> List[str]:
    """
    Returns a list of all the places a target file of the specified type could be located in
    Args:
        working_group: Optional parameter to specify the WG of the folder (needed for generating some paths)
        override_folder_path: If this parameter is included, it constructs a folder path for the selected server type
        meeting_folder_in_server: Used for public servers to generate the full URL (not really used for private server)
        server_type: Whether we want the address for the internal 3GPP WiFi (F2F) or public server
        document_type: Whether we are searching for a TDoc, TDocsByAgenda or Agenda file
        tdoc_type: Type of Tdoc (normal or revision). If not included, assumed normal

    Returns:

    """
    # To Do: add WG type
    match server_type:
        case ServerType.PRIVATE:
            host_address = 'http://' + private_server + '/'
        case _:
            host_address = 'https://' + public_server + '/'

    # Skip the rest if we are overriding the path
    if override_folder_path is not None:
        return [f'{host_address}{override_folder_path}']

    root_folder = working_group.get_wg_folder_name(server_type)
    match document_type:
        case DocumentType.CHAIR_NOTES:
            folders = [
                f'{root_folder}/INBOX/Chair_Notes'
            ] if server_type == ServerType.PRIVATE \
                else [
                f'{root_folder}/{meeting_folder_in_server}/INBOX/Chair_Notes'
            ]
        case DocumentType.TDOCS_BY_AGENDA | DocumentType.MEETING_ROOT:
            folders = [
                f'{root_folder}/'
            ] if server_type == ServerType.PRIVATE \
                else [
                f'{root_folder}/{meeting_folder_in_server}/'
            ]
        case DocumentType.AGENDA:
            folders = [
                f'{root_folder}/Agenda/',
                f'{root_folder}/INBOX/DRAFTS/_Session_Plan_Updates/'
            ] if server_type == ServerType.PRIVATE \
                else [
                f'{root_folder}/{meeting_folder_in_server}/Agenda/',
                f'{root_folder}/{meeting_folder_in_server}/INBOX/DRAFTS/_Session_Plan_Updates/'
            ]
        case _:
            # A TDoc
            match tdoc_type:
                case None | TdocType.NORMAL:
                    # Normal TDoc
                    folders = [f'{root_folder}/Docs/'] if server_type == ServerType.PRIVATE \
                        else [f'{root_folder}/{meeting_folder_in_server}/Docs/']
                case TdocType.DRAFT:
                    # Draft TDoc (sub-folders not included!)
                    folders = [f'{root_folder}/INBOX/DRAFTS/'] if server_type == ServerType.PRIVATE \
                        else [f'{root_folder}/{meeting_folder_in_server}/INBOX/DRAFTS/']
                case _:
                    # Revision
                    # No revisions in F2F meetings (at least during the F2F phase)
                    folders = [] if server_type == ServerType.PRIVATE \
                        else [
                        f'{root_folder}/{meeting_folder_in_server}/INBOX/Revisions/',
                        f'{root_folder}/{meeting_folder_in_server}/INBOX/e-mail_Approval/Revisions/']
    target_folders = [host_address + folder for folder in folders]

    print(f'Target folder for meeting '
          f'{meeting_folder_in_server}, '
          f'{server_type}, '
          f'{document_type}, '
          f'{tdoc_type}, '
          f'override {override_folder_path}: {target_folders}')
    return target_folders


def decode_string(str_to_decode: bytes, log_name, print_error=False) -> str | bytes:
    """
    Decodes an HTML (binary) input using different encodings
    Args:
        str_to_decode: The input to decode
        log_name: The name of the file (for logging)
        print_error: Whether to print decodign errors

    Returns:
        The decoded text (string), a bytes object if it could not be decoded
    """
    encodings_to_try = [
        'utf-8',
        'cp1252'
    ]
    for encoding_to_try in encodings_to_try:
        try:
            print(f"Trying to decode {log_name} using {encoding_to_try}")
            decoded_html_data = str_to_decode.decode(encoding=encoding_to_try)
            print(f"Successfully decoded {log_name} using {encoding_to_try}")
            return decoded_html_data
        except:
            if print_error:
                traceback.print_exc()

    print(f"Could not decode {log_name}. Returning HTML as-is")
    return str_to_decode


def get_remote_meeting_folder(
        meeting_folder_name,
        use_private_server=False,
        use_inbox=False,
        override_folder_path: str = None
):
    if override_folder_path is not None:
        if use_private_server:
            folder = host_private_server + '/' + override_folder_path
        else:
            folder = host_public_server + '/' + override_folder_path
    elif use_private_server:
        # e.g., http://10.10.10.10/ftp/SA/SA2/Docs/S2-2405873.zip
        folder = sa2_url_private_server
    else:
        url_prefix = sa2_url
        folder = url_prefix + meeting_folder_name + '/'
    if use_inbox and (override_folder_path is not None):
        folder = folder + 'Inbox/'
    return folder


def get_inbox_root(searching_for_a_file=False):
    if not we_are_in_meeting_network():
        folder = sa2_url_sync
    else:
        folder = sa2_url_private_server
    return folder


def we_are_in_meeting_network():
    # Before, 10.10.10.10 used only FTP, so we had to differentiate between files and folders. Now,
    # we can always just use HTTP (albeit no HTTPs in 10.10.10.10)
    ip_addresses = [i[4][0] for i in socket.getaddrinfo(socket.gethostname(), None)]
    matches = [re.match(r'10.10.(\d)+.(\d)+', ip_address) for ip_address in ip_addresses]
    matches = [match for match in matches if match is not None]
    ip_is_meeting_ip = (len(matches) != 0)
    return ip_is_meeting_ip


def get_sa2_folder(force_redownload=False):
    html = get_remote_file(
        sa2_url,
        cached_file_to_return_if_error_or_cache=get_sa2_root_folder_local_cache(),
        use_cached_file_if_available=not force_redownload
    )
    return html


def download_file_to_location(
        url: str,
        local_location: str,
        cache=False) -> bool:
    """
    Downloads a given file to a local location
    Args:
        cache: Whether to use HTTP caching
        url: The URL to download
        local_location: Where to download the file to

    Returns:
        bool: Whether the file could be successfully downloaded
    """
    try:
        file = get_remote_file(
            url,
            cache=cache,
            use_cached_file_if_available=True,
            cached_file_to_return_if_error_or_cache=local_location)
        if file is None:
            print(f'No file saved to disk. No data to write for {local_location}')
            return False
        # Create folder if needed
        local_folder = os.path.dirname(local_location)
        create_folder_if_needed(local_folder, create_dir=True)
        with open(local_location, 'wb') as output:
            print('Saved {0}'.format(local_location))
            output.write(file)
            return True
    except Exception as e:
        print(f'Could not download file {url} to {local_location}: {e}')
        return False


class FileToDownload(NamedTuple):
    remote_url: str
    local_filepath: str


def batch_download_file_to_location(files_to_download: List[FileToDownload], cache=False):
    """
    Downloads a list of URLs using a ThreadPoolExecutor
    Args:
        cache: Whether the session's cache should be used
        files_to_download: List of URLs to download and target local files to download to
    """
    # See https://docs.python.org/3/library/concurrent.futures.html
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        future_to_url = {executor.submit(
            download_file_to_location,
            file_to_download.remote_url,
            file_to_download.local_filepath,
            cache
        ): file_to_download for file_to_download in files_to_download}
        for future in concurrent.futures.as_completed(future_to_url):
            file_to_download = future_to_url[future]
            try:
                file_downloaded = future.result()
                if not file_downloaded:
                    print(f'Could not download {file_to_download.remote_url}')
            except Exception as exc:
                print('%r generated an exception: %s' % (file_to_download, exc))
