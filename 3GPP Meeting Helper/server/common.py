import concurrent.futures
import datetime
import os.path
import re
import socket
import traceback
from dataclasses import dataclass
from enum import Enum
from functools import cached_property
from typing import NamedTuple, List

import tdoc.utils
import utils

from config.networking import private_server, public_server, wg_folder_public_server, wg_folder_private_server
from server.connection import get_remote_file
from utils.local_cache import get_sa2_root_folder_local_cache, create_folder_if_needed, file_exists

"""Retrieves data from the 3GPP web server"""

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
    SYNC = 3


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

    @staticmethod
    def from_string(wg_str_from_tdoc: str):
        wg_str_from_tdoc = wg_str_from_tdoc.upper()
        match wg_str_from_tdoc:
            case 'SP':
                return WorkingGroup.SP
            case 'S1':
                return WorkingGroup.S1
            case 'S2':
                return WorkingGroup.S2
            case 'S3':
                return WorkingGroup.S3
            case 'S3LI':
                return WorkingGroup.S3LI
            case 'S4':
                return WorkingGroup.S4
            case 'S5':
                return WorkingGroup.S5
            case 'S6':
                return WorkingGroup.S6
            case 'RP':
                return WorkingGroup.RP
            case 'R1':
                return WorkingGroup.R1
            case 'R2':
                return WorkingGroup.R2
            case 'R3':
                return WorkingGroup.R3
            case 'R4':
                return WorkingGroup.R4
            case 'R5':
                return WorkingGroup.R5
            case 'CP':
                return WorkingGroup.CP
            case 'C1':
                return WorkingGroup.C1
            case 'C3':
                return WorkingGroup.C3
            case 'C4':
                return WorkingGroup.C4
            case 'C6':
                return WorkingGroup.C6
            case _:
                print(f'Could not parse WG {wg_str_from_tdoc}. Returning SA WG2')
                return WorkingGroup.S2

    def get_wg_folder_name(self, server_type: ServerType) -> str:
        # Groups have different names depending on where you access them!
        match server_type:
            case ServerType.PRIVATE:
                # When we are connected to 10.10.10.10
                # See https://www.3gpp.org/ftp/Meetings_3GPP_SYNC
                prefix = 'ftp'
                if server_type == ServerType.SYNC:
                    prefix = f'{prefix}/Meetings_3GPP_SYNC'
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
                        return f'{prefix}/CT/CT1'
                    case WorkingGroup.C3:
                        return f'{prefix}/CT/CT3'
                    case WorkingGroup.C4:
                        return f'{prefix}/CT/CT4'
                    case WorkingGroup.C6:
                        return f'{prefix}/CT/CT6'
                    case WorkingGroup.RP:
                        return f'{prefix}/RAN'
                    case WorkingGroup.R1:
                        return f'{prefix}/RAN/RAN1'
                    case WorkingGroup.R2:
                        return f'{prefix}/RAN/RAN2'
                    case WorkingGroup.R3:
                        return f'{prefix}/RAN/RAN3'
                    case WorkingGroup.R4:
                        return f'{prefix}/RAN/RAN4'
                    case WorkingGroup.R5:
                        return f'{prefix}/RAN/RAN5'
            case ServerType.SYNC:
                # When we are connected to the sync server
                # See https://www.3gpp.org/ftp/Meetings_3GPP_SYNC
                prefix = 'ftp'
                if server_type == ServerType.SYNC:
                    prefix = f'{prefix}/Meetings_3GPP_SYNC'
                match self:
                    case WorkingGroup.SP:
                        return f'{prefix}/SA'
                    case WorkingGroup.S1:
                        return f'{prefix}/SA1'
                    case WorkingGroup.S2:
                        return f'{prefix}/SA2'
                    case WorkingGroup.S3:
                        return f'{prefix}/SA3'
                    case WorkingGroup.S3LI:
                        return f'{prefix}/SA3LI'
                    case WorkingGroup.S4:
                        return f'{prefix}/SA4'
                    case WorkingGroup.S5:
                        return f'{prefix}/SA5'
                    case WorkingGroup.S6:
                        return f'{prefix}/SA6'
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

    def get_wg_inbox_folder(self, server_type: ServerType) -> str:
        """
        Returns the inbox folder for this meeting
        Args:
            server_type (object):
        """
        wg_folder = self.get_wg_folder_name(server_type)
        inbox_folder = f'{wg_folder}/Inbox'
        return inbox_folder


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

    wg_folder = working_group.get_wg_folder_name(server_type)
    match document_type:
        case DocumentType.CHAIR_NOTES:
            folders = [
                f'{wg_folder}/INBOX/Chair_Notes'
            ] if server_type == ServerType.PRIVATE or server_type == ServerType.SYNC \
                else [
                f'{wg_folder}/{meeting_folder_in_server}/INBOX/Chair_Notes'
            ]
        case DocumentType.TDOCS_BY_AGENDA | DocumentType.MEETING_ROOT:
            folders = [
                f'{wg_folder}/'
            ] if server_type == ServerType.PRIVATE or server_type == ServerType.SYNC \
                else [
                f'{wg_folder}/{meeting_folder_in_server}/'
            ]
        case DocumentType.AGENDA:
            folders = [
                f'{wg_folder}/Agenda/',
                f'{wg_folder}/INBOX/DRAFTS/_Session_Plan_Updates/',
                f'{wg_folder}/INBOX/Schedule_Updates/'
            ] if server_type == ServerType.PRIVATE or server_type == ServerType.SYNC \
                else [
                f'{wg_folder}/{meeting_folder_in_server}/Agenda/',
                f'{wg_folder}/{meeting_folder_in_server}/INBOX/DRAFTS/_Session_Plan_Updates/',
                f'{wg_folder}/{meeting_folder_in_server}/INBOX/Schedule_Updates/'
            ]
        case _:
            # A TDoc
            match tdoc_type:
                case None | TdocType.NORMAL:
                    # Normal TDoc
                    folders = [
                        f'{wg_folder}/Docs/',
                        f'{wg_folder}/INBOX/'
                    ] if server_type == ServerType.PRIVATE or server_type == ServerType.SYNC \
                        else [
                        f'{wg_folder}/{meeting_folder_in_server}/Docs/',
                        f'{wg_folder}/{meeting_folder_in_server}/INBOX/'
                    ]
                case TdocType.DRAFT:
                    # Draft TDoc (sub-folders not included!)
                    folders = [
                        f'{wg_folder}/INBOX/DRAFTS/'
                    ] if server_type == ServerType.PRIVATE or server_type == ServerType.SYNC \
                        else [f'{wg_folder}/{meeting_folder_in_server}/INBOX/DRAFTS/']
                case _:
                    # Revision
                    # No revisions in F2F meetings (at least during the F2F phase)
                    folders = [] if server_type == ServerType.PRIVATE or server_type == ServerType.SYNC \
                        else [
                        f'{wg_folder}/{meeting_folder_in_server}/INBOX/Revisions/',
                        f'{wg_folder}/{meeting_folder_in_server}/INBOX/e-mail_Approval/Revisions/']
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
        cache=False,
        force_download=False
) -> bool:
    """
    Downloads a given file to a local location
    Args:
        cache: Whether to use HTTP caching
        url: The URL to download
        local_location: Where to download the file to
        force_download: Whether to force a download

    Returns:
        bool: Whether the file could be successfully downloaded
    """
    try:
        if force_download:
            use_cached_file_if_available = False
        else:
            use_cached_file_if_available = True
        file = get_remote_file(
            url,
            cache=cache,
            use_cached_file_if_available=use_cached_file_if_available,
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
    force_download: bool


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
            cache,
            file_to_download.force_download
        ): file_to_download for file_to_download in files_to_download}
        for future in concurrent.futures.as_completed(future_to_url):
            file_to_download = future_to_url[future]
            try:
                file_downloaded = future.result()
                if not file_downloaded:
                    print(f'Could not download {file_to_download.remote_url}')
            except Exception as exc:
                print('%r generated an exception: %s' % (file_to_download, exc))


# Points to the 3GPP meeting information for each TSG, WG
meeting_pages_per_group: dict[str, str] = {
    'SP': 'https://www.3gpp.org/dynareport?code=Meetings-SP.htm',
    'S1': 'https://www.3gpp.org/dynareport?code=Meetings-S1.htm',
    'S2': 'https://www.3gpp.org/dynareport?code=Meetings-S2.htm',
    'S3': 'https://www.3gpp.org/dynareport?code=Meetings-S3.htm',
    'S4': 'https://www.3gpp.org/dynareport?code=Meetings-S4.htm',
    'S5': 'https://www.3gpp.org/dynareport?code=Meetings-S5.htm',
    'S6': 'https://www.3gpp.org/dynareport?code=Meetings-S6.htm',
    'CP': 'https://www.3gpp.org/dynareport?code=Meetings-CP.htm',
    'C1': 'https://www.3gpp.org/dynareport?code=Meetings-C1.htm',
    'C3': 'https://www.3gpp.org/dynareport?code=Meetings-C3.htm',
    'C4': 'https://www.3gpp.org/dynareport?code=Meetings-C4.htm',
    'C6': 'https://www.3gpp.org/dynareport?code=Meetings-C6.htm',
    'RP': 'https://www.3gpp.org/dynareport?code=Meetings-RP.htm',
    'R1': 'https://www.3gpp.org/dynareport?code=Meetings-R1.htm',
    'R2': 'https://www.3gpp.org/dynareport?code=Meetings-R2.htm',
    'R3': 'https://www.3gpp.org/dynareport?code=Meetings-R3.htm',
    'R4': 'https://www.3gpp.org/dynareport?code=Meetings-R4.htm',
    'R5': 'https://www.3gpp.org/dynareport?code=Meetings-R5.htm',
}

meeting_ftp_pages_per_group: dict[str, str] = {
    'SP': 'https://www.3gpp.org/ftp/tsg_sa/TSG_SA',
    'S1': 'https://www.3gpp.org/ftp/tsg_sa/WG1_Serv',
    'S2': 'https://www.3gpp.org/ftp/tsg_sa/WG2_Arch',
    'S3': 'https://www.3gpp.org/ftp/tsg_sa/WG3_Security',
    'S4': 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC',
    'S5': 'https://www.3gpp.org/ftp/tsg_sa/WG5_TM',
    'S6': 'https://www.3gpp.org/ftp/tsg_sa/WG6_MissionCritical',
    'CP': 'https://www.3gpp.org/ftp/tsg_ct/TSG_CT',
    'C1': 'https://www.3gpp.org/ftp/tsg_ct/WG1_mm-cc-sm_ex-CN1',
    'C3': 'https://www.3gpp.org/ftp/tsg_ct/WG2_capability_ex-T2',
    'C4': 'https://www.3gpp.org/ftp/tsg_ct/WG3_interworking_ex-CN3',
    'C6': 'https://www.3gpp.org/ftp/tsg_ct/WG4_protocollars_ex-CN4',
    'RP': 'https://www.3gpp.org/ftp/tsg_ran/TSG_RAN',
    'R1': 'https://www.3gpp.org/ftp/tsg_ran/WG1_RL1',
    'R2': 'https://www.3gpp.org/ftp/tsg_ran/WG2_RL2',
    'R3': 'https://www.3gpp.org/ftp/tsg_ran/WG3_Iu',
    'R4': 'https://www.3gpp.org/ftp/tsg_ran/WG4_Radio',
    'R5': 'https://www.3gpp.org/ftp/tsg_ran/WG5_Test_ex-T1',
}


# 3GPP Forge API repository
apis_3gpp_forge_url = 'https://forge.3gpp.org/rep/all/5G_APIs/'


def get_tdoc_details_url(tdoc_id: str):
    # e.g. https://portal.3gpp.org/ngppapp/CreateTDoc.aspx?mode=view&contributionUid=SP-241424
    return f'https://portal.3gpp.org/ngppapp/CreateTDoc.aspx?mode=view&contributionUid={tdoc_id}'


@dataclass(frozen=True)
class MeetingEntry:
    meeting_group: str
    meeting_number: str
    meeting_url_3gu: str
    meeting_name: str
    meeting_location: str
    meeting_url_invitation: str
    start_date: datetime.datetime
    meeting_url_agenda: str
    end_date: datetime.datetime
    meeting_url_report: str
    tdoc_start: tdoc.utils.GenericTdoc | None
    tdoc_end: tdoc.utils.GenericTdoc | None
    meeting_url_docs: str
    meeting_folder_url: str

    @property
    def meeting_folder(self) -> str | None:
        """
        The remote meeting folder name in the 3GPP server's group directory based on the meeting_folder URL
        Returns: The remote folder of the meeting in the 3GPP server. If the folder URL is not set, it may return None

        """
        folder_url = self.meeting_folder_url
        if folder_url is None or folder_url == '':
            return folder_url
        split_folder_url = [f for f in folder_url.split('/') if f != '']
        return split_folder_url[-1]

    @cached_property
    def meeting_id(self) -> str | None:
        """
        Parses the meeting ID from the Meeting's URL. This ID is used in 3GU to identify the meeting, e.g.
        https://portal.3gpp.org/Home.aspx#/meeting?MtgId=60623 -> 60623
        Returns: The ID of the meeting. None if the ID could not be parsed

        """
        if self.meeting_url_3gu is None:
            return None

        id_match = meeting_id_regex.match(self.meeting_url_3gu)

        if id_match is None:
            return None

        return id_match.group('meeting_id')

    @cached_property
    def meeting_calendar_ics_url(self) -> str | None:
        """
        Generates a URL for the 3GPP server containing the calendar entry in ICS format
        Returns: The URL of the ICS file

        """
        the_meeting_id = self.meeting_id
        if the_meeting_id is None:
            return None
        return f"https://portal.3gpp.org/webservices/Rest/Meetings.svc/GetiCal/{the_meeting_id}.ics"

    @cached_property
    def meeting_tdoc_list_url(self) -> str | None:
        """
        Returns, based on the meeting ID, the TDoc list URL from the 3GPP portal
        Returns: The URL, None if the meeting ID is not available/parseable
        """
        meeting_id = self.meeting_id
        if meeting_id is None:
            return None

        # e.g. https://portal.3gpp.org/ngppapp/TdocList.aspx?meetingId=60394
        return 'https://portal.3gpp.org/ngppapp/TdocList.aspx?meetingId=' + meeting_id

    @cached_property
    def meeting_tdoc_list_excel_url(self) -> str | None:
        """
        Returns, based on the meeting ID, the TDoc list URL for the Excel file from the 3GPP portal
        Returns: The URL, None if the meeting ID is not available/parseable
        """
        meeting_id = self.meeting_id
        if meeting_id is None:
            return None

        # e.g. https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=60394
        return 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=' + meeting_id

    def get_tdoc_url(self, tdoc_to_get: tdoc.utils.GenericTdoc | str):
        """
        For a string containing a potential TDoc, returns a URL concatenating the Docs folder and the input TDoc and
        adds a .'zip' extension.
        Args:
            tdoc_to_get: A TDoc ID. Either an object (GenericTdoc) or string. Note that the input is NOT checked!

        Returns: A URL

        """
        if isinstance(tdoc_to_get, tdoc.utils.GenericTdoc):
            tdoc_file = tdoc_to_get.__str__() + '.zip'
        else:
            tdoc_file = tdoc_to_get + '.zip'
        return self.meeting_url_docs + tdoc_file

    def get_tdoc_inbox_url(self, tdoc_to_get: tdoc.utils.GenericTdoc | str):
        """
        For a string containing a potential TDoc, returns a URL concatenating the Inbox folder and the input TDoc and
        adds a .'zip' extension.
        Args:
            tdoc_to_get: A TDoc ID. Either an object (GenericTdoc) or string. Note that the input is NOT checked!

        Returns: A URL

        """
        docs_url = self.get_tdoc_url(tdoc_to_get)
        try:
            inbox_url = re.sub('Docs', 'Inbox', string=docs_url, flags=re.IGNORECASE)
            return inbox_url
        except Exception as e:
            print(f'Could not generate inbox URL, returning Docs URL: {e}')
            return docs_url

    @cached_property
    def local_folder_path(self) -> str | None:
        """
        For a given meeting, returns the cache folder and creates it if it does not exist
        Returns:

        """
        folder_name = self.meeting_folder
        if folder_name is None:
            return None
        full_path = os.path.join(utils.local_cache.get_cache_folder(), folder_name)
        return full_path

    @property
    def local_agenda_folder_path(self) -> str:
        """
        For a given meeting, returns the cache folder located at meeting_folder/Agenda and creates
        it if it does not exist
        Returns:

        """
        full_path = os.path.join(self.local_folder_path, 'Agenda')
        utils.local_cache.create_folder_if_needed(full_path, create_dir=True)
        return full_path

    @property
    def local_export_folder_path(self) -> str:
        """
        For a given meeting, returns the cache folder located at meeting_folder/Export and creates
        it if it does not exist
        Returns:

        """
        full_path = os.path.join(self.local_folder_path, 'Export')
        utils.local_cache.create_folder_if_needed(full_path, create_dir=True)
        return full_path

    @cached_property
    def local_tdoc_list_excel_path(self):
        return os.path.join(self.local_agenda_folder_path, 'TDoc_List.xlsx')

    @cached_property
    def is_li(self):
        return '-LI' in self.meeting_number

    @cached_property
    def meeting_folders_3gpp_wifi_url(self) -> List[str]:
        wg = WorkingGroup.from_string(self.meeting_group)
        candidate_folders = get_document_or_folder_url(
            server_type=ServerType.PRIVATE,
            document_type=DocumentType.TDOC,
            meeting_folder_in_server='',
            tdoc_type=TdocType.NORMAL,
            working_group=wg
        )
        return candidate_folders

    @cached_property
    def working_group_enum(self) -> WorkingGroup:
        return WorkingGroup.from_string(self.meeting_group)

    def get_tdoc_3gpp_wifi_url(self, tdoc_id_str: str) -> List[str]:
        candidate_folders = self.meeting_folders_3gpp_wifi_url
        candidate_urls = [f'{f}{tdoc_id_str}.zip' for f in candidate_folders]
        return candidate_urls

    @property
    def meeting_is_now(self) -> bool:
        if self.start_date is None or self.end_date is None:
            return False

        # Add some time delta
        days_delta = datetime.timedelta(days=3)
        if self.start_date - days_delta < datetime.datetime.now() < self.end_date + days_delta:
            return True
        return False

    @cached_property
    def local_server_url(self):
        return f'{host_private_server}/{self.working_group_enum.get_wg_folder_name(ServerType.PRIVATE)}'

    @cached_property
    def sync_server_url(self):
        return f'{host_private_server}/{self.working_group_enum.get_wg_folder_name(ServerType.SYNC)}'

    def get_tdoc_local_path(self, tdoc_str: str) -> str | None:
        """
        Generates the local path for a given TDoc
        Args:
            tdoc_str: The TDoc for which the local path is queried

        Returns: The TDoc local path. None if it could not be generated, e.g. if the local folder cannot be established.
        """
        local_folder = self.local_folder_path
        if local_folder is None:
            return None
        local_file = os.path.join(
            local_folder,
            str(tdoc_str),
            f'{tdoc_str}.zip')
        local_file.replace(f'{os.path.pathsep}{os.path.pathsep}', f'{os.path.pathsep}')
        return local_file

    @cached_property
    def tdoc_excel_local_path(self)->str|None:
        download_folder = self.local_agenda_folder_path
        if download_folder is None:
            return None
        return os.path.join(download_folder, f'{self.meeting_name}_TDoc_List.xlsx')

    @property
    def tdoc_excel_exists_in_local_folder(self)->bool | None:
        local_path = self.tdoc_excel_local_path
        if local_path is None:
            return None
        return file_exists(local_path)

    def starts_in_given_year(self, year:int) -> bool:
        if self.start_date is None:
            return False
        return self.start_date.year == year


# Used to parse the meeting ID
meeting_id_regex = re.compile(r'.*meeting\?MtgId=(?P<meeting_id>[\d]+)')

class DocumentFileType(Enum):
    UNKNOWN = 0
    DOCX = 1
    DOC = 2
    PPTX = 3
    PDF = 4
    HTML = 5
    YAML = 6


class DownloadedTdocDocument(NamedTuple):
    title: str | None
    source: str | None
    url: str | None
    tdoc_id: str | None
    path: str | None

    @property
    def document_type(self) -> DocumentFileType:
        if self.path is None:
            return DocumentFileType.UNKNOWN
        try:
            extension = os.path.splitext(self.path)[1].replace('.','').lower()
        except Exception as e:
            print(f'Could not get file type for TDoc {self.tdoc_id}, {self.path}: {e}')
            return DocumentFileType.UNKNOWN

        match extension:
            case 'pdf':
                return DocumentFileType.PDF
            case 'html':
                return DocumentFileType.HTML
            case 'doc':
                return DocumentFileType.DOC
            case 'docx':
                return DocumentFileType.DOCX
            case 'pptx':
                return DocumentFileType.PPTX
            case 'yaml':
                return DocumentFileType.YAML
            case _:
                return DocumentFileType.UNKNOWN


class DownloadedData(NamedTuple):
    folder_path: str | None # Folder where the downloaded data is placed
    downloaded_word_documents: List[DownloadedTdocDocument] | None  # Downloaded Word Documents


class WiEntry(NamedTuple):
    uid: str
    code: str
    title: str
    release: str
    lead_body: str

    @property
    def cr_list_url(self) -> str:
        cr_list_url = f'https://portal.3gpp.org/ChangeRequests.aspx?q=1&specnumber=&release=all&workitem={self.uid}'
        return cr_list_url

    @property
    def spec_list_url(self) -> str:
        spec_list_url = f'https://portal.3gpp.org/Specifications.aspx?q=1&WiUid={self.uid}'
        return spec_list_url

    @property
    def wid_page_url(self) -> str:
        sid_page_url = f'https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId={self.uid}'
        return sid_page_url

    @property
    def wid_lead_body_list(self) -> List[str]:
        """
        Since the lead body may contain a list of comma-separated values, this property exposes an actual list of
        lead bodies that can be used to generate the URLs to the 3GPP site Returns: List of lead boies, e.g. [R3, S2]
        """
        lead_bodies = [body.strip() for body in self.lead_body.split(',')]
        return lead_bodies

    @property
    def wid_lead_body_list_urls(self) -> List[str]:
        lead_bodies = [f'https://www.3gpp.org/dynareport?code=TSG-WG--{body}--wis.htm' for
                       body in self.wid_lead_body_list]
        return lead_bodies
