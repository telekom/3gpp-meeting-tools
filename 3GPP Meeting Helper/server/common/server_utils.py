import concurrent.futures
import os.path
import re
import socket
import traceback
from typing import NamedTuple, List

import html2text

from config.networking import private_server, public_server, wg_folder_public_server, wg_folder_private_server
from server.common.server_enums import ServerType, DocumentType, TdocType, WorkingGroup, DocumentFileType
from server.connection import get_remote_file
from utils.local_cache import get_sa2_root_folder_local_cache, create_folder_if_needed

"""Retrieves data from the 3GPP web server"""

sync_folder = 'ftp/Meetings_3GPP_SYNC/SA2/'

host_public_server = 'https://' + public_server
host_private_server = 'http://' + private_server
sa2_url = host_public_server + '/' + wg_folder_public_server
sa2_url_sync = host_public_server + '/' + sync_folder
sa2_url_private_server = host_private_server + '/' + wg_folder_private_server

tdocs_by_agenda_for_checking_meeting_number_in_meeting = 'http://10.10.10.10/ftp/SA/SA2/TdocsByAgenda.htm'


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
        force_download=False,
        convert_html_to_txt=False
) -> bool:
    """
    Downloads a given file to a local location
    Args:
        convert_html_to_txt: Whether to finally convert the downloaded HTML file to TXT
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

        [root, ext] = os.path.splitext(local_location)
        if convert_html_to_txt and (ext=='.htm' or ext=='.html'):
            txt_location = f'{root}.txt'
            h = html2text.HTML2Text()
            h.ignore_links = False
            h.body_width = 0

            with open(local_location, 'r', encoding="utf-8") as f:
                html = f.read()
            text = h.handle(html)
            with open(txt_location, 'w', encoding="utf-8") as out:
                print('Saved {0}'.format(txt_location))
                out.write(text)

        return True

    except Exception as e:
        print(f'Could not download file {url} to {local_location}: {e}')
        return False


class FileToDownload(NamedTuple):
    remote_url: str
    local_filepath: str
    force_download: bool


def batch_download_file_to_location(
        files_to_download: List[FileToDownload],
        cache=False,
        convert_html_to_txt=False):
    """
    Downloads a list of URLs using a ThreadPoolExecutor
    Args:
        convert_html_to_txt: Whether to finally convert downloaded HTML files to TXT
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
            file_to_download.force_download,
            convert_html_to_txt
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


# Used to parse the meeting ID
meeting_id_regex = re.compile(r'.*meeting\?MtgId=(?P<meeting_id>[\d]+)')


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



