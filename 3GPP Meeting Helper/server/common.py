import concurrent.futures
import re
import socket
import traceback
from typing import NamedTuple, List

from server.connection import get_remote_file
from utils.local_cache import get_sa2_root_folder_local_cache

"""Retrieves data from the 3GPP web server"""
private_server = '10.10.10.10'
default_http_proxy = 'http://lanbctest:8080'
http_server = 'https://www.3gpp.org/ftp/'
group_folder = 'tsg_sa/WG2_Arch/'
sync_folder = 'Meetings_3GPP_SYNC/SA2/'
meeting_folder = 'SA/SA2/'
sa2_url = ''
sa2_url_sync = ''
sa2_url_meeting = ''


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


def get_sa2_tdoc_list(meeting_folder_name):
    url = get_remote_meeting_folder(meeting_folder_name, use_inbox=False) + 'TdocsByAgenda.htm'
    return get_remote_file(url)


def get_remote_meeting_folder(meeting_folder_name, use_inbox=False, searching_for_a_file=False):
    if not use_inbox:
        folder = sa2_url + meeting_folder_name + '/'
    else:
        folder = get_inbox_url(searching_for_a_file)
    return folder


def get_inbox_url(searching_for_a_file=False):
    return get_inbox_root(searching_for_a_file) + 'Inbox/'


def get_inbox_root(searching_for_a_file=False):
    if not we_are_in_meeting_network(searching_for_a_file):
        folder = sa2_url_sync
    else:
        folder = sa2_url_meeting
    return folder


def we_are_in_meeting_network(searching_for_a_file=False):
    # Since 10.10.10.10 uses only FTP, we will only return it for files, NOT
    # for folder searches
    if not searching_for_a_file:
        return False
    ip_addresses = [i[4][0] for i in socket.getaddrinfo(socket.gethostname(), None)]
    matches = [re.match(r'10.10.(\d)+.(\d)+', ip_address) for ip_address in ip_addresses]
    matches = [match for match in matches if match is not None]
    ip_is_meeting_ip = (len(matches) != 0)
    return ip_is_meeting_ip


def get_sa2_folder():
    html = get_remote_file(sa2_url, file_to_return_if_error=get_sa2_root_folder_local_cache())
    return html


def download_file_to_location(url: str, local_location: str) -> bool:
    """
    Downloads a given file to a local location
    Args:
        url:
        local_location:

    Returns:
        bool: Whether the file could be successfully downloaded
    """
    try:
        file = get_remote_file(url, cache=False)
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


def batch_download_file_to_location(files_to_download: List[FileToDownload]):
    """
    Downloads a list of URLs using a ThreadPoolExecutor
    Args:
        files_to_download: List of URLs to download and target local files to download to
    """
    # See https://docs.python.org/3/library/concurrent.futures.html
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        future_to_url = {executor.submit(
            download_file_to_location,
            file_to_download.remote_url,
            file_to_download.local_filepath): file_to_download for file_to_download in files_to_download}
        for future in concurrent.futures.as_completed(future_to_url):
            file_to_download = future_to_url[future]
            try:
                file_downloaded = future.result()
                if not file_downloaded:
                    print(f'Could not download {file_to_download.remote_url}')
            except Exception as exc:
                print('%r generated an exception: %s' % (file_to_download, exc))


def update_meeting_ftp_server(new_address):
    if (new_address is None) or (new_address == ''):
        return
    global private_server
    private_server = new_address
    update_urls()


def update_urls():
    global sa2_url, sa2_url_sync, sa2_url_meeting
    sa2_url = http_server + group_folder
    sa2_url_sync = http_server + sync_folder
    sa2_url_meeting = 'ftp://' + private_server + '/' + meeting_folder
