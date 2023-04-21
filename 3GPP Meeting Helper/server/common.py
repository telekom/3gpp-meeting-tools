import os
import re
import socket
import traceback
import zipfile
from ftplib import FTP
from urllib.parse import urlparse

import requests
from cachecontrol import CacheControl

root_folder = '3GPP_SA2_Meeting_Helper'

non_cached_http_session = requests.Session()
http_session = CacheControl(non_cached_http_session)

# Avoid getting sometimes 403s
http_session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36'})

# Timeout set for 3 seconds for connect, 15 seconds for the transmission itself. See
# https://requests.readthedocs.io/en/latest/user/advanced/ -> Timeouts
timeout_values = (3.05, 15)

folder_ftp_names_regex = re.compile(r'[\d-]+[ ]+.*[ ]+<DIR>[ ]+(.*[uU][pP][dD][aA][tT][eE].*)')

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
user_folder = '~'

def get_html(url, cache=True, try_update_folders=True, file_to_return_if_error=None):
    if file_to_return_if_error is not None:
        print('Returning {0} in case of HTTP(s) error'.format(file_to_return_if_error))
    try:
        o = urlparse(url)
    except:
        # Not an URL
        print('{0} not an URL'.format(url))
        return None
    try:
        if (o.scheme == 'http') or (o.scheme == 'https'):
            print('HTTP GET {0}'.format(url))
            if cache:
                r = http_session.get(url, timeout=timeout_values)
            else:
                r = non_cached_http_session.get(url, timeout=timeout_values)
            if r.status_code != 200:
                print('HTTP GET {0}: {1}'.format(url, r.status_code))
                if file_to_return_if_error is not None:
                    try:
                        with open(file_to_return_if_error, "rb") as f:
                            cached_file = f.read()
                            print("Returning cached file {0}".format(file_to_return_if_error))
                            return cached_file
                    except:
                        print("Could not read file {0}".format(file_to_return_if_error))
                        return None
                else:
                    return None

            html_content = r.content
            if file_to_return_if_error is not None:
                try:
                    # Write to cache
                    with open(file_to_return_if_error, "wb") as file:
                        file.write(html_content)
                    print('Cached content to to {0}'.format(file_to_return_if_error))
                except:
                    traceback.print_exc()
                    print('Could not cache file to {0}'.format(file_to_return_if_error))
            return html_content
        elif o.scheme == 'ftp':
            # Do FTP download
            print('FTP {0} RETR {1}'.format(o.netloc, o.path))
            try:
                with FTP(o.netloc) as ftp:
                    ftp.login()
                    # https://stackoverflow.com/questions/18772703/read-a-file-in-buffer-from-ftp-python/18773444
                    data = []

                    def handle_binary(more_data):
                        data.append(more_data)

                    try:
                        ftp.retrbinary('RETR {0}'.format(o.path), callback=handle_binary)
                    except Exception as ftp_exception:
                        if not try_update_folders:
                            raise ftp_exception

                        # Try into the "_Update" folders in Inbox and outside of Inbox
                        # Example address: '/SA/SA2/Inbox/S2-1912194.zip'
                        split_address = o.path.split('/')
                        tdoc_id = split_address[-1]
                        folders_to_test = list({'/'.join(split_address[0:-1]), '/'.join(split_address[0:-2])})
                        update_folders = []
                        for folder_to_test in folders_to_test:
                            dir_data = []

                            def handle_binary_dir(more_data):
                                dir_data.append(more_data)

                            try:
                                ftp.cwd(folder_to_test)
                                ftp.retrlines('LIST', handle_binary_dir)
                                folder_content_matches = [folder_ftp_names_regex.match(e) for e in dir_data]
                                folder_content_matches = ['{0}/{1}'.format(folder_to_test, e.group(1)) for e in
                                                          folder_content_matches if e is not None]
                                update_folders.extend(folder_content_matches)
                            except:
                                print('Could not scan directories in dir {0} in FTP server'.format(folder_to_test))
                        found_in_update_folder = False
                        last_exception = ftp_exception
                        if len(update_folders) > 0:
                            print('Searching for TDocs in update folders: {0}'.format(', '.join(update_folders)))
                        for update_folder in update_folders:
                            try:
                                data = []
                                ftp.retrbinary('RETR {0}/{1}'.format(update_folder, tdoc_id), callback=handle_binary)
                                found_in_update_folder = True
                                print('Found TDoc {0} in {1}'.format(tdoc_id, update_folder))
                            except Exception as x:
                                last_exception = x
                        if not found_in_update_folder:
                            raise last_exception

                    # https://stackoverflow.com/questions/17068100/joining-byte-list-with-python
                    data = b''.join(data)
                    return data
            except:
                print('FTP {0} RETR {1} ERROR'.format(o.netloc, o.path))
    except:
        if file_to_return_if_error is not None:
            try:
                # Read from cache
                with open(file_to_return_if_error, "rb") as file:
                    file_content = file.read()
                print('Could not load from {1}. Read cached content from {0}'.format(file_to_return_if_error, url))
                return file_content
            except:
                print('Could not read cache file from {0}'.format(file_to_return_if_error))
                return None
        else:
            traceback.print_exc()

        return None


def create_folder_if_needed(folder_name, create_dir):
    if create_dir and (not os.path.exists(folder_name)):
        os.makedirs(folder_name, exist_ok=True)


def decode_string(str_to_decode, log_name):
    try:
        return str_to_decode.decode('utf-8')
    except:
        print('Could not decode {0} using UTF-8. Trying cp1252'.format(log_name))
        try:
            decoded_html_data = str_to_decode.decode('cp1252')
            print('Successfully decoded data as cp1252')
            return decoded_html_data
        except:
            print('Could not decode {0} using cp1252. Returning HTML as-is'.format(log_name))
            return str_to_decode


def get_tmp_folder(create_dir=True):
    folder_name = os.path.expanduser(os.path.join(user_folder, root_folder, 'tmp'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name


def get_spec_folder(create_dir=True):
    folder_name = os.path.expanduser(os.path.join(user_folder, root_folder, 'specs'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name


def get_local_docs_filename(meeting_folder_name):
    folder = get_local_agenda_folder(meeting_folder_name, create_dir=True)
    return os.path.join(folder, 'Docs.htm')


def get_local_revisions_filename(meeting_folder_name):
    folder = get_local_agenda_folder(meeting_folder_name, create_dir=True)
    return os.path.join(folder, 'Revisions.htm')


def get_local_drafts_filename(meeting_folder_name):
    folder = get_local_agenda_folder(meeting_folder_name, create_dir=True)
    return os.path.join(folder, 'Drafts.htm')


def get_sa2_tdoc_list(meeting_folder_name):
    url = get_remote_meeting_folder(meeting_folder_name, use_inbox=False) + 'TdocsByAgenda.htm'
    return get_html(url)


def get_local_agenda_folder(meeting_folder_name, create_dir=True):
    local_folder_for_this_meeting = get_meeting_folder(meeting_folder_name)
    folder_name = os.path.join(local_folder_for_this_meeting, 'Agenda')
    if create_dir and (not os.path.exists(folder_name)):
        os.makedirs(folder_name, exist_ok=True)
    return folder_name


def get_meeting_folder(meeting_folder_name, create_dir=False):
    folder_name = os.path.join(get_cache_folder(create_dir=create_dir), meeting_folder_name)
    create_folder_if_needed(folder_name, create_dir)
    return folder_name


def get_cache_folder(create_dir=False):
    folder_name = os.path.expanduser(os.path.join(user_folder, root_folder, 'cache'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name


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
    html = get_html(sa2_url, file_to_return_if_error=get_sa2_root_folder_local_cache())
    return html


def get_sa2_root_folder_local_cache(create_dir=True):
    cache_folder = get_cache_folder(create_dir)
    inbox_cache = os.path.join(cache_folder, 'Wg2ArchCache.html')
    return inbox_cache


def unzip_files_in_zip_file(zip_file):
    tdoc_folder = os.path.split(zip_file)[0]
    zip_ref = zipfile.ZipFile(zip_file, 'r')
    files_in_zip = zip_ref.namelist()
    # Check if is there any file in the zip that does not exist. If not, then do not extract need_to_extract = any(
    # item == False for item in map(os.path.isfile, map(lambda x: os.path.join(tdoc_folder, x), files_in_zip)))
    # Removed check whether extracting is needed, as some people reused the same file name on different document
    # versions... Added exception catch as the file may probably be already open
    try:
        zip_ref.extractall(tdoc_folder)
    except:
        print('Could not extract files')
        traceback.print_exc()
    return [os.path.join(tdoc_folder, file) for file in files_in_zip]


def download_file_to_location(url, local_location):
    try:
        file = get_html(url, cache=False)
        with open(local_location, 'wb') as output:
            print('Saved {0}'.format(local_location))
            output.write(file)
    except:
        print('Could not download file {0} to {1}'.format(url, local_location))
        traceback.print_exc()


def write_data_and_open_file(data, local_file):
    """
    Writes input data to a binary file
    Args:
        data: The data to write
        local_file: The local file to which to write

    Returns:

    """
    if data is None:
        return
    with open(local_file, 'wb') as output:
        output.write(data)
