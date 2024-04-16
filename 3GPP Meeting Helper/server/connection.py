import re
import traceback
from ftplib import FTP
from typing import NamedTuple, Any, Tuple
from urllib.parse import urlparse

import requests
from cachecontrol import CacheControl

non_cached_http_session = requests.Session()
http_session = CacheControl(non_cached_http_session)

# Avoid getting sometimes 403s
http_session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36'})


class HttpRequestTimeout(NamedTuple):
    # Connect timeout: it is the number of seconds
    # Requests will wait for your client to establish a connection to a remote machine (corresponding to the connect())
    # call on the socket. It’s a good practice to set connect timeouts to slightly larger than a multiple of 3,
    # which is the default TCP packet retransmission window.
    connect_timeout: Any
    # Read timeout: Once your client has connected to the server and sent the
    # HTTP request, the read timeout is the number of seconds the client will wait for the server to send a response. (
    # Specifically, it’s the number of seconds that the client will wait between bytes sent from the server. In 99.9% of
    # cases, this is the time before the server sends the first byte).
    read_timeout: Any


# Timeout set for 3 seconds for connect, 15 seconds for the transmission itself. See
# https://requests.readthedocs.io/en/latest/user/advanced/ -> Timeouts
timeout_values = HttpRequestTimeout(3.05, 6)


def get_html(
        url,
        cache=True,
        try_update_folders=True,
        file_to_return_if_error=None,
        timeout: HttpRequestTimeout = None
) -> bytes | None:
    """
    Downloads a given HTML file
    Args:
        url: The URL of the file
        cache: Whether HTTP cache should be used (default=Yes)
        try_update_folders: Used for FTP retrieval
        file_to_return_if_error: Can override an error, in which case this file is returned
        timeout: Timeout value for the HTTP connection
        # ToDo if needed: Content-Disposition: attachment;filename="TDoc_List_Meeting_SA2#162.xlsx"

    Returns:

    """
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
            if timeout is None:
                timeout_tuple = (timeout_values.connect_timeout, timeout_values.read_timeout)
            else:
                timeout_tuple = (timeout.connect_timeout, timeout.read_timeout)
            if cache:
                r = http_session.get(url, timeout=timeout_tuple)
            else:
                r = non_cached_http_session.get(url, timeout=timeout_tuple)
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


folder_ftp_names_regex = re.compile(r'[\d-]+[ ]+.*[ ]+<DIR>[ ]+(.*[uU][pP][dD][aA][tT][eE].*)')
