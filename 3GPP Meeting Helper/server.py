import application
import requests
from cachecontrol import CacheControl
import os
import os.path
import zipfile
import socket
import re
from urllib.parse import urljoin
import parsing.html as html_parser
import parsing.word as word_parser
import tdoc
from urllib.parse import urlparse
from ftplib import FTP
import traceback
import tdoc
import datetime

http_server        = 'http://www.3gpp.org/ftp/'
group_folder       = 'tsg_sa/WG2_Arch/'
sync_folder        = 'Meetings_3GPP_SYNC/SA2/'
meeting_folder     = 'SA/SA2/'
private_server     = '10.10.10.10'
default_http_proxy = 'http://lanbctest:8080'
root_folder        = '3GPP_SA2_Meeting_Helper'

"""Retrieves data from the 3GPP web server"""
sa2_url         = ''
sa2_url_sync    = ''
sa2_url_meeting = ''
agenda_regex            = re.compile(r'.*Agenda.*[-_](v)?(?P<version>\d*).*\..*')
agenda_docx_regex       = re.compile(r'.*Agenda.*[-_](v)?(?P<version>\d*).*\.(docx|doc)')
agenda_version_regex    = re.compile(r'.*Agenda.*[-_]?(v)(?P<version>\d*).*\..*')
agenda_draft_docx_regex = re.compile(r'.*draft.*Agenda.*[-_](v)?(?P<version>\d*).*\.(docx|doc)')
folder_ftp_names_regex  = re.compile(r'[\d-]+[ ]+.*[ ]+<DIR>[ ]+(.*[uU][pP][dD][aA][tT][eE].*)')

non_cached_http_session = requests.Session()
http_session = CacheControl(non_cached_http_session)

def update_urls():
    global sa2_url
    global sa2_url_sync
    global sa2_url_meeting
    sa2_url         = http_server + group_folder
    sa2_url_sync    = http_server + sync_folder
    sa2_url_meeting = 'ftp://' + private_server + '/' + meeting_folder

update_urls()

def get_sa2_folder():
    html = get_html(sa2_url)
    return html

def get_sa2_inbox_current_tdoc(searching_for_a_file=False):
    url = get_inbox_root(searching_for_a_file) + 'CurDoc.htm'
    return get_html(url)

def get_sa2_inbox_tdoc_list():
    url = get_inbox_root(searching_for_a_file=True) + 'TdocsByAgenda.htm'
    return get_html(url)

def get_sa2_meeting_tdoc_list(meeting_folder):
    remote_folder = get_remote_meeting_folder(meeting_folder)
    url           = remote_folder + 'TdocsByAgenda.htm'
    returned_html = get_html(url)

    # Normal case
    if returned_html is not None:
        return returned_html

    # In some cases, the original TDocsByAgenda was removed (e.g. 136AH meeting). In this case, we have to look for a substitute
    folder_contents = get_html(remote_folder)
    parsed_folder = html_parser.parse_3gpp_http_ftp(folder_contents)
    tdocs_by_agenda_files = [file for file in parsed_folder.files if ('TdocsByAgenda' in file) and (('.htm' in file) or ('.html' in file))]
    if len(tdocs_by_agenda_files) > 0:
        file_to_get = tdocs_by_agenda_files[0]
        url         = remote_folder + file_to_get
        new_html    = get_html(url)
        return new_html
    else:
        return None

def get_sa2_tdoc_list(meeting_folder):
    url = get_remote_meeting_folder(meeting_folder, use_inbox=False) + 'TdocsByAgenda.htm'
    return get_html(url)

def get_html(url, cache=True, try_update_folders=True):
    try:
        o = urlparse(url)
    except:
        # Not an URL
        print('{0} not an URL'.format(url))
        return None
    try:
        if (o.scheme=='http') or (o.scheme=='https'):
            print('HTTP GET {0}'.format(url))
            if cache:
                r = http_session.get(url)
            else:
                r = non_cached_http_session.get(url)
            if r.status_code != 200:
                print('HTTP GET {0}: {1}'.format(url, r.status_code))
                return None
            return r.content
        elif (o.scheme=='ftp'):
            # Do FTP download
            print('FTP {0} RETR {1}'.format(o.netloc, o.path))
            try:
                with FTP(o.netloc) as ftp:
                    ftp.login()
                    # https://stackoverflow.com/questions/18772703/read-a-file-in-buffer-from-ftp-python/18773444
                    data = []
                    def handle_binary(more_data):
                        data.append(more_data)

                    ftp_exception = None
                    try:
                        ftp.retrbinary('RETR {0}'.format(o.path), callback=handle_binary)
                    except Exception as ftp_exception:
                        if not try_update_folders:
                            raise ftp_exception

                        # Try into the "_Update" folders in Inbox and outside of Inbox
                        # Example address: '/SA/SA2/Inbox/S2-1912194.zip'
                        split_address = o.path.split('/')
                        tdoc_id = split_address[-1]
                        folders_to_test = list(set(['/'.join(split_address[0:-1]), '/'.join(split_address[0:-2])]))
                        update_folders = []
                        for folder_to_test in folders_to_test:
                            dir_data = []
                            def handle_binary_dir(more_data):
                                dir_data.append(more_data)
                            try:
                                dir_contents = []
                                ftp.cwd(folder_to_test)
                                ftp.retrlines('LIST', handle_binary_dir) 
                                folder_content_matches = [ folder_ftp_names_regex.match(e) for e in dir_data]
                                folder_content_matches = ['{0}/{1}'.format(folder_to_test, e.group(1)) for e in folder_content_matches if e is not None ]
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
                                ftp.retrbinary('RETR {0}/{1}'.format(update_folder,tdoc_id), callback=handle_binary)
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
        traceback.print_exc()
        return None

def get_tdoc(meeting_folder_name, tdoc_id, use_inbox=False, return_url=False, searching_for_a_file=False):
    if not tdoc.is_tdoc(tdoc_id):
        if not return_url:
            return None
        else:
            return None, None
    tdoc_local_filename = get_local_filename(meeting_folder_name, tdoc_id)
    zip_file_url = get_remote_filename(meeting_folder_name, tdoc_id, use_inbox, searching_for_a_file)
    if not os.path.exists(tdoc_local_filename):
        # TODO: change to also FTP support
        tdoc_file = get_html(zip_file_url, cache=False)
        if tdoc_file is None:
            if use_inbox:
                # Retry without inbox
                return_value = get_tdoc(meeting_folder_name, tdoc_id, use_inbox=False)
            else:
                # No need to retry
                return_value = None
            if not return_url:
                return return_value
            else:
                return return_value, zip_file_url
        # Drive zip file to disk
        with open(tdoc_local_filename,'wb') as output:
            output.write(tdoc_file)

    # If the file does not now exist, there was an error (e.g. not found)
    if not os.path.exists(tdoc_local_filename):
        if not return_url:
            return None
        else:
            return None, None
            
    if not return_url:
        return unzip_tdoc_files(tdoc_local_filename)
    else:
        return unzip_tdoc_files(tdoc_local_filename), zip_file_url

def download_file_to_location(url, location):
    try:
        file = get_html(url, cache=False)
        with open(location,'wb') as output:
            output.write(file)
    except:
        print('Could not download file')
        traceback.print_exc()

def create_folder_if_needed(folder_name, create_dir):
    if create_dir and (not os.path.exists(folder_name)):
        os.makedirs(folder_name, exist_ok=True)

def get_tmp_folder(create_dir=True):
    folder_name = os.path.expanduser(os.path.join('~', root_folder, 'tmp'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name

def get_meeting_folder(meeting_folder_name, create_dir=False):
    folder_name = os.path.expanduser(os.path.join('~', root_folder, 'cache', meeting_folder_name))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name

def get_spec_folder(create_dir=True):
    folder_name = os.path.expanduser(os.path.join('~', root_folder, 'specs'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name

def get_local_folder(meeting_folder_name, tdoc_id, create_dir=True, email_approval=False):
    meeting_folder = get_meeting_folder(meeting_folder_name)
    folder_name = os.path.join(meeting_folder, tdoc_id)
    if email_approval:
        folder_name = os.path.join(folder_name, 'email approval')
    if create_dir and (not os.path.exists(folder_name)):
        os.makedirs(folder_name, exist_ok=True)
    return folder_name

def get_local_filename(meeting_folder_name, tdoc_id, create_dir=True):
    folder_name = get_local_folder(meeting_folder_name, tdoc_id, create_dir)
    return os.path.join(folder_name, tdoc_id + '.zip')

def get_local_agenda_folder(meeting_folder_name, create_dir=True):
    meeting_folder = get_meeting_folder(meeting_folder_name)
    folder_name = os.path.join(meeting_folder, 'Agenda')
    if create_dir and (not os.path.exists(folder_name)):
        os.makedirs(folder_name, exist_ok=True)
    return folder_name

def get_local_invitation_folder(meeting_folder_name, create_dir=True):
    meeting_folder = get_meeting_folder(meeting_folder_name)
    folder_name = os.path.join(meeting_folder, 'Invitation')
    if create_dir and (not os.path.exists(folder_name)):
        os.makedirs(folder_name, exist_ok=True)
    return folder_name

def get_local_tdocs_by_agenda_filename(meeting_folder_name):
    folder = get_local_agenda_folder(meeting_folder_name, create_dir=True)
    return os.path.join(folder, 'TdocsByAgenda.htm')

def get_remote_filename(meeting_folder_name, tdoc_id, use_inbox=False, searching_for_a_file=False):
    folder = get_remote_meeting_folder(meeting_folder_name, use_inbox, searching_for_a_file)
    if not use_inbox:
        folder += 'Docs/'
    elif use_inbox:
        # No need to add 'Docs/'
        pass

    return folder + tdoc_id + '.zip'

def get_remote_agenda_folder(meeting_folder_name, use_inbox=False):
    folder = get_remote_meeting_folder(meeting_folder_name, use_inbox)
    if use_inbox:
        folder += 'Drafts/'
    else:
        folder += 'Agenda/'
    return folder

def get_inbox_url(searching_for_a_file=False):
    return get_inbox_root(searching_for_a_file) + 'Inbox/'

def get_inbox_root(searching_for_a_file=False):
    if not we_are_in_meeting_network(searching_for_a_file):
        folder = sa2_url_sync
    else:
        folder = sa2_url_meeting
    return folder

def get_remote_meeting_folder(meeting_folder_name, use_inbox=False, searching_for_a_file=False):
    if not use_inbox:
        folder = sa2_url + meeting_folder_name + '/'
    else:
        folder = get_inbox_url(searching_for_a_file)
    return folder

def unzip_tdoc_files(zip_file):
    tdoc_folder = os.path.split(zip_file)[0]
    zip_ref = zipfile.ZipFile(zip_file, 'r')
    files_in_zip = zip_ref.namelist()
    # Check if is there any file in the zip that does not exist. If not, then do not extract
    need_to_extract = any(item == False for item in map(os.path.isfile, map(lambda x: os.path.join(tdoc_folder, x), files_in_zip)))
    if need_to_extract:
        zip_ref.extractall(tdoc_folder)
    return [os.path.join(tdoc_folder, file) for file in files_in_zip]

def get_agenda_files(meeting_folder, use_inbox=False):
    url  = get_remote_agenda_folder(meeting_folder, use_inbox=use_inbox)
    html = get_html(url)
    if html is None:
        return
    parsed_folder = html_parser.parse_3gpp_http_ftp(html)
    agenda_files = [file for file in parsed_folder.files if agenda_regex.match(file)]
    agenda_folder = get_local_agenda_folder(meeting_folder)
    real_agenda_files = []
    if len(agenda_files) == 0:
        return
    for agenda_file in agenda_files:
        agenda_url = urljoin(url, agenda_file)
        local_file = os.path.join(agenda_folder, agenda_file)
        filename, file_extension = os.path.splitext(local_file)
        if not os.path.isfile(local_file):
            html = get_html(agenda_url, cache=False)
            if html is None:
                continue
            with open(local_file,'wb') as output:
                output.write(html)
        if file_extension == '.zip':
            unzipped_files = unzip_tdoc_files(local_file)
            real_agenda_files.extend(unzipped_files)
        else:
            real_agenda_files.append(local_file)

def get_latest_agenda_file(agenda_files):
    if agenda_files is None:
        return None

    agenda_files = [ file for file in agenda_files if agenda_docx_regex.match(file) ]
    if len(agenda_files)==0:
        return None

    draft_agenda_files = [ file for file in agenda_files if agenda_draft_docx_regex.match(file) ]
    if (len(draft_agenda_files)>0) and (len(draft_agenda_files)!=len(agenda_files)):
        # Non-draft agendas have priority over draft agenda files
        agenda_files = [ agenda_file for agenda_file in agenda_files if agenda_file not in draft_agenda_files ]

    last_agenda = max(agenda_files, key=get_agenda_file_version_number)
    print('Most recent agenda file seems to be {0}'.format(last_agenda))
    return last_agenda

def get_agenda_file_version_number(x):
    if len(x)==0 or x[0]=='~':
        # Sanity check for empty strings and/or temporary files
        return -1
    tdoc_match = tdoc.tdoc_regex.match(x)
    if tdoc_match is not None:
        tdoc_year = float(tdoc_match.groupdict()['year'])
        tdoc_id   = float(tdoc_match.groupdict()['tdoc_number'])
        tdoc_number = tdoc_year*100000 + tdoc_id
    else:
        tdoc_number = 0
    x_without_dashes = x.replace('-','')
    agenda_match     = agenda_version_regex.match(x_without_dashes)
    agenda_version   = agenda_match.groupdict()['version']
    print('{0} is agenda version {1} for tdoc ID {2:.0f}'.format(x, agenda_version, tdoc_number))

    # Support up to 100 agenda versions. It should be OK...
    agenda_version = tdoc_number + float(agenda_version)/100
    return agenda_version

ai_names_cache = {}
def get_last_agenda(meeting_folder):
    agenda_folder = get_local_agenda_folder(meeting_folder)
    agenda_files = [ file for file in os.listdir(agenda_folder) ]

    last_agenda = get_latest_agenda_file(agenda_files)

    if last_agenda is None:
        return None, None

    agenda_path = os.path.join(agenda_folder,last_agenda)
    try:
        agenda_item_descriptions = word_parser.import_agenda(agenda_path)
    except:
        agenda_item_descriptions = {}
    ai_names_cache[meeting_folder] = agenda_item_descriptions

    return agenda_path, int(agenda_docx_regex.match(last_agenda).groupdict()['version'])

def get_ts_folder(series, release):
    address = 'http://www.3gpp.org/ftp/Specs/latest/Rel-{0}/{1}_series/'.format(release, series)
    return address

def we_are_in_meeting_network(searching_for_a_file=False):
    # Since 10.10.10.10 uses only FTP, we will only return it for files, NOT
    # for folder searches
    if not searching_for_a_file:
        return False
    ip_addresses = [i[4][0] for i in socket.getaddrinfo(socket.gethostname(), None)]
    matches = [re.match('10.10.\d*.\d*', ip_address) for ip_address in ip_addresses]
    matches = [match for match in matches if match is not None]
    ip_is_meeting_ip = (len(matches) != 0)
    return ip_is_meeting_ip

# Begin with updated URLs
update_urls()

def update_meeting_ftp_server(new_address):
    if (new_address is None) or (new_address == ''):
        return
    global private_server
    private_server = new_address
    update_urls()

def get_tdocs_by_agenda_for_selected_meeting(meeting_folder, inbox_active=False):
    # If the inbox is active, we need to download both and return the newest one
    html_inbox = None
    html_3gpp  = None

    datetime_inbox = datetime.datetime.min
    datetime_3gpp  = datetime.datetime.min

    if inbox_active:
        print('Getting TDocs by agenda from inbox')
        html_inbox = get_sa2_inbox_tdoc_list()
        datetime_inbox = html_parser.tdocs_by_agenda.get_tdoc_by_agenda_date(html_inbox)

    print('Getting TDocs by agenda from server')
    html_3gpp = get_sa2_meeting_tdoc_list(meeting_folder)
    datetime_3gpp = html_parser.tdocs_by_agenda.get_tdoc_by_agenda_date(html_3gpp)

    if datetime_3gpp is None:
        datetime_3gpp = datetime.datetime.min

    if datetime_inbox is None:
        datetime_inbox = datetime.datetime.min

    if inbox_active:
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

def download_agenda_file(meeting, inbox_active=False):
    try:
        meeting_server_folder = application.sa2_meeting_data.get_server_folder_for_meeting_choice(meeting)
        local_file            = get_local_tdocs_by_agenda_filename(meeting_server_folder)
        html                  = get_tdocs_by_agenda_for_selected_meeting(meeting_server_folder, inbox_active)
        if html is None:
            print('Agenda file for {0} not found'.format(meeting))
            return None
        tdoc.write_data_and_open_file(html, local_file, open_file=False)
        return local_file
    except:
        print('Could not download agenda file for {0}'.format(meeting))
        traceback.print_exc()
        return None
