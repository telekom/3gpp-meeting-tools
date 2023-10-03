import datetime
import os
import os.path
import re
import traceback
from urllib.parse import urljoin

import application.meeting_helper
import application.word
import parsing.html.common
import parsing.html.common as html_parser
import parsing.word.docx
import server.common
import tdoc.utils
import tdoc.utils
from server.common import get_html, get_local_revisions_filename, get_local_drafts_filename, get_local_agenda_folder, get_meeting_folder, \
    get_cache_folder, get_remote_meeting_folder, get_inbox_root, unzip_files_in_zip_file, update_urls

agenda_regex = re.compile(r'.*Agenda.*[-_]([ ]|%20)*([vr])?(?P<version>\d*).*\..*')
agenda_docx_regex = re.compile(r'.*Agenda.*[-_]([ ]|%20)*([vr])?(?P<version>\d*).*\.(docx|doc|zip)')
agenda_version_regex = re.compile(r'.*Agenda.*[-_]?([ ]|%20)*([vr])(?P<version>\d*).*\..*')
agenda_draft_docx_regex = re.compile(r'.*Agenda.*[-_]([ ]|%20)*([vr])?(?P<version>\d*).*\.(docx|doc|zip)')
folder_ftp_names_regex = re.compile(r'[\d-]+[ ]+.*[ ]+<DIR>[ ]+(.*[uU][pP][dD][aA][tT][eE].*)')


# tdoc_url = 'https://portal.3gpp.org/ngppapp/DownloadTDoc.aspx?contributionUid=S2-2202451'
# Then, search for javascript: window.location.href='https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_150E_Electronic_2022-04/Docs/S2-2202451.zip';//]]> -> extract


update_urls()


def get_sa2_inbox_current_tdoc(searching_for_a_file=False):
    url = get_inbox_root(searching_for_a_file) + 'CurDoc.htm'
    return get_html(url)


def get_sa2_inbox_tdoc_list(open_tdocs_by_agenda_in_browser=False):
    url = get_inbox_root(searching_for_a_file=True) + 'TdocsByAgenda.htm'
    if open_tdocs_by_agenda_in_browser:
        os.startfile(url)
    # Return back cached HTML if there is an error retrieving the remote HTML
    fallback_cache = get_inbox_tdocs_list_cache_local_cache()
    online_html = get_html(url, file_to_return_if_error=fallback_cache)
    return online_html


def get_sa2_meeting_tdoc_list(meeting_folder, save_file_to=None, open_tdocs_by_agenda_in_browser=False):
    remote_folder = get_remote_meeting_folder(meeting_folder)
    url = remote_folder + 'TdocsByAgenda.htm'
    returned_html = get_html(url, file_to_return_if_error=save_file_to)

    if open_tdocs_by_agenda_in_browser:
        os.startfile(url)

    # Normal case
    if returned_html is not None:
        return returned_html

    # In some cases, the original TDocsByAgenda was removed (e.g. 136AH meeting). In this case, we have to look for a substitute
    folder_contents = get_html(remote_folder)
    parsed_folder = html_parser.parse_3gpp_http_ftp(folder_contents)
    tdocs_by_agenda_files = [file for file in parsed_folder.files if
                             ('TdocsByAgenda' in file) and (('.htm' in file) or ('.html' in file))]
    if len(tdocs_by_agenda_files) > 0:
        file_to_get = tdocs_by_agenda_files[0]
        url = remote_folder + file_to_get
        new_html = get_html(url)
        return new_html
    else:
        print('Returned TdocsByAgenda as NONE. Something went wrong when retrieving TDocsByAgenda.htm...')
        return None


def get_sa2_docs_tdoc_list(meeting_folder, save_file_to=None):
    remote_folder = get_remote_meeting_folder(meeting_folder)
    url = remote_folder + 'Docs'
    returned_html = get_html(url, file_to_return_if_error=save_file_to)

    return returned_html


def get_sa2_revisions_tdoc_list(meeting_folder, save_file_to=None):
    remote_folder = get_remote_meeting_folder(meeting_folder)
    url = remote_folder + 'INBOX/Revisions'
    returned_html = get_html(url, file_to_return_if_error=save_file_to)

    return returned_html


def get_sa2_drafts_tdoc_list(meeting_folder):
    remote_folder = get_remote_meeting_folder(meeting_folder)
    url = remote_folder + 'INBOX/DRAFTS'
    returned_html = get_html(url)

    # In this case, we also need to retrieve all sub-pages
    # TO-DO!!!!

    return returned_html


def get_tdoc(
        meeting_folder_name,
        tdoc_id,
        use_inbox=False,
        return_url=False,
        searching_for_a_file=False,
        use_email_approval_inbox=False):
    if '*' in tdoc_id:
        is_draft = True
        tdoc_id = tdoc_id.replace('*', '')
    else:
        is_draft = False

    if not tdoc.utils.is_tdoc(tdoc_id):
        if not return_url:
            return None
        else:
            return None, None
    tdoc_local_filename = get_local_filename(meeting_folder_name, tdoc_id, is_draft=is_draft)
    zip_file_url = get_remote_filename(
        meeting_folder_name,
        tdoc_id,
        use_inbox,
        searching_for_a_file,
        use_email_approval_inbox=use_email_approval_inbox,
        is_draft=is_draft)
    if not os.path.exists(tdoc_local_filename):
        # TODO: change to also FTP support
        tdoc_file = get_html(zip_file_url, cache=False)
        if tdoc_file is None:
            if use_inbox:
                # Retry without inbox
                return_value = get_tdoc(meeting_folder_name, tdoc_id, use_inbox=False)
            else:
                if not use_email_approval_inbox:
                    # Retry in INBOX folder
                    return_value = get_tdoc(
                        meeting_folder_name,
                        tdoc_id,
                        use_inbox=False,
                        use_email_approval_inbox=True)
                else:
                    # No need to retry
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


def get_inbox_tdocs_list_cache_local_cache(create_dir=True):
    cache_folder = get_cache_folder(create_dir)
    inbox_cache = os.path.join(cache_folder, 'InboxCache.html')
    return inbox_cache


def get_local_folder(meeting_folder_name, tdoc_id, create_dir=True, email_approval=False, is_draft=False):
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


def get_local_invitation_folder(meeting_folder_name, create_dir=True):
    meeting_folder = get_meeting_folder(meeting_folder_name)
    folder_name = os.path.join(meeting_folder, 'Invitation')
    if create_dir and (not os.path.exists(folder_name)):
        os.makedirs(folder_name, exist_ok=True)
    return folder_name


def get_local_tdocs_by_agenda_filename(meeting_folder_name):
    folder = get_local_agenda_folder(meeting_folder_name, create_dir=True)
    return os.path.join(folder, 'TdocsByAgenda.htm')


def get_remote_filename(
        meeting_folder_name,
        tdoc_id,
        use_inbox=False,
        searching_for_a_file=False,
        use_email_approval_inbox=False,
        is_draft=False):
    folder = get_remote_meeting_folder(meeting_folder_name, use_inbox, searching_for_a_file)

    if not use_inbox:
        # Check if this is a TDoc revision. If yes, change the folder to the revisions folder. Need to see how this
        # works during a meeting, but this is something to test in 2021 :P
        year, tdoc_number, revision = tdoc.utils.get_tdoc_year(tdoc_id, include_revision=True)
        if revision is not None:
            if not is_draft:
                folder = get_remote_meeting_revisions_folder(folder)
            else:
                folder = get_remote_meeting_drafts_folder(folder)
        else:
            if use_email_approval_inbox:
                folder += 'Inbox/'
            else:
                folder += 'Docs/'
    elif use_inbox:
        # No need to add 'Docs/'
        pass

    return folder + tdoc_id + '.zip'


def get_remote_meeting_revisions_folder(meeting_folder_ending_with_slash):
    return meeting_folder_ending_with_slash + 'Inbox/Revisions/'


def get_remote_meeting_drafts_folder(meeting_folder_ending_with_slash):
    return meeting_folder_ending_with_slash + 'Inbox/DRAFTS/'


def get_remote_agenda_folder(meeting_folder_name, use_inbox=False):
    folder = get_remote_meeting_folder(meeting_folder_name, use_inbox)
    if use_inbox:
        folder += 'Drafts/'
    else:
        folder += 'Agenda/'
    return folder


def get_remote_agenda_update_folder_for_inbox(meeting_folder_name):
    folder = get_remote_meeting_folder(meeting_folder_name, use_inbox=True)
    folder += 'Agenda_Updates/'
    return folder


def get_agenda_files(meeting_folder_name, use_inbox=False):
    url = get_remote_agenda_folder(meeting_folder_name, use_inbox=use_inbox)
    html = get_html(url)
    if html is None:
        return
    if use_inbox:
        # Starting from SA2#157, Agenda updates have been placed in the /Inbox/Agenda_Updates folder
        url_inbox_agenda_updates = get_remote_agenda_update_folder_for_inbox(meeting_folder_name)
        html_inbox_agenda_updates = get_html(url_inbox_agenda_updates)
        if html_inbox_agenda_updates is not None:
            html = html_inbox_agenda_updates
            url = url_inbox_agenda_updates
    parsed_folder = html_parser.parse_3gpp_http_ftp(html)
    agenda_files = [file for file in parsed_folder.files if agenda_regex.match(file)]
    agenda_folder = get_local_agenda_folder(meeting_folder_name)
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
            with open(local_file, 'wb') as output:
                output.write(html)
        if file_extension == '.zip':
            unzipped_files = unzip_files_in_zip_file(local_file)
            real_agenda_files.extend(unzipped_files)
        else:
            real_agenda_files.append(local_file)


def get_latest_agenda_file(agenda_files):
    if agenda_files is None:
        return None

    # Remove case when Word temporary documents are stored
    agenda_files = [file for file in agenda_files if agenda_docx_regex.match(file) and ('~$' not in file)]
    if len(agenda_files) == 0:
        return None

    draft_agenda_files = [file for file in agenda_files if agenda_draft_docx_regex.match(file)]
    if (len(draft_agenda_files) > 0) and (len(draft_agenda_files) != len(agenda_files)):
        # Non-draft agendas have priority over draft agenda files
        agenda_files = [agenda_file for agenda_file in agenda_files if agenda_file not in draft_agenda_files]

    last_agenda = max(agenda_files, key=get_agenda_file_version_number)
    print('Most recent agenda file seems to be {0}'.format(last_agenda))
    return last_agenda


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
    print('{0} is agenda version {1} for tdoc ID {2:.0f}'.format(x, agenda_version, tdoc_number))

    # Support up to 100 agenda versions. It should be OK...
    agenda_version = tdoc_number + float(agenda_version) / 100
    return agenda_version


ai_names_cache = {}


def get_last_agenda(meeting_folder):
    agenda_folder = get_local_agenda_folder(meeting_folder)
    agenda_files = [file for file in os.listdir(agenda_folder)]

    last_agenda = get_latest_agenda_file(agenda_files)

    if last_agenda is None:
        return None, None

    agenda_path = os.path.join(agenda_folder, last_agenda)
    try:
        agenda_item_descriptions = parsing.word.docx.import_agenda(agenda_path)
    except:
        agenda_item_descriptions = {}
    ai_names_cache[meeting_folder] = agenda_item_descriptions

    # Convert agenda version
    agenda_version_str = ''
    try:
        last_agenda_match = agenda_docx_regex.match(last_agenda)
        agenda_version_str = last_agenda_match.groupdict()['version']
        agenda_version_int = int(agenda_version_str)
    except ValueError as e:
        print("Could not parse agenda version: {0}. Agenda name: {1}".format(agenda_version_str, last_agenda_match.group(0)))
        traceback.print_exc()
    return agenda_path, agenda_version_int


# Begin with updated URLs
update_urls()

def get_tdocs_by_agenda_for_selected_meeting(
        meeting_folder: str,
        inbox_active=False,
        save_file_to=None,
        open_tdocs_by_agenda_in_browser=False):
    """
    Returns the HTML of a TdocsByAgenda file for a given meeting
    Args:
        meeting_folder: The meeting folder as named in the 3GPP server
        inbox_active: Whether the inbox is to be used
        save_file_to: Where to save the file to
        open_tdocs_by_agenda_in_browser: Whether to open the file in the browser

    Returns: The HTML contents (bytes)
    """
    # If the inbox is active, we need to download both and return the newest one
    html_inbox = None
    html_3gpp = None

    datetime_inbox = datetime.datetime.min
    datetime_3gpp = datetime.datetime.min

    if inbox_active:
        print('Getting TDocs by agenda from inbox')
        html_inbox = get_sa2_inbox_tdoc_list(open_tdocs_by_agenda_in_browser=open_tdocs_by_agenda_in_browser)
        datetime_inbox = parsing.html.common.tdocs_by_agenda.get_tdoc_by_agenda_date(html_inbox)

    print('Getting TDocs by agenda from server')
    html_3gpp = get_sa2_meeting_tdoc_list(meeting_folder, save_file_to=save_file_to,
                                          open_tdocs_by_agenda_in_browser=open_tdocs_by_agenda_in_browser)
    datetime_3gpp = parsing.html.common.tdocs_by_agenda.get_tdoc_by_agenda_date(html_3gpp)

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


def download_agenda_file(meeting, inbox_active=False, open_tdocs_by_agenda_in_browser=False):
    try:
        meeting_server_folder = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(
            meeting)
        local_file = get_local_tdocs_by_agenda_filename(meeting_server_folder)
        html = get_tdocs_by_agenda_for_selected_meeting(meeting_server_folder, inbox_active,
                                                        open_tdocs_by_agenda_in_browser=open_tdocs_by_agenda_in_browser)
        if html is None:
            print('Agenda file for {0} not found'.format(meeting))
            return None
        server.common.write_data_and_open_file(html, local_file)
        return local_file
    except:
        print('Could not download agenda file for {0}'.format(meeting))
        traceback.print_exc()
        return None


def download_docs_file(meeting) -> str:
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
        local_file = server.common.get_local_docs_filename(meeting_server_folder)
        html = get_sa2_docs_tdoc_list(meeting_server_folder, save_file_to=local_file)
        if html is None:
            print('Docs file for {0} not found'.format(meeting))
            return None
        server.common.write_data_and_open_file(html, local_file)
        return local_file
    except:
        print('Could get not docs file for {0}'.format(meeting))
        traceback.print_exc()
        return None


def download_revisions_file(meeting) -> str:
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
        server.common.write_data_and_open_file(html, local_file)
        return local_file
    except:
        print('Could get not revisions file for {0}'.format(meeting))
        traceback.print_exc()
        return None


def download_drafts_file(meeting) -> str:
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
        server.common.write_data_and_open_file(html, local_file)
        return local_file
    except:
        print('Could not get drafts file for {0}'.format(meeting))
        traceback.print_exc()
        return None
