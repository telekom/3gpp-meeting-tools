import os
import traceback

import server.common
import utils.local_cache
from server.common import get_remote_meeting_folder, get_inbox_root
from server.connection import get_remote_file
from utils.local_cache import get_meeting_folder


def download_chairnotes_file(meeting):
    try:
        meeting_server_folder = meeting  # e.g. TSGS2_144E_Electronic
        print("Retrieving Chairman's Notes list for {0} meeting".format(meeting))
        local_file = get_local_chairnotes_filename(meeting_server_folder)
        html = get_sa2_chairnotes_list(meeting_server_folder)
        if html is None:
            print("Chairman's Notes file for {0} not found".format(meeting))
            return None
        utils.local_cache.write_data_and_open_file(html, local_file)
        return local_file
    except:
        print("Could not get Chairman's Notes file for {0}".format(meeting))
        traceback.print_exc()
        return None


def get_sa2_chairnotes_list(meeting_folder):
    remote_folder = get_remote_meeting_folder(meeting_folder)
    url = remote_folder + 'INBOX/Chair_Notes'
    returned_html = get_remote_file(url)

    return returned_html


def get_local_chairnotes_filename(meeting_folder_name):
    folder = get_local_chairnotes_folder(meeting_folder_name, create_dir=True)
    return os.path.join(folder, 'ChairNotes.htm')


def get_chairnotes(
        meeting_folder_name,
        chairnotes_df,
        use_inbox=False):

    filenames_list = list(chairnotes_df['file'])
    chairnotes_local_folder = get_local_chairnotes_folder(meeting_folder_name)
    filenames_local_path = [os.path.join(chairnotes_local_folder, e) for e in filenames_list]

    chairnotes_remote_folder = get_chairnotes_url()
    filenames_remote_url = [chairnotes_remote_folder + e for e in filenames_list]

    filenames_to_iterate = zip(filenames_local_path, filenames_remote_url)

    for t in filenames_to_iterate:
        local_filename = t[0]
        remote_url = t[1]
        if os.path.exists(local_filename):
            continue

        print('ToDo: download {0}'.format(remote_url))


def get_local_chairnotes_folder(meeting_folder_name, create_dir=True):
    local_folder_for_this_meeting = get_meeting_folder(meeting_folder_name)
    folder_name = os.path.join(local_folder_for_this_meeting, 'Agenda', 'ChairNotes')
    if create_dir and (not os.path.exists(folder_name)):
        os.makedirs(folder_name, exist_ok=True)
    return folder_name


def get_chairnotes_url(searching_for_a_file=False):
    return get_inbox_root(searching_for_a_file) + 'Inbox/' + 'Chair_Notes/'