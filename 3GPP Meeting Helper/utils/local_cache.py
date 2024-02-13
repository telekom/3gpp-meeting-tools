import os

from config.cache import user_folder, root_folder


def get_cache_folder(create_dir=False):
    folder_name = os.path.expanduser(os.path.join(user_folder, root_folder, 'cache'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name


def create_folder_if_needed(folder_name, create_dir):
    if create_dir and (not os.path.exists(folder_name)):
        os.makedirs(folder_name, exist_ok=True)


def get_local_docs_filename(meeting_folder_name):
    folder = get_local_agenda_folder(meeting_folder_name, create_dir=True)
    return os.path.join(folder, 'Docs.htm')


def get_local_revisions_filename(meeting_folder_name):
    folder = get_local_agenda_folder(meeting_folder_name, create_dir=True)
    return os.path.join(folder, 'Revisions.htm')


def get_local_drafts_filename(meeting_folder_name):
    folder = get_local_agenda_folder(meeting_folder_name, create_dir=True)
    return os.path.join(folder, 'Drafts.htm')


def get_meeting_folder(meeting_folder_name, create_dir=False):
    folder_name = os.path.join(get_cache_folder(create_dir=create_dir), meeting_folder_name)
    create_folder_if_needed(folder_name, create_dir)
    return folder_name


def get_local_agenda_folder(meeting_folder_name, create_dir=True):
    local_folder_for_this_meeting = get_meeting_folder(meeting_folder_name)
    folder_name = os.path.join(local_folder_for_this_meeting, 'Agenda')
    if create_dir and (not os.path.exists(folder_name)):
        os.makedirs(folder_name, exist_ok=True)
    return folder_name


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


def get_tmp_folder(create_dir=True):
    folder_name = os.path.expanduser(os.path.join(user_folder, root_folder, 'tmp'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name


def get_spec_folder(create_dir=True):
    folder_name = os.path.expanduser(os.path.join(user_folder, root_folder, 'specs'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name


def get_sa2_root_folder_local_cache(create_dir=True):
    cache_folder = get_cache_folder(create_dir)
    inbox_cache = os.path.join(cache_folder, 'Wg2ArchCache.html')
    return inbox_cache
