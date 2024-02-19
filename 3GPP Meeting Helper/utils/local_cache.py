import os
import traceback

import html2text

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


def get_local_agenda_folder(meeting_folder_name: str, create_dir=True) -> str:
    """
    Get the path of the folder where the agenda file is cached
    Args:
        meeting_folder_name: The meeting name
        create_dir: Whether the folder should be created if it does not exist

    Returns: Path

    """
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


def get_meeting_list_folder(create_dir=True) -> str:
    """
    Folder where the meeting lists are saved to
    Args:
        create_dir: Whether to create the directory if it does not exist

    Returns: The path of the folder

    """
    folder_name = os.path.expanduser(os.path.join(user_folder, root_folder, 'meetings'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name


def get_sa2_root_folder_local_cache(create_dir=True):
    cache_folder = get_cache_folder(create_dir)
    inbox_cache = os.path.join(cache_folder, 'Wg2ArchCache.html')
    return inbox_cache


def convert_html_file_to_markup(
        file_path: str,
        output_path: str= None,
        ignore_links=True,
        filter_text_function=None) -> str:
    """
    Converts a HTML file to Markdown
    Args:
        output_path: Optional path where to save this file. If not, same as original with .md extension
        filter_text_function: Additional function that can be passed on to further filter the text
        file_path: The file's path
        ignore_links: Whether links should be included

    Returns: The destination file

    """
    if not os.path.exists(file_path):
        return None
    [root, ext] = os.path.splitext(file_path)
    if ext not in ['.htm', '.html']:
        return None
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            html_content = file.read()
    except:
        print('Could not open file "{0}"'.format(html_content))
        return None
    h = html2text.HTML2Text()
    h.ignore_links = ignore_links
    markdown_text = h.handle(html_content)

    if filter_text_function is not None:
        markdown_text = filter_text_function(markdown_text)

    if output_path is None:
        destination_file = os.path.join(root + '.md')
    else:
        destination_file = output_path

    try:
        with open(destination_file, 'w', encoding='utf-8') as file:
            file.write(markdown_text)
        return destination_file
    except:
        print('Could not write file "{0}"'.format(destination_file))
        traceback.print_exc()
        return None
