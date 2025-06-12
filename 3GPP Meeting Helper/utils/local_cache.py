import hashlib
import os
import pickle
import traceback
from typing import Callable, Any, NamedTuple

import html2text
from pandas import DataFrame

from config.cache import CacheConfig


def get_cache_folder(create_dir=False):
    folder_name = os.path.expanduser(os.path.join(
        CacheConfig.user_folder,
        CacheConfig.root_folder,
        'cache'))
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
    folder_name = os.path.expanduser(os.path.join(CacheConfig.user_folder, CacheConfig.root_folder, 'tmp'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name

def get_export_folder(create_dir=True):
    folder_name = os.path.expanduser(os.path.join(CacheConfig.user_folder, CacheConfig.root_folder, 'export'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name


def get_webcache_file():
    folder_name = get_tmp_folder()
    return os.path.join(folder_name, '.webcache')


def get_spec_folder(create_dir=True):
    folder_name = os.path.expanduser(os.path.join(CacheConfig.user_folder, CacheConfig.root_folder, 'specs'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name


def get_meeting_list_folder(create_dir=True) -> str:
    """
    Folder where the meeting lists are saved to
    Args:
        create_dir: Whether to create the directory if it does not exist

    Returns: The path of the folder

    """
    folder_name = os.path.expanduser(os.path.join(CacheConfig.user_folder, CacheConfig.root_folder, 'meetings'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name


def get_sa2_root_folder_local_cache(create_dir=True):
    cache_folder = get_cache_folder(create_dir)
    inbox_cache = os.path.join(cache_folder, 'Wg2ArchCache.html')
    return inbox_cache


def convert_html_file_to_markup(
        file_path: str,
        output_path: str = None,
        ignore_links=True,
        filter_text_function: Callable[[str], str] = None) -> str | None:
    """
    Converts an HTML file to Markdown
    Args:
        output_path: Optional path where to save this file. If not, same as original with .md extension
        filter_text_function: Additional function (callable) that can be passed on to further filter the text. Takes as
         input a string containing the markup text and outputs the (further) filtered text
        file_path: The file's path
        ignore_links: Whether links should be included

    Returns: The destination file, None if the conversion could not be performed

    """
    if not os.path.exists(file_path):
        return None
    [root, ext] = os.path.splitext(file_path)
    if ext not in ['.htm', '.html']:
        return None
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            html_content = file.read()
    except Exception as e:
        print(f'Could not open file "{html_content}": {e}')
        return None
    h = html2text.HTML2Text()
    h.ignore_links = ignore_links
    markdown_text = h.handle(html_content)

    if filter_text_function is not None:
        markdown_text = filter_text_function(markdown_text)

    if output_path is None:
        destination_file = str(os.path.join(root + '.md'))  # To avoid IDE warnings
    else:
        destination_file = output_path

    try:
        with open(destination_file, 'w', encoding='utf-8') as file:
            file.write(markdown_text)
        return destination_file
    except Exception as e:
        print(f'Could not write file "{destination_file}": {e}')
        traceback.print_exc()
        return None


def file_exists(local_filename: str, print_log: bool = False) -> bool:
    """
    Returns whether the file exists, and if it exists, if it is NOT of null size
    Args:
        print_log: Whether to print a log related to whether the file exists
        local_filename: The file path
    """
    local_file_exists = os.path.exists(local_filename)
    if not local_file_exists:
        if print_log:
            print(f'{local_filename} does not exist')
        return False
    try:
        local_file_size = os.path.getsize(local_filename)
        if local_file_size == 0:
            print(f'File {local_filename} is of size 0. Most probably corrupted')
            return False
    except OSError as e:
        print(f"Could not ascertain downloaded file's size: {e}")
        traceback.print_exc()
        return False

    if print_log:
        print(f'{local_filename} exists')
    return True


def get_specs_cache_folder(create_dir=True):
    folder_name = os.path.expanduser(os.path.join(
        CacheConfig.user_folder,
        CacheConfig.root_folder,
        'specs',
        'server_cache'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name


def get_work_items_cache_folder(create_dir=True):
    folder_name = os.path.expanduser(
        os.path.join(CacheConfig.user_folder, CacheConfig.root_folder, 'work_items', 'server_cache'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name


def get_tdocs_by_agenda_filename(meeting_folder_name):
    folder = get_local_agenda_folder(meeting_folder_name, create_dir=True)
    return os.path.join(folder, 'TdocsByAgenda.htm')


def get_private_server_tdocs_by_agenda_local_cache(create_dir=True):
    cache_folder = get_cache_folder(create_dir)
    cache_file = os.path.join(cache_folder, '3gpp_server_TdocsByAgenda.html')
    return cache_file

class CachedMeetingTdocData(NamedTuple):
    tdocs_df: DataFrame
    wy_hyperlinks: dict[str, str]
    hash: str


    @staticmethod
    def get_cache(tdoc_excel_path:str, excel_hash:str=None):
        cached_data: CachedMeetingTdocData|None = retrieve_pickle_cache_for_file(
            file_path=tdoc_excel_path,
            file_prefix='TDocs_3GU',
            file_hash=excel_hash
        )
        return cached_data

    def store_cache(self, tdoc_excel_path:str):
        store_pickle_cache_for_file(
            file_path=tdoc_excel_path,
            file_prefix='TDocs_3GU',
            file_hash=self.hash,
            data=self
        )

def hash_file(file_path: str, chunk_size=4096) -> str|None:
    """
    Calculates the MD5 hash of a file by reading it in chunks.

    Args:
        file_path (str): The path to the file.
        chunk_size (int): The size of chunks to read from the file (in bytes).
                          Larger chunks can be faster for large files, but
                          use more memory. 4096 or 8192 are common values.

    Returns:
        str: The 32-character hexadecimal MD5 digest of the file, or None if an error occurs.
    """
    md5_hasher = hashlib.md5()
    try:
        with open(file_path, 'rb') as f:  # Open in binary read mode ('rb')
            while True:
                chunk = f.read(chunk_size)
                if not chunk:  # End of file
                    break
                md5_hasher.update(chunk)
        return md5_hasher.hexdigest()
    except FileNotFoundError:
        print(f"Error: File not found at '{file_path}'")
        return None
    except Exception as e:
        print(f"An error occurred while hashing the file: {e}")
        return None

def store_pickle_cache_for_file(
        file_path: str,
        file_prefix:str,
        data:Any,
        file_hash:str=None):

    file_folder =  os.path.dirname(file_path)
    if file_hash is None:
        file_hash = hash_file(file_path)
    target_file = os.path.join(file_folder, f'{file_prefix}_{file_hash}.pickle')
    if not os.path.exists(target_file):
        try:
            with open(target_file, 'wb') as file:
                pickle.dump(data, file)
            print(f"Object '{data}' successfully saved to '{file}'")
        except Exception as e:
            print(f"Error saving object: {e}")

def retrieve_pickle_cache_for_file(
        file_path: str,
        file_prefix:str,
        file_hash:str)->Any:

    file_folder = os.path.dirname(file_path)
    target_file = os.path.join(file_folder, f'{file_prefix}_{file_hash}.pickle')

    if not os.path.exists(target_file):
        return None

    try:
        with open(target_file, 'rb') as file:
            loaded_object = pickle.load(file)
        print(f"Object successfully loaded from '{target_file}'")
        print(f"Type of loaded object: {type(loaded_object)}")
        return loaded_object

    except FileNotFoundError:
        print(f"Error: The file '{target_file}' was not found.")
        return None
    except pickle.UnpicklingError as e:
        print(f"Error unpickling data from '{target_file}': {e}")
    except Exception as e:
        print(f"An unexpected error occurred loading {target_file}: {e}")


