import os.path
import pickle
from urllib.parse import urlparse
from typing import List, Tuple

import html2text

from parsing.html_specs import extract_releases_from_latest_folder, extract_spec_series_from_spec_folder, \
    extract_spec_files_from_spec_folder, extract_spec_versions_from_spec_file, cleanup_spec_name
from parsing.spec_types import SpecType
from server.common import get_html, root_folder, create_folder_if_needed, decode_string, download_file_to_location, \
    unzip_files_in_zip_file
import pandas as pd

specs_url = 'https://www.3gpp.org/ftp/Specs/latest'

# Specification page, e.g., https://www.3gpp.org/DynaReport/23501.htm
spec_page = 'https://www.3gpp.org/DynaReport/{0}.htm'

# Specification archive page, e.g., https://www.3gpp.org/ftp/Specs/archive/24_series/24.011
spec_archive_page = 'https://www.3gpp.org/ftp/Specs/archive/{0}_series/{1}'


def get_html_page_and_save_cache(url, cache, cache_file, cache_as_markup):
    html = get_html(url)
    if cache:
        if not cache_as_markup:
            print('Caching HTML: {0}'.format(cache_file))
            with open(cache_file, 'wb') as file:
                file.write(html)
            # If HTML is to be returned, bytes as-is are returned
            return html
        else:
            print('Caching Markup: {0}'.format(cache_file))
            html_decoded = decode_string(html, "cache_file".format(html))
            h = html2text.HTML2Text()
            h.ignore_links = False
            html_markup = h.handle(html_decoded)
            with open(cache_file, 'w', encoding='utf-8') as file:
                file.write(html_markup)
            # If markup is to be returned, string is returned
            return html_markup


def get_markup_from_cache(cache_file):
    print('Loading {0}'.format(cache_file))
    with open(cache_file, mode='r', encoding='utf-8') as file:
        markup = file.read()
    return markup


def get_specs_page(cache=False):
    cache_file = os.path.join(get_specs_cache_folder(), 'latest.md')
    if cache and os.path.exists(cache_file):
        markup = get_markup_from_cache(cache_file)
    else:
        markup = get_html_page_and_save_cache(specs_url, cache, cache_file, cache_as_markup=True)
    return markup


def get_release_folder(release_url, release_number, cache=False):
    cache_file = os.path.join(get_specs_cache_folder(), 'Rel_{0}.md'.format(release_number))
    if cache and os.path.exists(cache_file):
        markup = get_markup_from_cache(cache_file)
    else:
        markup = get_html_page_and_save_cache(release_url, cache, cache_file, cache_as_markup=True)
    return markup


def get_series_folder(series_url, release_number, series_number, cache=False):
    cache_file = os.path.join(
        get_specs_cache_folder(),
        'Specs_{0}_series_Rel_{1}.md'.format(series_number, release_number))
    if cache and os.path.exists(cache_file):
        markup = get_markup_from_cache(cache_file)
    else:
        markup = get_html_page_and_save_cache(series_url, cache, cache_file, cache_as_markup=True)
    return markup


def get_spec_remote_folder(spec_number, cache=False):
    # Clean up the dot as we do not use it as part of the file name
    spec_number = cleanup_spec_name(spec_number)

    cache_file = os.path.join(get_specs_cache_folder(), '{0}.md'.format(spec_number))
    if cache and os.path.exists(cache_file):
        markup = get_markup_from_cache(cache_file)
    else:
        markup = get_html_page_and_save_cache(spec_page.format(spec_number), cache, cache_file, cache_as_markup=True)
    return markup


def get_specs(cache=True) -> Tuple[pd.DataFrame, dict]:
    """
    Retrieves information related to the latest 3GPP specs (per Release) from the 3GPP server or a local cache.
    Args:
        cache: Whether caching is desired.

    Returns:
        A DataFrame containing the specification information from the 3GPP specification repository at
        https://www.3gpp.org/ftp/Specs/latest and and a dictionary containing specification metadata, e.g.
        the specification title, obtained from scraping all of the https://www.3gpp.org/DynaReport/{spec_name}.htm
        pages in the 3GPP server.

    """
    specs_df_cache_file = os.path.join(get_specs_cache_folder(), '_specs.pickle')

    # Load specs data from cache file
    if cache and os.path.exists(specs_df_cache_file):
        with open(specs_df_cache_file, "rb") as f:
            print('Loading spec cache from {0}'.format(specs_df_cache_file))
            specs_df, spec_metadata = pickle.load(f)
        return specs_df, spec_metadata

    html_latest_specs_bytes = get_specs_page(cache=cache)

    releases_data = extract_releases_from_latest_folder(
        html_latest_specs_bytes,
        base_url=specs_url)

    # For each release, extract data
    all_specs_data = []
    for release_data in releases_data:
        markup_release_data = get_release_folder(
            release_data.release_url,
            release_data.release,
            cache=cache)
        series_data_for_release = extract_spec_series_from_spec_folder(
            markup_release_data,
            release=release_data.release,
            base_url=release_data.release_url)

        # For each spec. series in each release, extract data
        for series_data in series_data_for_release:
            markup_series_data = get_series_folder(
                series_data.series_url,
                series_number=series_data.series,
                release_number=series_data.release,
                cache=cache)
            specs_data_for_series = extract_spec_files_from_spec_folder(
                markup_series_data,
                release=series_data.release,
                series=series_data.series,
                base_url=series_data.series_url)
            all_specs_data.extend(specs_data_for_series)

    specs_df = pd.DataFrame(all_specs_data)

    # Set TS/TR number as index
    specs_df.set_index("spec", inplace=True)

    unique_specs = list(specs_df.index.unique())
    unique_specs.sort()
    spec_metadata = {}
    for spec_to_download in unique_specs:
        spec_page_markup = get_spec_remote_folder(spec_to_download, cache=cache)
        spec_data = extract_spec_versions_from_spec_file(spec_page_markup)
        spec_key = spec_data.spec[0:2] + '.' + spec_data.spec[2:]
        spec_metadata[spec_key] = spec_data

    apply_spec_metadata_to_dataframe(specs_df, spec_metadata)

    if cache:
        with open(specs_df_cache_file, "wb") as f:
            print('Storing spec cache in {0}'.format(specs_df_cache_file))
            pickle.dump([specs_df, spec_metadata], f)

    return specs_df, spec_metadata


def get_specs_cache_folder(create_dir=True):
    folder_name = os.path.expanduser(os.path.join('~', root_folder, 'specs', 'server_cache'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name


def get_specs_folder(create_dir=True, spec_id=None):
    """
    Returns the folder where the specs are stored
    Args:
        create_dir: Whether the folder should be created if it does not exist.
        spec_id: If specified, the specification number (subfolder)

    Returns: The folder where the specs are stored

    """
    if spec_id is None:
        folder_name: str = os.path.expanduser(os.path.join('~', root_folder, 'specs'))
    else:
        folder_name: str = os.path.expanduser(os.path.join('~', root_folder, 'specs', spec_id))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name


def apply_spec_metadata_to_dataframe(specs_df, spec_metadata):
    specs_df['title'] = ''
    specs_df['responsible_group'] = ''
    specs_df['type'] = ''

    specs_list = specs_df.index.unique()
    for idx in specs_list:
        if idx in spec_metadata:
            specs_df.at[idx, 'title'] = spec_metadata[idx].title
            specs_df.at[idx, 'responsible_group'] = spec_metadata[idx].responsible_group
            if spec_metadata[idx].type == SpecType.TS:
                specs_df.at[idx, 'type'] = 'TS'
            elif spec_metadata[idx].type == SpecType.TR:
                specs_df.at[idx, 'type'] = 'TR'

    specs_df['search_column'] = specs_df.index + specs_df['title']
    specs_df['search_column'] = specs_df['search_column'].str.lower()


def file_version_to_version(file_version: str) -> str:
    """
    Converts the file version of a 3GPP spec, e.g., a00 to a version number, e.g., 10.0.0.
    Args:
        file_version: A three-letter string containing the file version.

    Returns: The specification version number

    """

    def letter_to_number(character: str) -> str:
        if character.isdigit():
            number = character
        else:
            number = '{0}'.format(ord(character) - ord('a') + 10)
        return number

    # File version must be three characters, e.g., a10, 821
    major_version = letter_to_number(file_version[0])
    middle_version = letter_to_number(file_version[1])
    minor_version = letter_to_number(file_version[2])
    version_number = '{0}.{1}.{2}'.format(major_version, middle_version, minor_version)
    return version_number


def version_to_file_version(version: str) -> str:
    """
    Converts the  version of a 3GPP spec, e.g., 10.0.0 to a file version number, e.g., a00.
    Args:
        version: The specification version number

    Returns: A three-letter string containing the file version.

    """
    split_version = [int(i) for i in version.split('.')]

    def number_to_letter(number: int) -> str:
        if number < 10:
            letter = '{0}'.format(number)
        else:
            letter = chr(ord('a') + number - 10)
        return letter

    # File version must be three characters, e.g., a10, 821
    first_letter = number_to_letter(split_version[0])
    second_letter = number_to_letter(split_version[1])
    third_letter = number_to_letter(split_version[2])
    return '{0}{1}{2}'.format(first_letter, second_letter, third_letter)


def download_spec_if_needed(spec_number: str, file_url: str) -> List[str]:
    """
    Downloads a given specification from said URL. The file is stored in a cache folder and only
    (re-)downloaded if needed.
    Args:
        spec_number: The specification number (with or without dots), e.g. 23.501. Needed to correctly place
        the downloaded file(s) in the cache folder
        file_url: The URL of the spec file (zip file typically) to download

    Returns:
        list[str]: Absolute path of the downloaded (and unzipped) files
    """
    spec_number = cleanup_spec_name(spec_number)
    spec_number = '{0}.{1}'.format(spec_number[0:2], spec_number[2:])
    local_folder = get_specs_folder(spec_id=spec_number)
    filename = os.path.basename(urlparse(file_url).path)
    local_filename = os.path.join(local_folder, filename)

    if not os.path.exists(local_filename):
        download_file_to_location(file_url, local_filename)
    files_in_zip = unzip_files_in_zip_file(local_filename)
    return files_in_zip


def get_url_for_spec_page(spec_number: str) -> str:
    """
    Returns the 3GPP specification page, e.g., https://www.3gpp.org/DynaReport/23501.htm
    Args:
        spec_number: The specification number. Either with dot or without

    Returns: The URL of the specification page

    """
    spec_number = cleanup_spec_name(spec_number)
    return spec_page.format(spec_number)


def get_archive_page_for_spec(spec_number_with_dot: str) -> Tuple[str, str]:
    """
    Returns the archive URL, for a given spec.
    Args:
        spec_number_with_dot: The specification number including a dot, e.g., 24.011

    Returns:
        The specification URL, e.g. https://www.3gpp.org/ftp/Specs/archive/24_series/24.011
        The series number
    """
    # https://www.3gpp.org/ftp/Specs/archive/24_series/24.011
    series_number = int(spec_number_with_dot.split('.')[0])
    series_number = '{:02d}'.format(series_number)
    spec_archive_url = 'https://www.3gpp.org/ftp/Specs/archive/{0}_series/{1}'.format(
        series_number,
        spec_number_with_dot)
    return spec_archive_url, series_number


def get_spec_archive_remote_folder(spec_number_with_dot, cache=False) -> Tuple[str, str, str]:
    """
    For a given specification, retrieves the 3GPP spec archive page,
    e.g., https://www.3gpp.org/ftp/Specs/archive/23_series/23.206
    Args:
        spec_number_with_dot: The specification number including the dot, e.g., 23.206
        cache: Whether the file should be cached or not

    Returns:
        A string tuple containing: the markup-converted text of the page remote URL of the page,
        the specs series number.
    """
    # Clean up the dot as we do not use it as part of the file name
    spec_number = cleanup_spec_name(spec_number_with_dot)

    archive_page_url, series_number = get_archive_page_for_spec(spec_number_with_dot)
    cache_file = os.path.join(get_specs_cache_folder(), 'archive_{0}.md'.format(spec_number))
    if cache and os.path.exists(cache_file):
        markup = get_markup_from_cache(cache_file)
    else:
        markup = get_html_page_and_save_cache(
            archive_page_url,
            cache,
            cache_file,
            cache_as_markup=True)
    return markup, archive_page_url, series_number
