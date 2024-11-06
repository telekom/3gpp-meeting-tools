import concurrent.futures
import os.path
import pickle
import re
import shutil
from urllib.parse import urlparse
from typing import List, Tuple, Dict, NamedTuple

import html2text
from pandas import DataFrame

import utils.local_cache
from parsing.html.specs import extract_releases_from_latest_folder, extract_spec_series_from_spec_folder, \
    extract_spec_files_from_spec_folder, extract_spec_versions_from_spec_file, cleanup_spec_name
from parsing.spec_types import SpecType, SpecVersionMapping, SpecSeries, SpecFile
from server.common import decode_string, download_file_to_location
from application.zip_files import unzip_files_in_zip_file
from server.connection import get_remote_file, HttpRequestTimeout
from utils.local_cache import create_folder_if_needed, file_exists, get_specs_cache_folder
from config.cache import CacheConfig
import pandas as pd

specs_url = 'https://www.3gpp.org/ftp/Specs/latest'

# Specification page, e.g., https://www.3gpp.org/DynaReport/23501.htm
# spec_page = 'https://www.3gpp.org/DynaReport/{0}.htm'
# Changed to https://www.3gpp.org/dynareport?code=29505.htm
spec_page = 'https://www.3gpp.org/dynareport?code={0}.htm'

# Specification archive page, e.g., https://www.3gpp.org/ftp/Specs/archive/24_series/24.011
spec_archive_page = 'https://www.3gpp.org/ftp/Specs/archive/{0}_series/{1}'

# Specification CRs page, e.g., https://portal.3gpp.org/ChangeRequests.aspx?q=1&specnumber=23.501
# For a specific WI would be e.g. https://portal.3gpp.org/ChangeRequests.aspx?q=1&specnumber=23.501&release=all&workitem=970025
spec_crs_page = 'https://portal.3gpp.org/ChangeRequests.aspx?q=1&specnumber={0}&release=all'

# Some specs may be so new that there is not an entry in the "latest" page yet
drafts_page = 'https://www.3gpp.org/ftp/Specs/latest-drafts'

# Higher timeout value for specs
timeout_values = HttpRequestTimeout(3.05, 25)


def get_html_page_and_save_cache(
        url: str,
        cache: bool,
        cache_file: str,
        cache_as_markup: bool) -> str | None:
    """

    Args:
        url: The URL of the page to retrieve
        cache: Whether the retrieved data should be written to a file
        cache_file: If yes, the file location
        cache_as_markup: Whether the retrieved HTML data should first be converted to Markup

    Returns:
        The retrieved data, either in HTML or Markup format
    """
    html = get_remote_file(url, timeout=timeout_values, cache=False)
    if html is None:
        print(f'Could NOT retrieve specs file for {url}')
        return None
    if cache_as_markup:
        html_decoded = decode_string(html, "cache_file".format(html))
        h = html2text.HTML2Text()
        h.ignore_links = False
        output_data = h.handle(html_decoded)

        # Make file smaller
        output_data = cleanup_spec_markup_file(in_markup=output_data, log_str=url)
    else:
        output_data = html

    if cache:
        if not cache_as_markup:
            print('Caching HTML: {0}'.format(cache_file))
            with open(cache_file, 'wb') as file:
                file.write(output_data)
        else:
            print('Caching Markup: {0}'.format(cache_file))
            with open(cache_file, 'w', encoding='utf-8') as file:
                file.write(output_data)

    # If HTML is to be returned, bytes as-is are returned
    # If markup is to be returned, string is returned
    return output_data


def cleanup_spec_markup_file(in_markup: str, log_str: str) -> str:
    """
    General cleanup of a markup file
    Args:
        log_str: Whether to print a log message
        in_markup: the markup text (input)

    Returns: Markup text after cleanup (output)

    """
    cleanup_markup = re.sub(r'\!\[\]\(images/.*\)', '', in_markup, flags=re.M)
    cleanup_markup = re.sub(r'\[[ ]*##LOC\[[\w]+]##[ ]*]\(javascript:void\\\(0\\\);\)', '', cleanup_markup, flags=re.M)
    cleanup_markup = (cleanup_markup
                      .replace(' "Click to show meeting details"', '')
                      .replace(' "Click to show meeting details"', '')
                      .replace('"Click to download this version"', '')
                      .replace('![icon](/ftp/geticon.axd?file=.zip)', '')
                      )

    markup_clenup_percentage = len(cleanup_markup) / len(in_markup) * 100
    print(f'Cleaning up markup file {log_str}. IN: {len(in_markup)}, OUT: {len(cleanup_markup)}, '
          f'{markup_clenup_percentage:.2f}%')
    return cleanup_markup


def get_markup_file(file_url: str, cache: bool, cache_file: str, force_download=False) -> str:
    """
    Downloads a given file and returns a markdown version of it. Can optionally use a file cache
    Args:
        file_url: The URL to retrieve
        cache: Whether to cache
        cache_file: If caching is used, the file path (i.e. file name) where to store the cached data
        force_download: Whether regardless of the cache parameter, the file should be downloaded
        (e.g. for cache updates)

    Returns:
        The markdown-converted HTML file
    """
    this_file_exists = os.path.exists(cache_file)
    if cache and this_file_exists and (not force_download):
        print('Loading {0}'.format(cache_file))
        with open(cache_file, mode='r', encoding='utf-8') as file:
            markup = file.read()
    else:
        markup = get_html_page_and_save_cache(file_url, cache, cache_file, cache_as_markup=True)
    if markup is None:
        print('Markup file at {0} could not be retrieved: cache={1}, file exists: {2}'.format(
            cache_file,
            cache,
            this_file_exists))

    return markup


def get_latest_specs_page(cache=False):
    cache_file = os.path.join(get_specs_cache_folder(), 'latest.md')
    markup = get_markup_file(specs_url, cache, cache_file)
    return markup


def get_release_folder_page(release_url, release_number, cache=False):
    cache_file = os.path.join(get_specs_cache_folder(), 'Rel_{0}.md'.format(release_number))
    markup = get_markup_file(release_url, cache, cache_file)
    return markup


def get_drafts_folder_page(cache=False):
    cache_file = os.path.join(get_specs_cache_folder(), 'drafts.md')
    markup = get_markup_file(drafts_page, cache, cache_file)
    return markup


def get_series_folder_page(series_url, release_number, series_number, cache=False):
    cache_file = os.path.join(
        get_specs_cache_folder(),
        'Specs_{0}_series_Rel_{1}.md'.format(series_number, release_number))
    markup = get_markup_file(series_url, cache, cache_file)
    return markup


def get_spec_page(spec_number: str, cache=False, force_download=False):
    # e.g. https://www.3gpp.org/DynaReport/23501.htm
    # Clean up the dot as we do not use it as part of the file name
    spec_number = cleanup_spec_name(spec_number)
    cache_file = os.path.join(get_specs_cache_folder(), '{0}.md'.format(spec_number))
    markup = get_markup_file(spec_page.format(spec_number), cache, cache_file, force_download=force_download)
    return markup


# Moved outside the function so that the last cached files can be stored in-memory between function calls. This is
# useful when reloading a single spec file
last_spec_metadata: dict[str, SpecVersionMapping] = {}

# Contains the last loaded specifications dataframe
last_specs_df: DataFrame | None = None


def get_specs(
        cache=True,
        check_for_new_specs=False,
        override_pickle_cache=False,
        load_only_spec_list: list[str] | None = None) -> Tuple[pd.DataFrame, dict[str, SpecVersionMapping]]:
    """
    Retrieves information related to the latest 3GPP specs (per Release) from the 3GPP server or a local cache.
    Args:
        override_pickle_cache: Whether the HTML/Markup cache should be used but the pickle file ignored (e.g. if an
            updated HTML file was loaded
        check_for_new_specs: Whether the cache should be updated with newly-found specs
        cache: Whether caching is desired. If yes, if existing, a cache file will be read. The cache file contains
        the last retrieved spec data
        load_only_spec_list: If the list is not empty, it contains a list of specifications that should be re-loaded. Otherwise,
            all specifications will be reloaded

    Returns:
        A DataFrame containing the specification information from the 3GPP specification repository at
        https://www.3gpp.org/ftp/Specs/latest and and a dictionary containing specification metadata, e.g.
        the specification title, obtained from scraping all of the https://www.3gpp.org/DynaReport/{spec_name}.htm
        pages in the 3GPP server.
        Also, metadata containing title and other information for the related specifications

    """
    if load_only_spec_list is None:
        load_only_spec_list = []

    try:
        current_file_directory = os.path.dirname(os.path.abspath(__file__))
        cache_file_target_folder = utils.local_cache.get_specs_cache_folder()
        file_name = 'server_cache.zip'
        source_cache_file = os.path.join(current_file_directory, file_name)
        target_cache_file = os.path.join(cache_file_target_folder, file_name)
        if not file_exists(target_cache_file) and file_exists(source_cache_file):
            print(f'Copying spec cache file to {target_cache_file}')
            utils.local_cache.create_folder_if_needed(cache_file_target_folder, create_dir=True)
            shutil.copyfile(source_cache_file, target_cache_file)
            unzipped_files = unzip_files_in_zip_file(target_cache_file)
            print(f'Unzipped {len(unzipped_files)} spec files to {cache_file_target_folder}')
    except Exception as e:
        print(f'Could not copy spec cache file: {e}')

    global last_specs_df, last_spec_metadata
    print('Loading specs: cache={0}, check for new specs={1}, override pickle cache={2}, load only={3}'.format(
        cache,
        check_for_new_specs,
        override_pickle_cache,
        load_only_spec_list))
    specs_df_cache_file = os.path.join(get_specs_cache_folder(), '_specs.pickle')

    # Load specs data from cache file
    if not override_pickle_cache:
        if cache and (not check_for_new_specs) and os.path.exists(specs_df_cache_file):
            with open(specs_df_cache_file, "rb") as f:
                print('Loading spec cache from {0}'.format(specs_df_cache_file))
                specs_df, last_spec_metadata = pickle.load(f)
            last_specs_df = specs_df
            return specs_df, last_spec_metadata

    if cache and (not check_for_new_specs):
        latest_and_series_cache = True
        print('Use of cached files enabled for all retrievals')
    else:
        latest_and_series_cache = False
        print('Checking for new specs. Disabling file cache for some retrievals')

    if len(load_only_spec_list) > 0:
        # Selectively reloading a known spec does not require to require all of the spec series data
        print('Skipping loading series data (loading only {0})'.format(load_only_spec_list))
        specs_df = last_specs_df
    else:
        print('Loading series data')

        # Get HTML page: https://www.3gpp.org/ftp/Specs/latest
        markup_latest_specs = get_latest_specs_page(cache=latest_and_series_cache)
        releases_data = extract_releases_from_latest_folder(
            markup_latest_specs,
            base_url=specs_url)

        # Retrieve information for all 3GPP releases
        series_data_per_release = []
        for release_data in releases_data:
            # For each Release, get the corresponding page, e.g. https://www.3gpp.org/ftp/Specs/latest/Rel-10
            markup_release_data = get_release_folder_page(
                release_data.release_url,
                release_data.release,
                cache=latest_and_series_cache)
            series_data_for_release = extract_spec_series_from_spec_folder(
                markup_release_data,
                release=release_data.release,
                base_url=release_data.release_url)
            series_data_per_release.append(series_data_for_release)

        all_specs_data = []

        def task_per_series(series_to_process: SpecSeries) -> List[SpecFile]:
            markup_series_data = get_series_folder_page(
                series_data.series_url,
                series_number=series_data.series,
                release_number=series_data.release,
                cache=latest_and_series_cache)
            specs_data_for_series = extract_spec_files_from_spec_folder(
                markup_series_data,
                release=series_data.release,
                series=series_data.series,
                base_url=series_data.series_url)
            return specs_data_for_series

        all_spec_series: List[SpecSeries] = []
        for series_data_for_release in series_data_per_release:
            all_spec_series.extend(series_data_for_release)

        # For each spec. series in each release, extract data
        # For each release, extract data, e.g. from https://www.3gpp.org/ftp/Specs/latest/Rel-10/23_series
        # for series_data in all_spec_series:
        #     all_specs_data.extend(task_per_series(series_data))

        # See https://docs.python.org/3/library/concurrent.futures.html
        # For each spec. series in each release, extract data
        # For each release, extract data, e.g. from https://www.3gpp.org/ftp/Specs/latest/Rel-10/23_series
        with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
            future_to_spec = {executor.submit(
                task_per_series,
                series_data): series_data for series_data in all_spec_series}
            for future in concurrent.futures.as_completed(future_to_spec):
                series_data = future_to_spec[future]
                try:
                    specs_data_for_series = future.result()
                    all_specs_data.extend(specs_data_for_series)
                except Exception as exc:
                    print('%r generated an exception: %s' % (series_data, exc))

        # Retrieve Drafts folder
        # Get drafts page: https://www.3gpp.org/ftp/Specs/latest-drafts
        markup_draft_specs = get_drafts_folder_page(cache=latest_and_series_cache)
        specs_data_for_drafts = extract_spec_files_from_spec_folder(
            markup_draft_specs,
            release='Draft',
            series=None,
            base_url=drafts_page,
            auto_fill=True)
        all_specs_data.extend(specs_data_for_drafts)

        # Convert specs data into DataFrame
        specs_df = pd.DataFrame(all_specs_data)

        # Set TS/TR number as index
        specs_df.set_index("spec", inplace=True)

    # If only one or more specs need to be reloaded, reload only those ones
    if len(load_only_spec_list) > 0:
        unique_specs = list(set(load_only_spec_list))
        print('Will only reload specs={0}'.format(unique_specs))
    else:
        unique_specs: List[str] = list(specs_df.index.unique())
        print('Will reload all specs')
    unique_specs.sort()

    # Download each spec's page, e.g. https://www.3gpp.org/DynaReport/23501.htm (or from cache)
    class DownloadedSpecData(NamedTuple):
        spec_key: str
        spec_data: SpecVersionMapping

    def get_spec_data(spec_to_download_str: str, cache_bool: bool) -> DownloadedSpecData:
        spec_page_markup_from_server = get_spec_page(spec_to_download_str, cache=cache_bool)
        spec_data_from_markdown = extract_spec_versions_from_spec_file(spec_page_markup_from_server)
        spec_key_from_markdown = spec_data_from_markdown.spec[0:2] + '.' + spec_data_from_markdown.spec[2:]
        return DownloadedSpecData(spec_key=spec_key_from_markdown, spec_data=spec_data_from_markdown)

    # See https://docs.python.org/3/library/concurrent.futures.html
    # 10 Executor threads because the spec. page is quite slow to load
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
        future_to_spec = {executor.submit(
            get_spec_data,
            spec_to_download_str,
            cache): spec_to_download_str for spec_to_download_str in unique_specs}
        for future in concurrent.futures.as_completed(future_to_spec):
            spec_to_download = future_to_spec[future]
            try:
                downloaded_spec = future.result()
                last_spec_metadata[downloaded_spec.spec_key] = downloaded_spec.spec_data
            except Exception as exc:
                print('%r generated an exception: %s' % (spec_to_download, exc))

    apply_spec_metadata_to_dataframe(specs_df, last_spec_metadata)

    if cache:
        with open(specs_df_cache_file, "wb") as f:
            print('Storing spec cache in {0}'.format(specs_df_cache_file))
            pickle.dump([specs_df, last_spec_metadata], f)

    last_specs_df = specs_df
    return specs_df, last_spec_metadata


def get_specs_folder(create_dir=True, spec_id=None):
    """
    Returns the folder where the specs are stored
    Args:
        create_dir: Whether the folder should be created if it does not exist.
        spec_id: If specified, the specification number (subfolder)

    Returns: The folder where the specs are stored

    """
    if spec_id is None:
        folder_name: str = os.path.expanduser(os.path.join(
            CacheConfig.user_folder,
            CacheConfig.root_folder,
            'specs'))
    else:
        folder_name: str = os.path.expanduser(os.path.join(
            CacheConfig.user_folder,
            CacheConfig.root_folder,
            'specs',
            spec_id))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name


def apply_spec_metadata_to_dataframe(specs_df: pd.DataFrame, spec_metadata: Dict[str, SpecVersionMapping]):
    """
    Applies metadata information to the sepecifications DataFrame
    Args:
        specs_df: The DataFrame containing the specification data
        spec_metadata: A dictionary with specification metadata extracted from parsing the specification pages
    """
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


class SpecVersionMetadata(NamedTuple):
    major_version: str
    middle_version: str
    minor_version: str

    @property
    def release(self) -> str:
        return self.major_version


def file_version_to_version_metadata(file_version: str) -> SpecVersionMetadata:
    version = file_version_to_version(file_version)
    version_split = version.split('.')
    return SpecVersionMetadata(
        major_version=version_split[0],
        middle_version=version_split[1],
        minor_version=version_split[2])


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


def download_spec_if_needed(
        spec_number: str,
        file_url: str,
        return_only_target_local_filename: bool = False) -> List[str] | str:
    """
    Downloads a given specification from said URL. The file is stored in a cache folder and only
    (re-)downloaded if needed.
    Args:
        return_only_target_local_filename: If this parameter is set, it just returns the target spec file to download
        spec_number: The specification number (with or without dots), e.g. 23.501. Needed to correctly place
        the downloaded file(s) in the cache folder
        file_url: The URL of the spec file (zip file typically) to download

    Returns:
        list[str]: Absolute path of the downloaded (and unzipped) files. Only the local file (zip) if
        return_only_target_local_filename is set to True
    """
    spec_number = cleanup_spec_name(spec_number)
    spec_number = '{0}.{1}'.format(spec_number[0:2], spec_number[2:])
    local_folder = get_specs_folder(spec_id=spec_number)
    filename = os.path.basename(urlparse(file_url).path)
    local_filename = os.path.join(local_folder, filename)

    if return_only_target_local_filename:
        return local_filename

    if not file_exists(local_filename):
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


def get_url_for_crs_page(spec_number: str, wi_uid: str = None) -> str:
    """
    Returns the 3GPP CRs page for a given specification page, e.g., https://portal.3gpp.org/ChangeRequests.aspx?q=1&specnumber=23.501
    Args:
        wi_uid: Optional 3GPP WI number
        spec_number: The specification number. Either with dot or without

    Returns: The URL of the specification page

    """
    spec_number = cleanup_spec_name(spec_number)
    spec_number = '{0}.{1}'.format(spec_number[0:2], spec_number[2:])
    return_url = spec_crs_page.format(spec_number)
    if wi_uid is not None:
        return_url = return_url + '&workitem={0}'.format(wi_uid)
    return return_url


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


def get_spec_archive_remote_folder(
        spec_number_with_dot,
        cache=False,
        force_download=False) -> Tuple[str, str, str]:
    """
    For a given specification, retrieves the 3GPP spec archive page,
    e.g., https://www.3gpp.org/ftp/Specs/archive/23_series/23.206
    Args:
        force_download: Whether to download regardless of cache status
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
    markup = get_markup_file(archive_page_url, cache, cache_file, force_download=force_download)
    return markup, archive_page_url, series_number
