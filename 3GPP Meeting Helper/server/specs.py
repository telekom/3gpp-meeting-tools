import os.path

import html2text

from parsing.html_specs import extract_releases_from_latest_folder, extract_spec_series_from_spec_folder, \
    extract_spec_files_from_spec_folder
from server.common import get_html, root_folder, create_folder_if_needed, decode_string
import pandas as pd

specs_url = 'https://www.3gpp.org/ftp/Specs/latest'

# Specification page, e.g., https://www.3gpp.org/DynaReport/23501.htm
spec_page = 'https://www.3gpp.org/DynaReport/{0}.htm'


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
    cache_file = os.path.join(get_specs_cache_folder(), 'Specs_{0}_series_Rel_{1}.md'.format(series_number, release_number))
    if cache and os.path.exists(cache_file):
        markup = get_markup_from_cache(cache_file)
    else:
        markup = get_html_page_and_save_cache(series_url, cache, cache_file, cache_as_markup=True)
    return markup


def get_spec_remote_folder(spec_number, cache=False):
    # Clean up the dot as we do not use it as part of the file name
    spec_number = spec_number.replace('.', '')

    cache_file = os.path.join(get_specs_cache_folder(), '{0}.md'.format(spec_number))
    if cache and os.path.exists(cache_file):
        markup = get_markup_from_cache(cache_file)
    else:
        markup = get_html_page_and_save_cache(spec_page.format(spec_number), cache, cache_file, cache_as_markup=True)
    return markup


def get_specs(cache=True):
    html_latest_specs_bytes = get_specs_page(cache=cache)

    releases_data = extract_releases_from_latest_folder(
        html_latest_specs_bytes,
        base_url=specs_url)

    # For each release, extract data
    all_specs_data = []
    for release_data in releases_data:
        html_release_data_bytes = get_release_folder(
            release_data.release_url,
            release_data.release,
            cache=cache)
        series_data_for_release = extract_spec_series_from_spec_folder(
            html_release_data_bytes,
            release=release_data.release,
            base_url=release_data.release_url)

        # For each spec. series in each release, extract data
        for series_data in series_data_for_release:
            html_series_data_bytes = get_series_folder(
                series_data.series_url,
                series_number=series_data.series,
                release_number=series_data.release,
                cache=cache)
            specs_data_for_series = extract_spec_files_from_spec_folder(
                html_series_data_bytes,
                release=series_data.release,
                series=series_data.series,
                base_url=series_data.series_url)
            all_specs_data.extend(specs_data_for_series)

    specs_df = pd.DataFrame(all_specs_data)

    # Set TS/TR number as index
    specs_df.set_index("spec", inplace=True)

    unique_specs = list(specs_df.index.unique())
    unique_specs.sort()
    for spec_to_download in unique_specs:
        get_spec_remote_folder(spec_to_download, cache=cache)
    return specs_df


def get_specs_cache_folder(create_dir=True):
    folder_name = os.path.expanduser(os.path.join('~', root_folder, 'specs', 'server_cache'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name
