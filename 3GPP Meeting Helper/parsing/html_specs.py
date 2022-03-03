import html2text
import pandas as pd
import re
import traceback
import collections

SpecReleases = collections.namedtuple('SpecReleases', 'folder release base_url release_url')
SpecSeries = collections.namedtuple('SpecSeries', 'folder series release base_url series_url')
SpecFile = collections.namedtuple('SpecFile', 'file spec version series release base_url spec_url')
SpecVersionMapping = collections.namedtuple('SpecVersionMapping', 'spec title version_mapping')


def extract_releases_from_latest_folder(latest_specs_page_text, base_url):
    """
    Extracts the 3GPP release information from the HTML of https://www.3gpp.org/ftp/Specs/latest
    Args:
        latest_specs_page_text (str): Markup content from which to extract the releases information
        base_url (str): URL where this HTML was extracted. Note that it should NOT end in '/'

    Returns:
        list(SpecReleases): List of releases found in the 3GPP site
    """
    releases = [SpecReleases(m.group(0), m.group(1), base_url, base_url + '/' + m.group(0)) for m in
                re.finditer(r'Rel-([\d]{1,2})', latest_specs_page_text)]
    releases = list(set(releases))

    return releases


def extract_spec_series_from_spec_folder(series_specs_page_text, base_url=None, release=None):
    """
    Extracts the 3GPP series information from the HTML of a release, e.g., https://www.3gpp.org/ftp/Specs/latest/Rel-17
    Args:
        series_specs_page_text (str): Markup content from which to extract the releases information
        base_url (str): URL where this HTML was extracted. Note that it should NOT end in '/'
        release (str): The release NUMBER this extraction relates to

    Returns:
        list(SpecSeries): List of series found in the 3GPP site
    """
    # SpecSeries = collections.namedtuple('SpecSeries', 'folder series release base_url series_url')
    series = [SpecSeries(m.group(0), m.group(1), release, base_url, base_url + '/' + m.group(0))
              for m in re.finditer(r'([\d]{1,2})_series', series_specs_page_text)]
    series = list(set(series))

    return series


# Need to account to TR numbers and old release numbers. Examples:
# 23700-07-h00.zip, 23003-aa0.zip, 23034-800.zip
spec_zipfile_regex = re.compile(r'([\d]{5}(-[\d]{2})?)-([\d\w]*).zip')


def extract_spec_files_from_spec_folder(specs_page_markup, base_url, release, series):
    """
    Extracts the 3GPP series information from the HTML of a release, e.g., https://www.3gpp.org/ftp/Specs/latest/Rel-18/23_series
    Args:
        specs_page_markup (str): Markup content from which to extract the releases information
        base_url (str): URL where this HTML was extracted. Note that it should NOT end in '/'
        release (str): The release number (not folder) this extraction relates to
        series (str): The series number (not folder) this extraction relates to

    Returns:
        list(SpecFile): List of specs found in the 3GPP site
    """
    # Need to account to TR numbers and old release numbers. Examples:
    # 23700-07-h00.zip, 23003-aa0.zip, 23034-800.zip
    specs = [
        SpecFile(m.group(0), '{0}.{1}'.format(m.group(1)[0:2], m.group(1)[2:]), m.group(3), series, release, base_url,
                 base_url + '/' + m.group(0)) for m in spec_zipfile_regex.finditer(specs_page_markup)]
    specs = list(set(specs))

    return specs


# e.g., [1.0.0](https://www.3gpp.org/ftp/Specs/archive/21_series/21.101/21101-100.zip
spec_version_regex = re.compile(r'\[([\d\.]+)]\(https://.*/(.*\.zip)')


def extract_spec_versions_from_spec_file(spec_page_markup):
    """
    Extracts the 3GPP series information from the HTML of a release, e.g., https://www.3gpp.org/ftp/Specs/latest/Rel-18/23_series
    Args:
        spec_page_markup (str): Markup content from which to extract the releases information

    Returns:
        list(SpecFile): List of specs found in the 3GPP site
    """

    spec_number = extract_text_between_delimiters(
        spec_page_markup,
        start_delimiter="Specification #: ",
        stop_delimiter="  * General").replace('.', '')
    spec_title = extract_text_between_delimiters(
        spec_page_markup,
        start_delimiter="Title: |  ",
        stop_delimiter="Status: |  ")
    print('Spec {0}: {1}'.format(spec_title, spec_number))

    # [16.0.0](https://www.3gpp.org/ftp//Specs/archive/21_series/21.101/21101-g00.zip
    # m Group 1: 16.0.0
    # m Group 2: 21101-g00.zip
    # m_file Group 1: 21101
    # m_file Group 2: g00
    spec_versions = {m_file.group(2): m.group(1) for m in spec_version_regex.finditer(spec_page_markup)
                     if (m_file := spec_zipfile_regex.match(m.group(2))) is not None}

    spec_mapping = SpecVersionMapping(spec_number, spec_title, spec_versions)
    print(spec_mapping)

    return spec_mapping


def extract_text_between_delimiters(text, start_delimiter, stop_delimiter):
    start = text.find(start_delimiter) + len(start_delimiter)
    stop = text.find(stop_delimiter)
    extracted_text = text[start:stop].replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
    return extracted_text
