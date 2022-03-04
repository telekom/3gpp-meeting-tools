import collections
import re
# https://stackoverflow.com/questions/52623204/how-to-specify-method-return-type-list-of-what-in-python
# Edit: With the new 3.9 version of Python, you can annotate types without importing from the typing module
from typing import List

SpecReleases = collections.namedtuple('SpecReleases', 'folder release base_url release_url')
SpecSeries = collections.namedtuple('SpecSeries', 'folder series release base_url series_url')
SpecFile = collections.namedtuple('SpecFile', 'file spec version series release base_url spec_url')

# version_mapping: 16.0.0->g00, version_mapping_inv: g00->16.0.0
SpecVersionMapping = collections.namedtuple(
    'SpecVersionMapping',
    'spec title version_mapping version_mapping_inv responsible_group')


def extract_releases_from_latest_folder(latest_specs_page_text, base_url) -> List[SpecReleases]:
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
    releases = sorted(releases, key=lambda tup: int(tup.release))

    return releases


def extract_spec_series_from_spec_folder(series_specs_page_text, base_url=None, release=None) -> List[SpecSeries]:
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
    series = sorted(series, key=lambda tup: int(tup.series))

    return series


# Need to account to TR numbers and old release numbers. Examples:
# 23700-07-h00.zip, 23003-aa0.zip, 23034-800.zip
spec_zipfile_regex = re.compile(r'([\d]{5}(-[\d]{2})?)-([\d\w]*).zip')


def extract_spec_files_from_spec_folder(specs_page_markup, base_url, release, series) -> List[SpecFile]:
    """
    Extracts the 3GPP series information from the HTML of a release,
    e.g., https://www.3gpp.org/ftp/Specs/latest/Rel-18/23_series. Also works for spec archive pages, e.g.,
    https://www.3gpp.org/ftp/Specs/archive/23_series/23.206
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


def extract_spec_versions_from_spec_file(spec_page_markup) -> SpecVersionMapping:
    """
    Extracts the 3GPP series information from the HTML of a release, e.g., https://www.3gpp.org/ftp/Specs/latest/Rel-18/23_series
    Args:
        spec_page_markup (str): Markup content from which to extract the releases information

    Returns:
        list(SpecFile): List of specs found in the 3GPP site
    """

    # Without the dot: 23.501->23501
    spec_number = extract_text_between_delimiters(
        spec_page_markup,
        start_delimiter="Specification #: ",
        stop_delimiter="  * General").replace('.', '')
    spec_title = extract_text_between_delimiters(
        spec_page_markup,
        start_delimiter="Title: |  ",
        stop_delimiter="Status: |  ")
    responsible_group = extract_text_between_delimiters(
        spec_page_markup,
        start_delimiter="Primary responsible group: |  ",
        stop_delimiter="Secondary responsible groups: |  ")
    # print('Spec {0}: {1}'.format(spec_number, spec_title))

    # [16.0.0](https://www.3gpp.org/ftp//Specs/archive/21_series/21.101/21101-g00.zip
    # m Group 1: 16.0.0
    # m Group 2: 21101-g00.zip
    # m_file Group 1: 21101
    # m_file Group 3: g00
    spec_versions = {m_file.group(3): m.group(1)
                     for m in spec_version_regex.finditer(spec_page_markup)
                     if (m_file := spec_zipfile_regex.match(m.group(2))) is not None}
    spec_versions_inv = {v: k for k, v in spec_versions.items()}

    spec_mapping = SpecVersionMapping(spec_number, spec_title, spec_versions, spec_versions_inv, responsible_group)
    # if spec_number=='23501' or spec_number=='23.501':
    #     print(spec_mapping)

    return spec_mapping


def extract_text_between_delimiters(text, start_delimiter, stop_delimiter) -> str:
    start = text.find(start_delimiter) + len(start_delimiter)
    stop = text.find(stop_delimiter)
    extracted_text = text[start:stop].replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ').replace('---|---','').strip()
    return extracted_text
