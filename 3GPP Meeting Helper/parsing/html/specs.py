import re
# https://stackoverflow.com/questions/52623204/how-to-specify-method-return-type-list-of-what-in-python
# Edit: With the new 3.9 version of Python, you can annotate types without importing from the typing module
from typing import List

from parsing.spec_types import SpecType, SpecReleases, SpecSeries, SpecFile, SpecVersionMapping
from server.common.server_utils import WiEntry


def extract_releases_from_latest_folder(latest_specs_page_text: str, base_url: str) -> List[SpecReleases]:
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
spec_zipfile_regex = re.compile(r'([\d]{5}(-[\d]{1,2})?)-([\d\w]*).zip')
spec_related_wis_regex = re.compile(r'(?P<uid>\d+) *\| *(?P<acronym>[\w, -_]+) *\| *(?P<name>[\w, -()-]+) *\| *(?P<groups>[\w, ]+) *\|')

def extract_spec_files_from_spec_folder(
        specs_page_markup: str,
        base_url: str,
        release: str,
        series: str,
        auto_fill=False) -> List[SpecFile]:
    """
    Extracts the 3GPP series information from the HTML of a release,
    e.g., https://www.3gpp.org/ftp/Specs/latest/Rel-18/23_series. Also works for spec archive pages, e.g.,
    https://www.3gpp.org/ftp/Specs/archive/23_series/23.206
    Args:
        auto_fill: Whether the list should auto-fill series based on the spec number
        specs_page_markup (str): Markup content from which to extract the releases information
        base_url (str): URL where this HTML was extracted. Note that it should NOT end in '/'
        release (str): The release number (not folder) this extraction relates to
        series (str): The series number (not folder) this extraction relates to

    Returns:
        list(SpecFile): List of specs found in the 3GPP site
    """
    # Need to account to TR numbers and old release numbers. Examples:
    # 23700-07-h00.zip, 23003-aa0.zip, 23034-800.zip

    print(f'Base URL: {base_url}')
    versions = [(m.group(0), m.group(1), m.group(3)) for m in spec_zipfile_regex.finditer(specs_page_markup)]
    versions = [v for v in versions if len(v[1]) > 2]
    print(f'Versions: {versions}')
    specs = [
        SpecFile(
            v[0],
            '{0}.{1}'.format(v[1][0:2], v[1][2:]),
            v[2],
            series,
            release,
            base_url,
            base_url + '/' + v[0]) for v in versions]

    if auto_fill:
        specs_autofill = []
        for spec in specs:
            specs_autofill.append(SpecFile(
                spec[0],
                spec[1],
                spec[2],
                spec.spec[0:2],
                spec[4],
                spec[5],
                spec[6],
            ))
        specs = specs_autofill

    specs = list(set(specs))

    return specs


# e.g., [1.0.0](https://www.3gpp.org/ftp/Specs/archive/21_series/21.101/21101-100.zip
spec_version_regex = re.compile(r'\[([\d\.]+)]\(https://.*/(.*\.zip)')

# e.g. |  2022-04-20 |
spec_upload_date_regex = re.compile(r'\|  ([\d]{4}-[\d]{2}-[\d]{2}) \|')


def extract_spec_versions_from_spec_file(spec_page_markup) -> SpecVersionMapping:
    """
    Extracts the 3GPP series information from the HTML of specification, e.g. https://www.3gpp.org/DynaReport/23501.htm
    The file is expected to be passed on in Markup form, not HTML. A call to html2text.HTML2Text() can provide this
    conversion
    Args:
        spec_page_markup (str): Markup content from which to extract the releases information

    Returns:
        list(SpecFile): List of specs found in the 3GPP site
    """

    # Without the dot: 23.501->23501
    spec_number = cleanup_spec_name(extract_text_between_delimiters(
        spec_page_markup,
        start_delimiter="Specification #: ",
        stop_delimiter="  * General"))
    spec_title = extract_text_between_delimiters(
        spec_page_markup,
        start_delimiter="Title: |  ",
        stop_delimiter="Status: |  ")
    responsible_group = extract_text_between_delimiters(
        spec_page_markup,
        start_delimiter="Primary responsible group: |  ",
        stop_delimiter="Secondary responsible groups: |  ")
    spec_type = extract_text_between_delimiters(
        spec_page_markup,
        start_delimiter="Type: |  ",
        stop_delimiter="Initial planned Release: |  ")
    spec_initial_release = extract_text_between_delimiters(
        spec_page_markup,
        start_delimiter="Initial planned Release: |  ",
        stop_delimiter="Internal: |")

    if 'Technical report (TR)' in spec_type:
        spec_type_enum = SpecType.TR
    elif 'Technical specification (TS)' in spec_type:
        spec_type_enum = SpecType.TS
    else:
        spec_type_enum = SpecType.Unknown

    upload_dates = [m.group(1) for m in spec_upload_date_regex.finditer(spec_page_markup) ]

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

    try:
        related_wis_relevant_markup = [e for e in spec_page_markup.split('{1}') if e is not None and 'Related Work Items' in e]
        related_wis_relevant_markup = '\n'.join(related_wis_relevant_markup)
    except Exception as e:
        print(f'Could not extract related WIs for {spec_number}: {e}')
        related_wis_relevant_markup = ''

    related_wis = [WiEntry(
        uid = m.group('uid').strip(),
        code = m.group('acronym').strip(),
        title = m.group('name').strip(),
        lead_body = m.group('groups').strip(),
        release = ''
    ) for m in spec_related_wis_regex.finditer(related_wis_relevant_markup)]
    print(f'Related WIs: {related_wis}')

    spec_mapping = SpecVersionMapping(
        spec_number,
        spec_title,
        spec_versions,
        spec_versions_inv,
        responsible_group,
        spec_type_enum,
        spec_initial_release,
        upload_dates,
        related_wis=related_wis
    )
    # if spec_number=='23501' or spec_number=='23.501':
    #     print(spec_mapping)

    return spec_mapping


def extract_text_between_delimiters(text, start_delimiter, stop_delimiter) -> str:
    start = text.find(start_delimiter) + len(start_delimiter)
    stop = text.find(stop_delimiter)
    extracted_text = text[start:stop].replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ').replace('---|---',
                                                                                                         '').strip()
    return extracted_text


def cleanup_spec_name(spec_id: str, clean_type=True, clean_dots=True) -> str:
    """
    For a given textual Specification name, return the string used for indexing specs, i.e., without dots and/or
    without a TS/TR string
    Args:
        clean_dots: Whether to clean the dot in the spec, e.g. 23.501->23501
        clean_type: Whether to clean the specification type, e.g. TS 23.501->23.501
        spec_id: The specification name, e.g., TS 23.501, 23.501, 23501

    Returns: The specification names as used in the DataFrame indexes, e.g., 23501

    """
    if spec_id is None:
        return None
    clean_spec_id = spec_id
    if clean_type:
        clean_spec_id = clean_spec_id.replace('TS', '').replace('TR', '').strip()
    if clean_dots:
        clean_spec_id = clean_spec_id.replace('.', '').strip()
    return clean_spec_id
