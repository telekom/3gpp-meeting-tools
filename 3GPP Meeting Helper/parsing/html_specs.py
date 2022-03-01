import html2text
import pandas as pd
import re
import traceback
import collections

SpecReleases = collections.namedtuple('SpecReleases', 'folder release')
SpecSeries = collections.namedtuple('SpecSeries', 'folder series')
SpecFile = collections.namedtuple('SpecFile', 'file spec version')


def extract_releases_from_latest_folder(html_content):
    # https://www.3gpp.org/ftp/Specs/latest
    h = html2text.HTML2Text()
    # Ignore converting links from HTML
    h.ignore_links = True
    latest_specs_page_text = h.handle(html_content)
    releases = [SpecReleases(m.group(0), m.group(1)) for m in re.finditer(r'Rel-([\d]{1,2})', latest_specs_page_text)]
    releases = list(set(releases))

    return releases


def extract_spec_series_from_spec_folder(html_content):
    # https://www.3gpp.org/ftp/Specs/latest/Rel-18/23_series
    h = html2text.HTML2Text()
    # Ignore converting links from HTML
    h.ignore_links = True
    series_specs_page_text = h.handle(html_content)
    series = [SpecSeries(m.group(0), m.group(1)) for m in re.finditer(r'([\d]{1,2})_series', series_specs_page_text)]
    series = list(set(series))

    return series


def extract_spec_files_from_spec_folder(html_content):
    # https://www.3gpp.org/ftp/Specs/latest/Rel-18/23_series
    h = html2text.HTML2Text()
    # Ignore converting links from HTML
    h.ignore_links = True
    specs_page_text = h.handle(html_content)
    specs = [SpecFile(m.group(0), m.group(1), m.group(2)) for m in re.finditer(r'([\d]{5})-(.*).zip', specs_page_text)]
    specs = list(set(specs))

    return specs
