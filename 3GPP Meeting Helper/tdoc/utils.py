import collections
import re

# Common Regular Expressions for parsing TDoc names

spec_file_regex = re.compile(r'(?P<series>[\d]{2})(\.)?(?P<number>[\d]{3})(-(?P<version>[\w\d]*))?(\.zip)?')
spec_number_regex = re.compile(r'(?P<series>[\d]{2})\.(?P<number>[\d]{3})')

# Title of a TDocs as in the TDocsByAgenda file
title_cr_regex = re.compile(r'([\d]{2}\.[\d]{3}) CR([\d]{1,4})')

# Used in Word processing for identifying a TDocs as such
tdoc_regex_simple = re.compile(r'[S\d]*-\d\d[\d]+')

# Originally from the config folder. Used throught the document
tdoc_regex = re.compile(r'(?P<group>[S\d]*)-(?P<year>\d\d)(?P<tdoc_number>[\d]+)(?P<revision>r[\d][\d])?')

TS = collections.namedtuple('TS', 'series number version match')


def is_tdoc(tdoc):
    if (tdoc is None) or (tdoc == ''):
        return False
    tdoc = tdoc.strip()
    regex_match = tdoc_regex.match(tdoc)
    if regex_match is None:
        return False
    return regex_match.group(0) == tdoc


def is_ts(tdoc):
    if (tdoc is None) or (tdoc == ''):
        return False
    tdoc = tdoc.strip()
    regex_match = spec_number_regex.match(tdoc)
    if regex_match is None:
        return False
    return regex_match.group(0) == tdoc


def parse_ts_number(ts):
    if ts is None:
        return None
    regex_match = spec_file_regex.match(ts)
    if regex_match is None:
        return None
    grouptdict = regex_match.groupdict()
    full_match = regex_match.group(0)
    if ts != full_match:
        return None
    try:
        series = int(grouptdict['series'])
        number = int(grouptdict['number'])
        version = grouptdict['version']
        if version is None:
            version = ''
    except:
        return None
    return TS(series, number, version, full_match)


def get_tdoc_year(tdoc, include_revision=False):
    """
    Opens a given TDoc identified by the TDoc ID
    Args:
        tdoc: The TDoc ID
        include_revision: Whether the revision number (e.g. S2-220012r01) is also returned

    Returns: The TDoc year, TDoc number. The revision if \include_revision is True

    """
    # Drafts have an asterisk with the revision number
    if '*' in tdoc:
        this_is_a_draft = True
    else:
        this_is_a_draft = False
    tdoc = tdoc.replace('*', '')

    if not is_tdoc(tdoc):
        if not include_revision:
            return None, None
        return None, None, None
    regex_match = tdoc_regex.match(tdoc)
    if regex_match is None:
        return None
    match_groups = regex_match.groupdict()
    year = int(match_groups['year']) + 2000
    tdoc_number = int(match_groups['tdoc_number'])

    if not include_revision:
        return year, tdoc_number

    try:
        revision = match_groups['revision']
        return year, tdoc_number, revision
    except:
        return year, tdoc_number, None