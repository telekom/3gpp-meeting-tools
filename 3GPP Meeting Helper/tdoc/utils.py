import collections
import re
import threading
import traceback
from typing import Callable, List, Any
import parsing.word.pywin32 as word_parser

# Common Regular Expressions for parsing TDoc names

spec_file_regex = re.compile(r'(?P<series>[\d]{2})(\.)?(?P<number>[\d]{3})(-(?P<version>[\w\d]*))?(\.zip)?')
spec_number_regex = re.compile(r'(?P<series>[\d]{2})\.(?P<number>[\d]{3})')

# Title of a TDocs as in the TDocsByAgenda file
title_cr_regex = re.compile(r'([\d]{2}\.[\d]{3}) CR([\d]{1,4})')

# Used in Word processing for identifying a SA2 TDoc as such
tdoc_sa2_regex_simple = re.compile(r'[S\d]*-\d\d[\d]+')

# Originally from the config folder. Used through the document
tdoc_regex_str = r'(?P<group>[S\d]*)-(?P<year>\d\d)(?P<tdoc_number>[\d]+)(?P<revision>r[\d][\d])?'
tdoc_regex = re.compile(tdoc_regex_str)

TS = collections.namedtuple('TS', 'series number version match')

# Generic TDoc regex used for Tdoc search. Note that early meetings did not use years
# Sometimes "_" is (erroneously) used, so I had to relax the regex
tdoc_generic_regex = re.compile(r'(?P<group>[\w\d]+)[\-_](?P<number>[\d]+)')

# Sometimes using a different format for SA3-LI
tdoc_sa3li_regex = re.compile(r'(?P<group>[sS]3i)(?P<number>[\d]+)')


def is_sa2_tdoc(tdoc: str):
    if (tdoc is None) or (tdoc == ''):
        return False
    tdoc = tdoc.strip()
    regex_match = tdoc_regex.match(tdoc)
    if regex_match is None:
        return False
    return regex_match.group(0) == tdoc


class GenericTdoc:

    def __init__(self, tdoc_id: str):
        if tdoc_id is None:
            raise ValueError('tdoc_id cannot be None')
        tdoc_match = tdoc_generic_regex.match(tdoc_id)
        if tdoc_match is None:
            tdoc_match = tdoc_sa3li_regex.match(tdoc_id)
            if tdoc_match is None:
                raise ValueError('tdoc_id is not a TDoc number including group and TDoc number')
        self._group = tdoc_match.group('group')
        self._number = int(tdoc_match.group('number'))
        self._tdoc = tdoc_id

    @property
    def group(self):
        return self._group

    @property
    def number(self):
        return self._number

    @property
    def tdoc(self):
        return self._tdoc

    def __str__(self) -> str:
        return f'{self.tdoc}'

    def __repr__(self) -> str:
        return f'GenericTdoc(\'{self.tdoc}\')'


def is_generic_tdoc(tdoc: str) -> GenericTdoc | None:
    """
    Parses a TDoc ID and returns (if matching) information regarding this TDoc.
    Args:
        tdoc: A TDoc ID

    Returns: Parsed TDoc information. None if this is not a valid TDoc ID

    """
    if (tdoc is None) or (tdoc == ''):
        return None
    tdoc = tdoc.strip()
    try:
        return GenericTdoc(tdoc)
    except ValueError as e:
        print(f'Tdoc {tdoc} is not a valid TDoc: {e}')
        return None


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

    if not is_sa2_tdoc(tdoc):
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


def do_something_on_thread(
        task: Callable[[], None] | None,
        before_starting: Callable[[], None] | None = None,
        after_task: Callable[[], None] | None = None,
        on_error_log: str = None):
    """
    Does something on a Thread (e.g. bulk download)
    Args:
        on_error_log: What to print in case of an exception
        task: The task to do
        before_starting: Something to do before starting the task
        after_task: Something to do after the task is finished or if an exception is thrown
    """
    if before_starting is not None:
        before_starting()

    def thread_task():
        try:
            task()
        except:
            if on_error_log is not None:
                print(on_error_log)
            traceback.print_exc()
        finally:
            if after_task is not None:
                after_task()

    t = threading.Thread(target=thread_task)
    t.start()


open_tdoc_for_compare_fn: Callable[[str, List[Any]], None] | None = None


def compare_tdocs(
        entry_1: str | None = None,
        entry_2: str | None = None,
        get_entry_1_fn: Callable[..., str] | None = None,
        get_entry_2_fn: Callable[..., str] | None = None
):
    try:
        tdocs_1 = []
        tdocs_2 = []
        if (entry_1 is None) and (get_entry_1_fn is not None):
            entry_1 = get_entry_1_fn()
        if (entry_2 is None) and (get_entry_2_fn is not None):
            entry_2 = get_entry_1_fn()
        match_1 = tdoc_regex.match(entry_1)
        match_2 = tdoc_regex.match(entry_2)

        # Strip revision number from any input (we will search for the matching document on the list)
        search_1 = '{0}-{1}{2}'.format(match_1.group('group'), match_1.group('year'), match_1.group('tdoc_number'))
        search_2 = '{0}-{1}{2}'.format(match_2.group('group'), match_2.group('year'), match_2.group('tdoc_number'))

        # Download (cache) documents to compare
        if open_tdoc_for_compare_fn is None:
            print(f'Could not open documents. Document open function not set')
            return

        open_tdoc_for_compare_fn(
            entry_1,
            tdocs_1)
        open_tdoc_for_compare_fn(
            entry_2,
            tdocs_2)

        # There may be several documents (e.g. other TDocs as attachment). Strip the list to the most likely
        # candidates to be the actual TDoc
        tdocs_1 = [e for e in tdocs_1 if search_1 in e]
        tdocs_2 = [e for e in tdocs_2 if search_2 in e]

        print('TDoc to compare 1: {0}'.format(tdocs_1))
        print('TDoc to compare 2: {0}'.format(tdocs_2))

        if len(tdocs_1) == 0 or len(tdocs_2) == 0:
            print('Need two TDocs to compare. One of them does not contain TDocs')
            return

        tdocs_1 = tdocs_1[0]
        tdocs_2 = tdocs_2[0]

        word_parser.compare_documents(tdocs_1, tdocs_2)
    except:
        print('Could not compare documents')
        traceback.print_exc()
