import collections
import datetime
import hashlib
import os
import pickle
import re
import traceback

import pandas as pd
from lxml import html as lh

import config.contributor_names
from parsing.html.common import get_cache_filepath, current_cache_version, comment_span, join_results, max_recursion, \
    sort_and_remove_duplicates_from_list
from parsing.html.common_tools import tdoc_regex_str, parse_tdoc_comments
from parsing.html.tdocs_by_agenda_v3 import assert_if_tdocs_by_agenda_post_sa2_159, parse_tdocs_by_agenda_v3
from server.common import decode_string
from tdoc.utils import title_cr_regex


class TdocsByAgendaData(object):
    """Contains the information for the "Tdocs by Agenda" meeting file"""
    revision_of_regex = re.compile(r'.*Revision of (?P<tdoc>[S\d-]*)( from (?P<previous_meeting>S[\d\w#-]*))?')
    revised_to_regex = re.compile(
        r'.*Revised([ ]?(off-line|in parallel session|in drafting session))? to (?P<tdoc>[S\d-]*)')
    merge_of_regex = re.compile(r'.*merging (?P<tdoc>(( and )?(, )?(part of)?( )?[S\d-]*)+)')
    merged_to_regex = re.compile(r'.*Merged (into|with) (?P<tdoc>[S\d-]*)')

    # Strings to match:
    #  - Draft Agenda for SA WG2#128
    #  - Draft Agenda for SA WG2#129#BIS
    #  - SA WG2 #137 Meeting Agenda
    #  - SA2#155 Meeting Agenda
    meeting_number_regex = re.compile(r'((SA WG2[ ]?)|(SA2))#(?P<meeting>[\d]{1,3}[#]?[\w]*)')

    creation_date_regex = re.compile(
        r'>Created: <(B|b)>((?P<year>[\d]{4})-(?P<month>[\d]{2})-(?P<day>[\d]{2}) (?P<hour>[\d]{2}):(?P<minute>[\d]{2}))</(B|b)>&nbsp[;]?&nbsp[;]?&nbsp[;]?&nbsp[;]?')
    creation_date_regex_if_fails = re.compile(
        r'<o:LastSaved>((?P<year>[\d]{4})-(?P<month>[\d]{2})-(?P<day>[\d]{2})T(?P<hour>[\d]{2}):(?P<minute>[\d]{2}))')

    # Hardcoded list of typos to correct
    tdoc_typos = {
        'S2-181812649': 'S2-1812649',
        'S2-19000963': 'S2-1900963'
    }

    def get_tdoc_by_agenda_html(path_or_html, return_raw_html=False):
        try:
            # Initial check added, as it seemed to sometimes break things if the cache directory does not exist in fresh installations
            if (len(path_or_html)) < 1000 and os.path.isfile(path_or_html):
                # If the HTML is input into the open function, it will throw an exception
                print('Reading TDocsByAgenda from {0}'.format(path_or_html))
                with open(path_or_html, 'rb') as f:  # with can auto close the file like f.close() does
                    html = f.read()
            else:
                html = path_or_html
        except Exception as e:
            html = path_or_html
            print(f'Could not parse HTML or path: {e}')

        if html is None:
            print('No HTML to parse. Maybe a communication error occurred?')
            return None

        print('TDocsByAgenda HTML: {0} bytes'.format(len(html)))

        if return_raw_html:
            return decode_string(html, 'TDocsByAgenda')
        else:
            try:
                parsed_html = lh.fromstring(html)
            except Exception as e:
                print(f'Could not parse TDocs by Agenda HTML: {e}')
                parsed_html = None
            return parsed_html

    def get_tdoc_by_agenda_date(path_or_html: str) -> datetime.datetime:
        html = TdocsByAgendaData.get_tdoc_by_agenda_html(path_or_html, return_raw_html=True)
        email_approval_results = False

        if html is None:
            print('Cannot not read TDocs by Agenda file')
            return None

        try:
            search_result = TdocsByAgendaData.creation_date_regex.search(html)
            if search_result is None:
                print('Could not parse date from TDocs by Agenda file, trying with LastSaved')
                search_result = TdocsByAgendaData.creation_date_regex_if_fails.search(html)
                email_approval_results = (html.find('This CR was agreed') != -1)
                if email_approval_results:
                    print(
                        'TDocsByAgenda contains e-mail approval results. Assuming last version and returning now() as time')
                    return datetime.datetime.now()
                if search_result is None:
                    print('Could not parse date from TDocs by Agenda file')
                    return None
            year = int(search_result.group('year'))
            month = int(search_result.group('month'))
            day = int(search_result.group('day'))
            hour = int(search_result.group('hour'))
            minute = int(search_result.group('minute'))
            parsed_date = datetime.datetime(year=year, month=month, day=day, hour=hour, minute=minute)
            return parsed_date
        except:
            print('Error parsing date of TDocs by Agenda file')
            traceback.print_exc()
            return None

    def __init__(self, path_or_html, v=2, html_hash='', meeting_server_folder=''):
        self.tdocs: pd.DataFrame | None = None

        print('Parsing TDocsByAgenda file: version {0}'.format(v))
        raw_html = TdocsByAgendaData.get_tdoc_by_agenda_html(path_or_html, return_raw_html=True)

        try:
            self.meeting_number = TdocsByAgendaData.get_meeting_number(raw_html)
        except Exception as e:
            self.meeting_number = 'Unknown'
            print(f'Could not get Meeting number: {e}')
        print('Parsed meeting number: {0}'.format(self.meeting_number))
        self.meeting_server_folder: str = meeting_server_folder

        dataframe_from_cache = False

        if v == 1:
            # print('XPath fro title: ' + html.xpath('//P/FONT/B').tostring())
            html = TdocsByAgendaData.get_tdoc_by_agenda_html(path_or_html)
            dataframe = TdocsByAgendaData.read_tdocs_by_agenda(html)
        else:
            if meeting_server_folder != '' and html_hash != '':
                cache_file_name = get_cache_filepath(meeting_server_folder, html_hash)
                try:
                    if cache_file_name is not None and os.path.exists(cache_file_name):
                        with open(cache_file_name, 'rb') as f:
                            # Unpickle the 'data' dictionary using the highest protocol available.
                            cache = pickle.load(f)
                            if cache['cache_version'] == current_cache_version:
                                dataframe = cache['tdocs']
                                dataframe_from_cache = True
                                print('Loaded TDocsByAgenda from file cache: {0}'.format(cache_file_name))
                                remove_old_cache = False
                            else:
                                print('Cache version mismatch. Not reading from cache. Removing old cache')
                                remove_old_cache = True

                        if remove_old_cache:
                            print('New cache version. Removing old cache (will get re-saved with new version)')
                            os.remove(cache_file_name)

                except:
                    print(
                        'Could not load file cache for meeting {0}, hash {1}'.format(meeting_server_folder, html_hash))
                    traceback.print_exc()

            if not dataframe_from_cache:
                dataframe = TdocsByAgendaData.read_tdocs_by_agenda_v2(raw_html, force_html=True)

        # Cleanup Unicode characters (see https://stackoverflow.com/questions/42306755/how-to-remove-illegal-characters-so-a-dataframe-can-write-to-excel)
        if not dataframe_from_cache:
            print('Cleaning up Unicode characters so that Excel export does not crash')
            dataframe = dataframe.map(
                lambda x: x.encode('unicode_escape').decode('utf-8') if isinstance(x, str) else x)

        # Cleanup comments. Sometimes we have "span" tags polluting comments
        if not dataframe_from_cache:
            print('Cleaning up comments column')
            try:
                dataframe['Comments'] = dataframe['Comments'].apply(lambda x: TdocsByAgendaData.clean_up_comment(x))
            except:
                print('Could not clean-up comments')
                traceback.print_exc()

        # Other cleanups that happened over the time
        dataframe['Title'] = dataframe['Title'].apply(lambda x: TdocsByAgendaData.clean_up_title(x))

        # Assign dataframe
        self.tdocs = dataframe

        if not dataframe_from_cache:
            TdocsByAgendaData.get_original_and_final_tdocs(self.tdocs)
            self.others_cosigners, self.tdocs = config.contributor_names.add_contributor_columns_to_tdoc_list(
                self.tdocs, self.meeting_server_folder)
            self.contributor_columns = config.contributor_names.get_contributor_columns()
        else:
            # get_original_and_final_tdocs should already be in the cache
            self.others_cosigners = cache['others_cosigners']
            self.contributor_columns = cache['contributor_columns']
        config.contributor_names.reset_others()

    def clean_up_comment(comment_str):
        comment_match = comment_span.match(comment_str)
        if comment_match is None:
            return comment_str
        fixed_comment = '{0}. {1}'.format(comment_match.group(1), comment_match.group(2))
        return fixed_comment

    def clean_up_title(title_str: str) -> str:
        if title_str is None:
            return title_str

        return title_str.replace(r'&apos;', "'").replace(r'&amp;', "&").replace(r'&#39;', "'")

    def get_meeting_number(tdocs_by_agenda_html:str) -> str:
        """
        Returns: The meeting number based on the HTML of TDocsByAgenda
        """
        print('Parsing TDocsByAgenda meeting number')
        meeting_number_match = TdocsByAgendaData.meeting_number_regex.search(tdocs_by_agenda_html)
        if meeting_number_match is None:
            print('Could not parse meeting number from TDocsByAgenda HTML')
            return 'Unknown'
        meeting_number = meeting_number_match.groupdict()['meeting']
        meeting_number = meeting_number.replace('#', '')
        print('Meeting number: {0}'.format(meeting_number))
        return meeting_number.upper()

    def read_tdocs_by_agenda_v2(path_or_html, force_html=False) -> pd.DataFrame:
        # New HTML-parsing method. Regex-based. It works better with malformed and broken HTML files, which appear to happen quite often
        if force_html:
            html = path_or_html
        else:
            html = TdocsByAgendaData.get_tdoc_by_agenda_html(path_or_html, return_raw_html=True)
        print('TDocsByAgenda: HTML file length: {0}'.format(len(html)))

        if assert_if_tdocs_by_agenda_post_sa2_159(html):
            print("TDocsByAgenda is newer than SA2#159")
            df_tdocs = parse_tdocs_by_agenda_v3(html)
            # Post-processing
            df_tdocs = TdocsByAgendaData.post_process_df_tdocs(df_tdocs)
            return df_tdocs
        else:
            print("TDocsByAgenda is prior to SA2#159")

        # Remove beginning
        pre_cleaning_html_length = len(html)
        start_of_tdocs = html.find('List for Meeting:')
        if start_of_tdocs > -1:
            html = html[start_of_tdocs:-1]
            print('Skipped first {0} characters until "List for Meeting:"'.format(start_of_tdocs))

        uncapitalized_separator = '<tr valign='
        capitalized_separator = '<TR VALIGN='
        if html.find(uncapitalized_separator) != -1:
            splitter = uncapitalized_separator
        else:
            splitter = capitalized_separator

        separators = ['(td|TD)', '(bordercolor|BORDERCOLOR)', '(bgcolor|BGCOLOR)', '(font|FONT)',
                      '(font-size|FONT-SIZE)', '(face|FACE)', '(color|COLOR)', '(th|TH)']

        # Strings that will be substituted
        s = {
            'TD': separators[0],
            'TH': separators[7],
            'BORDERCOLOR': separators[1],
            'BGCOLOR': separators[2],
            'FONT': separators[3],
            'FONT-STYLE': '[\\\"]?FONT-SIZE:[\\d]+pt[\\\"]?',
            'FACE': separators[5],
            'COLOR': separators[6],
            'B': 'b',  # Bold tag is always lower-case
            'HEX': '([\\\"]?(#[0-9a-fA-F]+)?[\\\"]?)?',
            '=': '[ ]*=[ ]*',
            'ARIAL': '[\\\"]?Arial[\\\"]?',
            'CRLF': '[\r\n]?',
            'TDOC': tdoc_regex_str,
            'INTEGER': '[\\\"]?[\\d]*[\\\"]?',
        }

        # Generate big regex that will substitute all of the unwanted strings for easier parsing
        strings_to_sub = [
            ' </{B}></{FONT}>'.format(**s),
            '<{FONT} style{=}{FONT-STYLE} {FACE}{=}({ARIAL})?[ \"]{COLOR}{=}{HEX}>{CRLF}'.format(**s),
            '</{FONT}>'.format(**s),
            '<{TD} {BORDERCOLOR}{=}{HEX} {BGCOLOR}{=}{HEX}>'.format(**s),
            '<{TD} [wW][iI][dD][tT][hH]={INTEGER} {BORDERCOLOR}{=}{HEX} {BGCOLOR}{=}{HEX}>'.format(**s),
            '<{TH} (width={INTEGER} )?({TH}[ ]?(={INTEGER})?)?[ ]?{BGCOLOR}{=}{HEX} {BORDERCOLOR}{=}{HEX}[ ]?>'.format(
                **s),
            '<[/]?[bBiI]>',
            '<TD  WIDTH="75" BORDERCOLOR=#000000 BGCOLOR =#C0C0C0>',
            '<[/]?{TH}>'.format(**s)
        ]
        full_sub = '|'.join(strings_to_sub)

        # Cleanup HTML tags
        html = re.sub(full_sub, '', html)

        post_cleaning_html_length = len(html)
        gain = post_cleaning_html_length / pre_cleaning_html_length
        print('TDocsByAgenda: HTML file length after cleaning: {0} ({1:.0%})'.format(len(html), gain))

        split_html = html.split(splitter)
        print('TDocsByAgenda: split in {0} segments. Parsing segments for TDoc info'.format(len(split_html)))

        tdoc_columns = None
        for element in split_html:
            if 'TD#' in element:
                try:
                    element_clean = re.sub('</{TD}>([\r\n])*'.format(**s), '\n', element)
                    tdoc_columns = [e.strip() for e in element_clean.split('\n')[1:-4]]
                    print('Columns found: {0}'.format(tdoc_columns))
                    break
                except:
                    pass

        # Build regex to parse columns
        tdocs_by_agenda_regex = '.*)</{TD}>{CRLF}'.format(**s)
        tdocs_by_agenda_email_disc_regex = '.*({CRLF}=== Close of meeting ===)?)</{TD}>{CRLF}'.format(**s)
        tdocs_by_agenda_tdoc_regex = '[ ]*' + '.*(?P<TD>{TDOC}).*</{TD}>{CRLF}'.format(**s)
        tdoc_by_agenda_simple_tdoc = re.compile(r'(?P<TD>{TDOC})'.format(**s))

        # Take default columns if none are found
        if tdoc_columns is None:
            tdocs_by_agenda_regex_full = '(?P<AI>{0}{1}(?P<TYPE>{0}(?P<Doc For>{0}(?P<Subject>{0}(?P<Source>{0}(?P<Rel>{0}(?P<Work Item>{0}(?P<Comments>{0}(?P<Result>{0}'.format(
                tdocs_by_agenda_regex, tdocs_by_agenda_tdoc_regex)
            column_indices = {
                'AI': 0,
                'TD#': 1,
                'Type': 2,
                'Doc For': 3,
                'Subject': 4,
                'Source': 5,
                'Rel': 6,
                'Work Item': 7,
                'Comments': 8,
                'Result': 9,
                'Subject': None,
                'e-mail_Discussion': None
            }
            total_columns = 10
        # Detect columns in file
        else:
            # Column is sometimes called "Title" sometimes "Subject"
            column_indices = {
                'AI': None,
                'TD#': None,
                'Type': None,
                'Doc For': None,
                'Subject': None,
                'Source': None,
                'Rel': None,
                'Work Item': None,
                'Comments': None,
                'Result': None,
                'Title': None,
                'e-mail_Discussion': None
            }
            for e in ['AI', 'TD#', 'Type', 'Doc For', 'Subject', 'Source', 'Title', 'Rel', 'Work Item', 'Comments',
                      'Result', 'e-mail_Discussion']:
                if e in tdoc_columns:
                    found_index = tdoc_columns.index(e)
                    column_indices[e] = found_index

            # Sort parameter list according to their column index to generate search regex
            column_indices_list = [e for e in column_indices.items()]
            column_indices_list.sort(key=lambda x: x[1] if x[1] is not None else -1)

            # Generate Regex
            tdocs_by_agenda_regex_full = ''
            next_expected_index = None
            total_columns = 0
            for column_name, column_index in column_indices_list:
                if column_index is not None and next_expected_index is None:
                    if 'TD#' in column_name:
                        tdocs_by_agenda_regex_full = '{0}{1}'.format(tdocs_by_agenda_regex_full,
                                                                     tdocs_by_agenda_tdoc_regex)
                        total_columns += 1
                    elif 'e-mail_Discussion' in column_name:
                        tdocs_by_agenda_regex_full = '{0}(?P<{1}>{2}'.format(tdocs_by_agenda_regex_full,
                                                                             column_name.replace(' ', '').replace('-',
                                                                                                                  ''),
                                                                             tdocs_by_agenda_email_disc_regex)
                        total_columns += 1
                    else:
                        tdocs_by_agenda_regex_full = '{0}(?P<{1}>{2}'.format(tdocs_by_agenda_regex_full,
                                                                             column_name.replace(' ', '').replace('-',
                                                                                                                  ''),
                                                                             tdocs_by_agenda_regex)
                        total_columns += 1

        print('TDocsByAgenda: Matching regular expressions')

        TdocMatch = collections.namedtuple('TdocMatch',
                                           'AI TD Type DocFor Title Source Rel WorkItem Comments Result ParsedComments')
        line_separator = re.compile(r'</[tT][dD]>')
        if column_indices['Subject'] is not None:
            title_name = 'Subject'
        else:
            title_name = 'Title'

        def convert_split_to_tdoc_info(info, td_column_idx):
            info_split = line_separator.split(info)
            if len(info_split) != total_columns + 1:
                return None

            tdoc_info = info_split[0:-1]
            found_tdoc = tdoc_by_agenda_simple_tdoc.search(tdoc_info[td_column_idx])
            if found_tdoc is None:
                return None
            tdoc_id = found_tdoc.group('TD')
            tdoc_info[td_column_idx] = tdoc_id
            tdoc_info = [e.strip() for e in tdoc_info]

            try:
                def get_str(column_name):
                    idx = column_indices[column_name]
                    if idx is None:
                        return ''
                    return tdoc_info[idx]

                ai = get_str('AI').replace('"', '').replace('TOP>', '').replace('\n', '').strip()
                comments = get_str('Comments').replace('<span title="">', '').replace('</span>', '').replace(
                    '<span title=" ">', '').strip()
                parsed_tdoc_info = TdocMatch(
                    ai,
                    tdoc_id,
                    get_str('Type'),
                    get_str('Doc For'),
                    get_str(title_name),
                    get_str('Source'),
                    get_str('Rel'),
                    get_str('Work Item'),
                    comments,
                    get_str('Result'),
                    parse_tdoc_comments(comments))
                return parsed_tdoc_info
            except:
                return None

        td_column_idx = column_indices['TD#']
        tdoc_matches = [convert_split_to_tdoc_info(e, td_column_idx) for e in split_html]
        tdoc_matches = [e for e in tdoc_matches if e is not None]

        print('TDocsByAgenda: Organizing TDoc matches')
        print('TDocsByAgenda: found {0} TDocs'.format(len(tdoc_matches)))
        column_for_dictionary = list(zip(*tdoc_matches))

        print('TDocsByAgenda: Generating TDocs DataFrame')
        d = {
            'AI': column_for_dictionary[0],
            'TD#': column_for_dictionary[1],
            'Type': column_for_dictionary[2],
            'Doc For': column_for_dictionary[3],
            'Title': column_for_dictionary[4],
            'Source': column_for_dictionary[5],
            'Rel': column_for_dictionary[6],
            'Work Item': column_for_dictionary[7],
            'Comments': column_for_dictionary[8],
            'Result': column_for_dictionary[9],
            'Revision of': [e.revision_of for e in column_for_dictionary[10]],
            'Revised to': [e.revised_to for e in column_for_dictionary[10]],
            'Merge of': [e.merge_of for e in column_for_dictionary[10]],
            'Merged to': [e.merged_to for e in column_for_dictionary[10]],
            '#': range(len(tdoc_matches))
        }

        print('TDocsByAgenda: Done')

        # Some cleaning for some special characters
        d['Source'] = [at_n_t.replace('&amp;', '&') for at_n_t in d['Source']]

        df_tdocs = pd.DataFrame(data=d)
        df_tdocs = df_tdocs.set_index('TD#')
        df_tdocs.index.name = 'TD#'
        print('TDocsByAgenda: {0} TDocs entries parsed'.format(len(df_tdocs)))
        email_approval_tdocs = df_tdocs[(df_tdocs['Result'] == 'For e-mail approval')]
        n_email_approval = len(email_approval_tdocs)
        print('TDocsByAgenda: {0} TDocs marked as "For e-mail approval"'.format(n_email_approval))

        # Post-processing
        df_tdocs = TdocsByAgendaData.post_process_df_tdocs(df_tdocs)

        return df_tdocs

    def read_tdocs_by_agenda(doc):
        # Legacy method for parsing TdocsByAgenda file. HTML-based parsing assumes a correct HTML file
        word_exported = False
        tdoc_table = doc.xpath('//table')[-1]
        tdoc_header = [element.text_content().strip() for element in tdoc_table.xpath('thead')[0].xpath('tr/th')]
        if len(tdoc_header) == 0:
            tdoc_header = [element.text_content().strip() for element in tdoc_table.xpath('thead')[0].xpath('tr/td')]
            tdoc_header = [e for e in tdoc_header if (e is not None) and (e != '')]
            if len(tdoc_header) > 0:
                word_exported = True
        tdoc_header.extend(['Revision of', 'Revised to', 'Merge of', 'Merged to', '#'])

        if not word_exported:
            tdocs = tdoc_table.xpath('tbody/tr')
        else:
            tdocs = tdoc_table.xpath('tr')

        rows = []
        row_idx = 0
        for tdoc in tdocs:
            tdoc_cols = tdoc.xpath('td')
            cols_as_array = [element.text_content().strip() for element in tdoc_cols]
            if word_exported:
                cols_as_array = cols_as_array[1:]
            new_cols = ['', '', '', '', row_idx]
            cols_as_array.extend(new_cols)

            # Sort out non-TDoc rows (e.g. separator rows)
            if (cols_as_array[1] == '-') or (cols_as_array[1] == ''):
                continue

            # The row index is the last column
            cols_as_array[-1] = row_idx

            added_columns = len(new_cols)

            # Extract chairman comments (auto-generated text from Word macro)
            comments = cols_as_array[-added_columns - 2]
            # print(cols_as_array)
            # print('Comments (row ' + str(-len(new_cols)-2) + '): ' + comments)

            # revision_of_match = tdocs_by_agenda.revision_of_regex.match(comments)
            # revised_to_match = tdocs_by_agenda.revised_to_regex.match(comments)
            # merge_of_match = tdocs_by_agenda.merge_of_regex.match(comments)
            # merged_to_match = tdocs_by_agenda.merged_to_regex.match(comments)

            parsed_comments = parse_tdoc_comments(comments)
            cols_as_array[-(added_columns - 2)] = parsed_comments.merge_of
            cols_as_array[-(added_columns - 3)] = parsed_comments.merged_to
            cols_as_array[-(len(new_cols))] = parsed_comments.revision_of
            cols_as_array[-(added_columns - 1)] = parsed_comments.revised_to

            rows.append(cols_as_array)
            row_idx += 1

        df_tdocs = pd.DataFrame(rows, columns=tdoc_header)
        df_tdocs = df_tdocs.set_index('TD#')
        print('{0} TDocs entries parsed'.format(len(df_tdocs)))

        # Post-processing
        df_tdocs = TdocsByAgendaData.post_process_df_tdocs(df_tdocs)

        return df_tdocs

    def post_process_df_tdocs(df_tdocs):
        # Remove duplicates. It was seen in S2-134 that sometimes duplicate #TD can be present
        # See https://www.dataquest.io/blog/settingwithcopywarning/ as o why the "copy()"
        df_tdocs = df_tdocs.loc[~df_tdocs.index.duplicated(keep='last')].copy()
        print('TDocsByAgenda: {0} TDocs entries parsed after de-duplication'.format(len(df_tdocs)))

        # Fix LS OUTs that have a wrong source
        print('Fixing wrong LS OUT sources')
        ls_with_wrong_source_idx = (df_tdocs['Type'] == 'LS OUT') & (df_tdocs['Revision of'] != '') & (
                df_tdocs['Source'] == 'SA WG2')
        ls_outs_with_wrong_source = df_tdocs.loc[ls_with_wrong_source_idx, :]
        for idx, row in ls_outs_with_wrong_source.iterrows():
            try:
                revision_idx = row['Revision of']
                old_source = df_tdocs.at[revision_idx, 'Source']
                df_tdocs.at[idx, 'Source'] = old_source
                print('{0}: {1}->{2}'.format(idx, row['Source'], df_tdocs.at[idx, 'Source']))
            except:
                pass
        print('Fixed {0} LS OUT sources'.format(len(ls_outs_with_wrong_source)))

        # Extract CR info
        print('TDocsByAgenda: Parsing TS and CR numbers')
        df_tdocs.loc[:, 'TS'] = ''
        df_tdocs.loc[:, 'CR'] = ''

        matches = [(idx, title_cr_regex.match(row['Title'])) for idx, row in df_tdocs.iterrows()]
        matches = [match for match in matches if match[1] is not None]
        for match in matches:
            df_tdocs.at[match[0], 'TS'] = match[1].group(1)
            df_tdocs.at[match[0], 'CR'] = match[1].group(2)
        print('TDocsByAgenda: Parsing TS and CR numbers Done')

        email_approval_tdocs = df_tdocs[(df_tdocs['Result'] == 'For e-mail approval')]
        n_email_approval = len(email_approval_tdocs)
        print('TDocsByAgenda: {0} TDocs marked as "For e-mail approval" after de-duplication'.format(n_email_approval))

        return df_tdocs

    def get_original_and_final_tdocs(df_tdocs):
        print('TDocsByAgenda: Tracking original/final tdocs')

        # Final pass to write the original predecessors and final children
        # Added .astype(str) to avoid some exceptions
        df_tdocs['Original TDocs'] = df_tdocs['Revision of'].astype(str) + ',' + df_tdocs['Merge of'].astype(str)
        df_tdocs['Final TDocs'] = df_tdocs['Revised to'].astype(str) + ',' + df_tdocs['Merged to'].astype(str)

        for index in df_tdocs.index:
            original_tdocs = df_tdocs.at[index, 'Original TDocs']
            final_tdocs = df_tdocs.at[index, 'Final TDocs']

            if original_tdocs != '':
                if original_tdocs == ',':
                    df_tdocs.at[index, 'Original TDocs'] = index
                else:
                    df_tdocs.at[index, 'Original TDocs'] = TdocsByAgendaData.get_original_tdocs(
                        original_tdocs,
                        df_tdocs,
                        index, 0).replace(',', ', ')
            if final_tdocs != '':
                if final_tdocs == ',':
                    df_tdocs.at[index, 'Final TDocs'] = index
                else:
                    df_tdocs.at[index, 'Final TDocs'] = TdocsByAgendaData.get_final_tdocs(
                        final_tdocs,
                        df_tdocs,
                        index,
                        0).replace(',', ', ')

        print('TDocsByAgenda: Finished tracking original/final tdocs')

    # Given a TDoc, returns the TDoc or TDocs that originated this TDoc
    def get_original_tdocs(tdocs, df_tdocs, original_index, n_recursion):
        tdocs_split = tdocs.split(',')
        if len(tdocs_split) > 1:
            tdocs_split = [e for e in tdocs_split if (e != '') and (e is not None)]
        if len(tdocs_split) > 1:
            return join_results(tdocs_split, df_tdocs, TdocsByAgendaData.get_original_tdocs, original_index, n_recursion)

        # We know that length is 1
        tdoc = tdocs_split[0].strip()

        # Fix for 137E final TDocsByAgenda
        if n_recursion > max_recursion:
            print('Maximum recursion reached ({0}) for {1}. Stopping search.'.format(max_recursion, original_index))
            return tdoc

        if tdoc not in df_tdocs.index:
            return tdoc

        try:
            revision_of = df_tdocs.at[tdoc, 'Revision of']
            merge_of = df_tdocs.at[tdoc, 'Merge of']
        except:
            # Case when the tdoc is not found
            print("Original TDoc: '{0}' not found. Stopping recursive search".format(tdoc))
            return tdoc

        # Check for circular reference (found one in the final TDocsByAgenda of 137E)
        if revision_of == original_index:
            revision_of = ''
        if merge_of == original_index:
            merge_of = ''

        if revision_of == '' and merge_of == '':
            return tdoc

        # Merge a list of "Revision of and Merge of" and execute recursively
        if revision_of == '' or merge_of == '':
            all_parents = revision_of + merge_of
        else:
            all_parents = ', '.join([revision_of, merge_of])

        return sort_and_remove_duplicates_from_list(
            TdocsByAgendaData.get_original_tdocs(all_parents, df_tdocs, original_index, n_recursion + 1))

    # Given a TDoc, returns the TDoc or TDocs that ultimately originate from this TDoc
    def get_final_tdocs(tdocs, df_tdocs, original_index, n_recursion):
        tdocs_split = tdocs.split(',')
        if len(tdocs_split) > 1:
            tdocs_split = [e for e in tdocs_split if (e != '') and (e is not None)]
        if len(tdocs_split) > 1:
            return join_results(tdocs_split, df_tdocs, TdocsByAgendaData.get_final_tdocs, original_index, n_recursion)

        # We know that length is 1
        tdoc = tdocs_split[0].strip()

        # Fix for 137E final TDocsByAgenda
        if n_recursion > max_recursion:
            print('Maximum recursion reached ({0}) for {1}. Stopping search.'.format(max_recursion, original_index))
            return tdoc

        if tdoc not in df_tdocs.index:
            tdoc = TdocsByAgendaData.try_to_correct_tdoc_typo(tdoc)

        try:
            revisions = df_tdocs.at[tdoc, 'Revised to']
            merges = df_tdocs.at[tdoc, 'Merged to']
        except:
            # Case when the tdoc is not found
            print("Final TDoc: '{0}' not found. Stopping recursive search".format(tdoc))
            return tdoc

        # This means that the TDoc has no more children
        if revisions == '' and merges == '':
            return tdoc

        # Merge a list of "Revised to and Merged to" and execute recursively
        if revisions == '' or merges == '':
            all_children = revisions + merges
        else:
            all_children = ', '.join([revisions, merges])

        return sort_and_remove_duplicates_from_list(
            TdocsByAgendaData.get_final_tdocs(all_children, df_tdocs, original_index, n_recursion + 1))

    def try_to_correct_tdoc_typo(tdoc):
        if tdoc not in TdocsByAgendaData.tdoc_typos.keys():
            return tdoc
        return TdocsByAgendaData.tdoc_typos[tdoc]


def get_tdocs_by_agenda_with_cache(path_or_html, meeting_server_folder='') -> TdocsByAgendaData:
    if (path_or_html is None) or (path_or_html == ''):
        print('Parse TDocsByAgenda skipped. path_or_html={0}, meeting_server_folder={1}'.format(path_or_html,
                                                                                                meeting_server_folder))
        return None

    global tdocs_by_document_cache

    print('Retrieving TDocsByAgenda (cache is enabled)')

    # If this is an HTML
    if len(path_or_html) > 1000:
        print('TDocsByAgenda retrieval based on HTML content')

        # Changed to hashlib as it is reinitialized beween sessions.
        # See https://stackoverflow.com/questions/27522626/hash-function-in-python-3-3-returns-different-results-between-sessions
        m = hashlib.md5()
        m.update(path_or_html)
        html_hash = m.hexdigest()

        # Retrieve
        if html_hash in tdocs_by_document_cache:
            print('Retrieving TdocsByAgenda from parsed document cache: {0}'.format(html_hash))
            last_tdocs_by_agenda = tdocs_by_document_cache[html_hash]
        else:
            print('TdocsByAgenda {0} not in cache'.format(html_hash))
            last_tdocs_by_agenda = TdocsByAgendaData(
                path_or_html,
                html_hash=html_hash,
                meeting_server_folder=meeting_server_folder)

            # I found out tht this was not a good idea for the inbox. File cache should be enough
            # print('Storing TdocsByAgenda with hash {0} in memory cache'.format(html_hash))
            # tdocs_by_document_cache[html_hash] = last_tdocs_by_agenda

            # Save TDocsByAgenda data in a pickle file so that we can plot graphs later on
            try:
                data_to_save = {
                    'contributor_columns': last_tdocs_by_agenda.contributor_columns,
                    'others_cosigners': last_tdocs_by_agenda.others_cosigners,
                    'tdocs': last_tdocs_by_agenda.tdocs,
                    'cache_version': current_cache_version
                }

                cache_file_name = get_cache_filepath(meeting_server_folder, html_hash)
                if cache_file_name is not None and not os.path.exists(cache_file_name):
                    with open(cache_file_name, 'wb') as f:
                        # Pickle the 'data' dictionary using the highest protocol available.
                        pickle.dump(data_to_save, f, pickle.HIGHEST_PROTOCOL)
                        print('Saved TDocsByAgenda cache to file {0}'.format(cache_file_name))
            except:
                print('Could not cache TDocsByAgenda for meeting {0}'.format(meeting_server_folder))
                print('Object to serialize:')
                print(data_to_save)
                traceback.print_exc()
    else:
        # Path-based fetching uses no hash
        print('TDocsByAgenda retrieval based on path')
        the_tdocs_by_agenda = TdocsByAgendaData(path_or_html)
        last_tdocs_by_agenda = the_tdocs_by_agenda

    return last_tdocs_by_agenda

# Storing all of them is easier, and the cache should not grow that big in the end
tdocs_by_document_cache = {}