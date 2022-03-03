import os
import re
import collections
import traceback
from docx import Document
from enum import Enum

import application.word
import server.tdoc
import config.contributor_names as contributor_names
import os.path
from datetime import datetime

from application.word import get_word, open_word_document
from tdoc.utils import tdoc_regex

title_regex = re.compile(r'Title:[\s\n]*(?P<title>.*)[\s\n]*\n', re.MULTILINE)
source_regex = re.compile(r'Source:[\s\n]*(?P<source>.*)[\s\n]*\n', re.MULTILINE)
cr_regex = re.compile(r'Title:[\s\n]*(?P<title>.*)[\n]*Source to WG:[\s\n]*(?P<source>.*)[\s\n]*Source to TSG',
                      re.MULTILINE)

color_magenta = (234, 10, 142)
color_black = (0, 0, 0)
color_white = (255, 255, 255)
color_green = (0, 97, 0)
color_light_green = (198, 239, 206)
color_light_gray = (242, 242, 242)
color_light_yellow = (255, 242, 204)

# See https://docs.microsoft.com/en-us/office/vba/api/Word.WdBuiltinStyle
toc_section_style = -2  # 'Überschrift 1'
source_section_style = -2  # 'Überschrift 1'
source_subsection_style = -3  # 'Überschrift 2'
tdoc_list_section_style = -2  # 'Überschrift 1'
tdoc_list_ai_section_style = -3  # 'Überschrift 2'
tdoc_list_ai_subsection_style = -4  # 'Überschrift 3'
standard_style = -1  # 'Standard'

Tdoc = collections.namedtuple('TDoc', 'title source')
TdocStats = collections.namedtuple('TdocStats',
                                   'tdoc_count tdoc_handled_count result_agreed_count result_revised_count result_noted_count')


# https://stackoverflow.com/questions/11444207/setting-a-cells-fill-rgb-color-with-pywin32-in-excel
def rgb_to_hex(rgb):
    '''
    ws.Cells(1, i).Interior.color uses bgr in hex

    '''
    bgr = (rgb[2], rgb[1], rgb[0])
    strValue = '%02x%02x%02x' % bgr
    # print(strValue)
    iValue = int(strValue, 16)
    return iValue


def get_metadata_from_doc(doc):
    if doc is None:
        return Tdoc(title=None, source=None)
    try:
        starting_text = ''
        tdoc_is_cr = False
        # 2000 character search range due to CRs
        # https://docs.microsoft.com/en-us/office/vba/api/word.document.range
        # https://docs.microsoft.com/en-us/office/vba/api/word.range.text
        try:
            starting_text = doc.Range(Start=0, End=2000).Text.replace(u'\r\x07', '\n').replace(u'\n\x07', '\n')
        except:
            try:
                starting_text = doc.Range(Start=0, End=1000).Text.replace(u'\r\x07', '\n').replace(u'\n\x07', '\n')
            except:
                starting_text = doc.Range(Start=0, End=500).Text.replace(u'\r\x07', '\n').replace(u'\n\x07', '\n')

        # Case for CRs
        if ('CHANGE REQUEST' in starting_text) and ('http://www.3gpp.org/Change-Requests' in starting_text):
            # starting_text = doc.Range(Start=0, End=2000).Text.replace(u'\r\x07', '\n').replace(u'\n\x07', '\n')
            tdoc_is_cr = True

        # Replace non-breaking spaces
        starting_text = starting_text.replace(u'\xa0', u' ').replace('\r', '\n')

        # print(starting_text)
        # doc.Close()
    finally:
        # word.Quit()
        pass

    title = None
    source = None

    # https://stackoverflow.com/questions/32134396/python-regular-expression-caret-not-working-in-multiline-modes
    if starting_text == '':
        return Tdoc(title=None, source=None)

    if not tdoc_is_cr:
        title_match = title_regex.search(starting_text)
        if title_match:
            title = title_match.groupdict()['title'].strip()

        source_match = source_regex.search(starting_text)
        if source_match:
            source = source_match.groupdict()['source'].strip()
    else:
        cr_match = cr_regex.search(starting_text)
        if cr_match:
            title = cr_match.groupdict()['title'].strip()
            source = cr_match.groupdict()['source'].strip()

    return Tdoc(title=title, source=source)


def parse_document(filename):
    doc = open_word_document(filename)
    return get_metadata_from_doc(doc)


def insert_text_and_format(doc, text, style, old_style, insert_range=None):
    if insert_range is None:
        insert_range = doc.Content
    original_start = insert_range.Start
    original_end = insert_range.End
    print('Text range before insertion: {0}-{1}'.format(original_start, original_end))
    insert_range.InsertAfter(text)
    modified_start = insert_range.Start
    modified_end = insert_range.End
    print('Text range after insertion: {0}-{1}'.format(insert_range.Start, insert_range.End))

    start_difference = modified_start - original_start
    end_different = modified_end - modified_end

    # Format title and undo formatting for rest
    if style is not None:
        # Move range to be in modified_range
        insert_range.MoveStart(1, start_difference)
        insert_range.MoveEnd(1, end_different)

        # Apply style
        insert_range.Style = style

        # Move back to original
        insert_range.MoveStart(1, -start_difference)
        insert_range.MoveEnd(1, -end_different)

    # Move to end of selection
    # 0=wdCollapseEnd, see https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdcollapsedirection?view=word-pia
    insert_range.Collapse(0)
    if old_style is not None:
        insert_range.Style = old_style

    # Return new position
    return insert_range


class TDocType(Enum):
    LS = 2
    CR = 1
    ALL = 3
    CR_AND_NOTED_DISCUSSION = 4
    WID_NEW = 5


def merge_cells(criteria_column, cols_to_merge, table):
    # Merges cells on columns based on the specified criteria column

    # Avoid addressing errors by changing first the furthest cells
    cols_to_merge.sort(reverse=True)

    last_cell_was_empty = False
    if len(criteria_column) > 0:
        cell_list_to_traverse = list(enumerate(criteria_column))
        empty_cell_start = cell_list_to_traverse[-1][0]
        if (criteria_column is not None) and (cols_to_merge != []):
            for idx, cell in reversed(cell_list_to_traverse):
                cell_text = cell.Range.Text
                current_cell_row = cell.RowIndex
                if (cell_text == '') or (cell_text == '\r\x07'):
                    if not last_cell_was_empty:
                        # Change the starting merge cell range only if we are the starting one
                        empty_cell_start = current_cell_row
                    last_cell_was_empty = True
                else:
                    if (idx != empty_cell_start) and last_cell_was_empty:
                        print('Merging cells from rows {0} to {1}'.format(current_cell_row, empty_cell_start))
                        for col_idx in cols_to_merge:
                            start_merge = table.Cell(Row=current_cell_row, Column=col_idx + 1)
                            end_merge = table.Cell(Row=empty_cell_start, Column=col_idx + 1)
                            start_merge.Merge(end_merge)
                    last_cell_was_empty = False


def format_table(table):
    # Formatting
    table.Rows.AllowBreakAcrossPages = False
    table.Borders.Enable = True
    # Line styles https://docs.microsoft.com/en-us/office/vba/api/word.wdlinestyle
    table.Borders.InsideLineStyle = 1
    table.Borders.OutsideLineStyle = 1
    # Line Widths https://docs.microsoft.com/en-us/office/vba/api/word.WdLineWidth
    table.Borders.InsideLineWidth = 8
    table.Borders.OutsideLineWidth = 8
    # Line color https://docs.microsoft.com/en-us/office/vba/api/word.WdColor
    table.Borders.InsideColor = 0
    table.Borders.OutsideColor = 0

    # Header cell
    header_row = table.Rows[0]
    header_row.Range.Font.Bold = True
    header_row.HeadingFormat = True

    # Vertical alignment
    for idx, row in enumerate(table.Rows):
        # wdCellAlignVerticalCente=1. See https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdcellverticalalignment?view=word-pia
        row.Cells.VerticalAlignment = 1
        row.Range.ParagraphFormat.SpaceBefore = 0
        row.Range.ParagraphFormat.SpaceAfter = 0


def fill_in_table(
        doc,
        df,
        type,
        meeting_folder,
        insert_range=None,
        title_style='Überschrift 2',
        standard_style='Standard',
        status_to_show=[],
        status_to_ignore=[],
        show_comments=True,
        show_statistics=True):
    if (df is None) or (doc is None) or (len(df) == 0):
        return insert_range

    # Filter only wanted status and generate some statistics
    df_original = df
    df = df.copy()
    if (status_to_show is not None) and (len(status_to_show) > 0):
        df = df[df['Result'].isin(status_to_show)]
    elif (status_to_ignore is not None) and (len(status_to_ignore) > 0):
        # More robust evaluation
        lowercase_results = df['Result'].str.lower()
        lowercase_ignores = [e.lower() for e in status_to_ignore]
        df = df[~lowercase_results.isin(lowercase_ignores)]

    # Check length after filtering
    if len(df) == 0:
        return insert_range

    if type == TDocType.CR_AND_NOTED_DISCUSSION:
        type = TDocType.CR
        add_noted_disc_to_crs = True
    else:
        add_noted_disc_to_crs = False

    use_blocks = False
    if type == TDocType.LS:
        if show_comments:
            columns = ['Type', 'Title', 'Source', 'Rel', 'Comments', 'Result']
        else:
            columns = ['Type', 'Title', 'Source', 'Rel', 'Result']
        df = df[df['Type'].str.contains("LS")]
        table_title = 'Liasons\n'
        use_blocks = True
    elif type == TDocType.CR:
        if show_comments:
            columns = ['Title', 'Source', 'Work Item', 'Comments', 'Result']
        else:
            columns = ['Title', 'Source', 'Work Item', 'Result']
        df = df[df['Type'].str.contains("CR")]
        if add_noted_disc_to_crs:
            noted_disc_tdocs_idx = df_original['Type'].str.contains("DISCUSSION") & df_original['Result'].str.contains(
                'Noted')
            noted_disc_tdocs = df_original[noted_disc_tdocs_idx]
            if len(noted_disc_tdocs) > 0:
                # Put the discussion papers first
                noted_disc_tdocs['Result'] = 'Noted (Discussion)'
                df = noted_disc_tdocs.append(df)

        table_title = 'CRs\n'
        use_blocks = True
    elif type == TDocType.WID_NEW:
        if show_comments:
            columns = ['Title', 'Source', 'Work Item', 'Comments', 'Result']
        else:
            columns = ['Title', 'Source', 'Work Item', 'Result']
        df = df[df['Type'].str.contains("WID NEW")]
        table_title = 'New WIDs\n'
        use_blocks = True
    else:
        columns = df.columns
        table_title = 'All TDocs\n'

    if len(df) == 0:
        return insert_range

    print('Generating Word table report for {0}, {1} TDocs. Position: {2}-{3}'.format(type, len(df), insert_range.Start,
                                                                                      insert_range.End))

    if use_blocks:
        current_block = ''
        current_block_start_row = 2
        current_block_started = False

    # Table title
    insert_range = insert_text_and_format(doc, table_title, title_style, standard_style, insert_range=insert_range)

    if show_statistics:
        stats_str, df_stats = get_tdoc_statistics(df, type=type)
        insert_range = insert_text_and_format(doc, stats_str, standard_style, standard_style, insert_range=insert_range)

    # 6=wdStory
    # References: https://docs.microsoft.com/en-us/office/vba/api/word.wdunits, 
    # https://software-solutions-online.com/word-vba-move-cursor-to-end-of-document/
    if insert_range is None:
        cursor = doc.ActiveWindow.Selection.EndKey(6)
        insert_range = doc.ActiveWindow.Selection.Range
    table = doc.Tables.Add(insert_range, NumRows=len(df) + 1, NumColumns=len(columns) + 1)

    # Add column names
    table.Cell(Row=1, Column=1).Range.Text = 'TD#'
    for idx, col_name in enumerate(columns):
        table.Cell(Row=1, Column=idx + 2).Range.Text = col_name

    current_row = table.Rows.Last

    tdocs = df.index.tolist()
    server_urls = dict([(tdoc, server.tdoc.get_remote_filename(meeting_folder, tdoc, use_inbox=False)) for tdoc in tdocs])

    # Fill in TDoc data
    row_idx = 2
    last_cr = ''
    last_cr_sources = ''

    # Rows that will be kept empty for later merging if the merge criteria is met, i.e.
    # For CRs: same CR number
    # For LSs: same title
    rows_to_ignore_if_merged = ['CR', 'Title', 'Rel', 'Type', 'Work Item']
    if type == TDocType.LS:
        rows_to_ignore_if_merged.append('Source')

    for idx, row in df.iterrows():
        tdoc_index_cell = table.Cell(Row=row_idx, Column=1).Range
        tdoc_index_cell.Text = idx

        doc.Hyperlinks.Add(tdoc_index_cell, server_urls[idx])
        tdoc_index_cell.Font.Color = rgb_to_hex(color_magenta)

        # Check whether we need to skip some values due to cell merging
        cells_to_merge = False
        if use_blocks:
            if type == TDocType.CR:
                merge_criteria = row['CR']
            else:
                merge_criteria = row['Title']
            if merge_criteria != '':
                if merge_criteria != current_block:
                    current_block = merge_criteria
                    current_block_start_row = row_idx
                    current_block_start_row = True
                else:
                    cells_to_merge = True
        for col_idx, value in enumerate(columns):
            # Skip cells meant for merging
            if cells_to_merge and (value in rows_to_ignore_if_merged):
                continue

            cell_content = row[value]

            # Summarized comments column
            if value == 'Comments':
                cell_content = ''
                strings_to_merge = []
                if row['Revised to']:
                    strings_to_merge.append('Revised to: {0}'.format(row['Revised to']))
                if row['Merged to']:
                    strings_to_merge.append('Merged to: {0}'.format(row['Merged to']))
                if len(strings_to_merge) > 0:
                    cell_content = '\n'.join(strings_to_merge)

                    table.Cell(Row=row_idx, Column=col_idx + 2).Range.Text = cell_content
                if cell_content != '':
                    # Add TDoc links to added text
                    cell_range = table.Cell(Row=row_idx, Column=col_idx + 2).Range
                    current_cell_text = cell_range.Text
                    found_tdocs = tdoc_regex.finditer(current_cell_text)
                    content_length = len(current_cell_text)
                    for m in found_tdocs:
                        m_start = m.start(0)
                        m_end = m.end(0)
                        m_tdoc = m.group(0)
                        try:
                            try:
                                tdoc_url = server_urls[m_tdoc]
                            except:
                                # May not be in this set of URLs. Note that the URL *may* not exist if it is from another meeting!
                                server_urls[m_tdoc] = server.tdoc.get_remote_filename(
                                    meeting_folder,
                                    m_tdoc,
                                    use_inbox=False)

                                tdoc_url = server_urls[m_tdoc]
                            tdoc_range = table.Cell(Row=row_idx, Column=col_idx + 2).Range
                            range_start = tdoc_range.Start
                            range_end = tdoc_range.End
                            # wdCharacter=1 https://docs.microsoft.com/en-us/office/vba/api/word.wdunits
                            tdoc_range.MoveStart(1, m_start)
                            tdoc_range.MoveEnd(1, m_end - content_length + 1)
                            doc.Hyperlinks.Add(tdoc_range, tdoc_url)
                            tdoc_range.Font.Color = rgb_to_hex(color_magenta)
                        except:
                            print('Could not add link for {0}'.format(m_tdoc))
                            traceback.print_exc()
            elif value == 'Source':
                cr_number = row['CR']
                if cr_number == '':
                    # Not a CR, just copy the source list
                    table.Cell(Row=row_idx, Column=col_idx + 2).Range.Text = cell_content
                else:
                    if cr_number == last_cr:
                        # Print only the diff
                        table.Cell(Row=row_idx, Column=col_idx + 2).Range.Text = diff_sources(last_cr_sources,
                                                                                              cell_content)
                    else:
                        # New CR, print full sources
                        table.Cell(Row=row_idx, Column=col_idx + 2).Range.Text = cell_content
            else:
                table.Cell(Row=row_idx, Column=col_idx + 2).Range.Text = cell_content

            # Formatting for Results column (CRs)
            if value == 'Result':
                if re.search('approved', cell_content, re.IGNORECASE) or re.search('agreed', cell_content,
                                                                                   re.IGNORECASE) or re.search(
                        'replied to', cell_content, re.IGNORECASE):
                    table.Cell(Row=row_idx, Column=col_idx + 2).Shading.BackgroundPatternColor = rgb_to_hex(
                        color_light_green)
                elif re.search('revised', cell_content, re.IGNORECASE):
                    table.Cell(Row=row_idx, Column=col_idx + 2).Shading.BackgroundPatternColor = rgb_to_hex(
                        color_light_gray)
                elif re.search('noted', cell_content, re.IGNORECASE) or re.search('postponed', cell_content,
                                                                                  re.IGNORECASE):
                    table.Cell(Row=row_idx, Column=col_idx + 2).Shading.BackgroundPatternColor = rgb_to_hex(
                        color_light_yellow)
        # End of row processing
        last_cr = row['CR']
        last_cr_sources = row['Source']
        row_idx = row_idx + 1

    print('Formatting table')

    # Formatting
    format_table(table)

    # Merge cells
    criteria_column = []
    # Zero-indexed!! Not like Word!
    cols_to_merge = []
    if type == TDocType.LS:
        # Merge Title and Rel columns
        criteria_column = list(table.Columns[2].Cells)
        cols_to_merge = [4, 3, 2, 1]
    elif (type == TDocType.CR) or (type == TDocType.WID_NEW):
        # Merge Title and Work Item column
        criteria_column = list(table.Columns[1].Cells)
        cols_to_merge = [3, 1]

    merge_cells(criteria_column, cols_to_merge, table)

    # Additional merge for CR sources, as they may contain a diff
    if type == TDocType.CR:
        criteria_column = list(table.Columns[2].Cells)
        cols_to_merge = [2]
        merge_cells(criteria_column, cols_to_merge, table)

    # Column Auto-fit
    if (type == TDocType.LS) or (type == TDocType.WID_NEW):
        table.Columns(2).AutoFit()  # Type
        table.Columns(5).AutoFit()  # Rel
        if show_comments:
            table.Columns(6).AutoFit()  # Result
        else:
            table.Columns(5).AutoFit()  # Result
        table.Columns(3).AutoFit()  # Title
    if type == TDocType.CR:
        table.Columns(1).AutoFit()  # TD#
        if show_comments:
            table.Columns(5).AutoFit()  # Result
        else:
            table.Columns(4).AutoFit()  # Result
        table.Columns(2).AutoFit()  # Title

    # Table looks better with a carriage return at the end
    end_of_table = table.Range
    end_of_table.Collapse(0)
    end_of_section = insert_text_and_format(doc, '\n', standard_style, standard_style, insert_range=end_of_table)
    print('Finished generating TDoc report table. Table end: {0}-{1}'.format(end_of_section.Start, end_of_section.End))
    return end_of_section


def add_ai_to_wi(wi, df):
    idx = df.loc[df['Work Item'] == wi, 'Work Item'].index[0]
    wi_plus_ai = '{0} {1}'.format(df.at[idx, 'AI'], wi)
    return wi_plus_ai


def insert_index_at_begin(doc):
    # Insert TOC in new page
    document_start = doc.Content
    document_start.Collapse(1)

    insert_range = insert_text_and_format(doc, 'Index\n', toc_section_style, toc_section_style,
                                          insert_range=document_start)
    doc.TablesOfContents.Add(
        insert_range,
        UseHeadingStyles=True,
        UseFields=True,
        UpperHeadingLevel=1,
        LowerHeadingLevel=2)
    toc_end = doc.TablesOfContents[0].Range
    toc_end.Collapse(0)
    # wdPageBreak=7 https://docs.microsoft.com/en-us/office/vba/api/word.wdbreaktype
    toc_end.InsertBreak(7)


def get_tdoc_statistics(df, type=TDocType.ALL, show_noted_discussion_tdocs_with_crs=False):
    tdoc_count = 0
    tdoc_handled_count = 0
    result_agreed_count = 0
    result_revised_count = 0
    result_noted_count = 0

    is_cr = (type == TDocType.CR) | (type == TDocType.CR_AND_NOTED_DISCUSSION)

    if (df is None) or (len(df) == 0):
        return '', TdocStats(tdoc_count, tdoc_handled_count, result_agreed_count, result_revised_count,
                             result_noted_count)

    lowercase_results = df['Result'].str.lower()

    tdoc_count = len(df)
    tdoc_handled_count = tdoc_count - len(lowercase_results[lowercase_results.str.contains('withdrawn')]) - len(
        lowercase_results[lowercase_results.str.contains('not handled')]) - len(
        lowercase_results[lowercase_results == ''])
    result_agreed_count = len(lowercase_results[lowercase_results.str.contains('agreed')])
    result_revised_count = len(lowercase_results[lowercase_results.str.contains('revised')])
    result_noted_count = len(lowercase_results[lowercase_results.str.contains('noted')])
    cr_count = len([cr for cr in df['CR'].unique() if (cr is not None) and (cr != '')])

    # Assume that DataFrame is not empty
    stats_str = '{0:,} TDocs'.format(tdoc_count)

    # Add text only if needed
    if is_cr:
        stats_str = stats_str + ' ({:,} CRs)'.format(cr_count)

    suffixes = []
    if tdoc_handled_count > 0:
        suffixes.append(', {0:,} handled ({1:.1%})'.format(tdoc_handled_count, tdoc_handled_count / tdoc_count))
    if result_agreed_count > 0:
        suffixes.append(', {0:,} agreed ({1:.1%})'.format(result_agreed_count, result_agreed_count / tdoc_count))
    if result_revised_count > 0:
        suffixes.append(', {0:,} revised ({1:.1%})'.format(result_revised_count, result_revised_count / tdoc_count))
    if result_noted_count > 0:
        suffixes.append(', {0:,} noted ({1:.1%})'.format(result_noted_count, result_noted_count / tdoc_count))

    if len(suffixes) == 0:
        suffixes = ['']
    if len(suffixes) > 0:
        suffixes[0] = suffixes[0].replace(',', '.')
    if len(suffixes) > 1:
        suffixes[-1] = suffixes[-1].replace(',', ' and')
    stats_str = stats_str + ''.join(suffixes) + '\n'

    return stats_str, TdocStats(tdoc_count, tdoc_handled_count, result_agreed_count, result_revised_count,
                                result_noted_count)


def insert_doc_data_to_doc(
        df,  # Data source (Pandas DataFrame)
        doc,  # Word document
        meeting_folder,  # Used to generate the linksto the documents
        add_toc=True,  # Whether to add a ToC at the beginning of the document
        insert_range=None,  # Where to insert the tables/text
        section_title=None,  # Title of the section (e.g. "Full Contribution Summary")
        sort_by_wi=False,  # Whether the DataFrame should be sorted by Work Item OR only based on TDoc number
        title_style='Überschrift 1',  # Word style to use for the title
        subtitle_style='Überschrift 2',  # Word style to use for the subtitle
        status_to_show=[],
        # Show only TDocs with the given status. Use "None" or [] to ignore this option. Note that status_to_show has
        # precedence over status_to_ignore
        status_to_ignore=[],
        # Do not show TDocs with the given status. Use "None" or [] to ignore this option. Note that status_to_show
        # has precedence over status_to_ignore
        show_comments=True,  # Whether to show the "Comments" column for CRs. For a more compact view it can be ignored
        show_withdrawn_crs=True,  # If set to False, adds CRs with status 'Withdrawn' to 'status_to_ignore'
        show_noted_discussion_tdocs_with_crs=False,  # Whether Noted discussion CRs should be shown with CRs
        show_noted_lss=True,  # If set to True, removes status 'Noted' from 'status_to_ignore'
        show_statistics=True):  # If set to True, shown a statistics entry

    if (df is None) or (doc is None) or (len(df) == 0):
        return insert_range

    # Check if there is something to output. We output only CRs and LSs
    number_of_LS = len(df[df['Type'].str.contains("LS")].index)
    number_of_CR = len(df[df['Type'].str.contains("CR")].index)
    number_of_WID_NEW = len(df[df['Type'].str.contains("WID NEW")].index)

    if number_of_LS + number_of_CR + number_of_WID_NEW < 1:
        print('No CRs/LSs to output for section {0}'.format(section_title))
        return insert_range

    # Sort DataFrame so that all CR revisions are together
    if sort_by_wi:
        df = df.sort_values(by=['Work Item', 'Type', 'CR', 'TD#'], ascending=True)
    else:
        df = df.sort_values(by=['Type', 'CR', 'TD#'], ascending=True)

    if section_title is None:
        # Unique values of Work Item for title
        work_items = [e for e in df['Work Item'].unique().tolist() if (e is not None) and (e != '')]
        work_items = [add_ai_to_wi(e, df) for e in work_items]
        section_title = '; '.join(work_items)

    insert_range = insert_text_and_format(doc, section_title + '\n', title_style, standard_style,
                                          insert_range=insert_range)

    if show_statistics:
        stats_str, df_stats = get_tdoc_statistics(df)
        insert_range = insert_text_and_format(doc, stats_str, standard_style, standard_style, insert_range=insert_range)

    ignore_list_for_lss = status_to_ignore.copy()
    if show_noted_lss:
        ignore_list_for_lss = [s for s in ignore_list_for_lss if s != 'Noted']

    # Add LSs
    insert_range = fill_in_table(
        doc, df,
        TDocType.LS,
        meeting_folder,
        insert_range=insert_range,
        title_style=subtitle_style,
        standard_style=standard_style,
        status_to_show=status_to_show,
        status_to_ignore=ignore_list_for_lss,
        show_comments=show_comments,
        show_statistics=show_statistics)

    ignore_list_for_crs = status_to_ignore.copy()
    if not show_withdrawn_crs:
        ignore_list_for_crs.append('Withdrawn')

    if not show_noted_discussion_tdocs_with_crs:
        cr_type = TDocType.CR
    else:
        cr_type = TDocType.CR_AND_NOTED_DISCUSSION

    # Add CRs
    insert_range = fill_in_table(
        doc, df,
        cr_type,
        meeting_folder,
        insert_range=insert_range,
        title_style=subtitle_style,
        standard_style=standard_style,
        status_to_show=status_to_show,
        status_to_ignore=ignore_list_for_crs,
        show_comments=show_comments,
        show_statistics=show_statistics)

    # Add new WIDs
    insert_range = fill_in_table(
        doc, df,
        TDocType.WID_NEW,
        meeting_folder,
        insert_range=insert_range,
        title_style=subtitle_style,
        standard_style=standard_style,
        status_to_show=status_to_show,
        status_to_ignore=ignore_list_for_lss,
        show_comments=show_comments,
        show_statistics=show_statistics)

    if add_toc:
        insert_index_at_begin(doc)

    return insert_range


def filter_wi_list(wis_list):
    if wis_list is None or len(wis_list) < 1:
        return ''

    wis_list = [wid for wid in wis_list if (wid is not None) and (wid != '')]
    if len(wis_list) < 1:
        return ''

    joined_list = ','.join(wis_list)
    wis = [wi.strip() for wi in joined_list.split(',')]
    wis = list(set(wis))
    title = ', '.join(wis)
    return title


def insert_cr_summary_to_report(
        df,
        doc,
        contributor_ranking_count=20,
        insert_range=None,
        source=None,
        status_to_ignore=[]):
    if df is None:
        return insert_range

    section_title = 'Contribution Summary\n'
    insert_range = insert_text_and_format(doc, section_title, source_section_style, standard_style,
                                          insert_range=insert_range)

    # Only CRs and no revisions
    # df_cr_only = df[df['Type'].str.contains('CR')]
    # df_filtered = df_cr_only[~df_cr_only['Result'].str.contains('Revised')]
    # cr_count = len(df_filtered['CR'].unique())

    df_filtered = df.copy()
    all_contribution_count = len(df)

    contribution_count = []
    df_columns = df_filtered.columns
    for source_item, column in contributor_names.contributor_columns.items():
        if not column in df_columns:
            continue
        tdocs_for_source = df_filtered[column]
        if len(tdocs_for_source) == 0:
            continue
        source_count = tdocs_for_source.sum()
        contribution_count.append((source_item, source_count))
    contribution_count.sort(key=lambda x: x[1], reverse=True)
    contribution_count = [(item[0], item[1], idx) for idx, item in enumerate(contribution_count)]

    if len(contribution_count) < contributor_ranking_count:
        contribution_count_limited = contribution_count
    else:
        contribution_count_limited = contribution_count[0:contributor_ranking_count]

    all_in_ranking = [i[0] for i in contribution_count_limited]
    if (source is not None) and (source not in all_in_ranking):
        add_rows_for_source = True
        table_rows = len(contribution_count_limited) + 2
    else:
        add_rows_for_source = False
        table_rows = len(contribution_count_limited) + 1

    ignored_contributions = ''
    if (status_to_ignore is not None) and (len(status_to_ignore) > 0):
        ignored_contributions = ' Not showing {0} contributions in contribution summary sections below'.format(
            ', '.join(status_to_ignore))

    if source is None:
        legend_str = '{0:,} TDocs total.{1}\nTop {2:,} contributor list:\n'.format(all_contribution_count,
                                                                                   ignored_contributions,
                                                                                   contributor_ranking_count)
    else:
        legend_str = '{0:,} TDocs total.{1}\nTop {2:,} contributor list plus {3}:\n'.format(all_contribution_count,
                                                                                            ignored_contributions,
                                                                                            contributor_ranking_count,
                                                                                            source)
    insert_range = insert_text_and_format(
        doc,
        legend_str,
        standard_style,
        standard_style,
        insert_range=insert_range)

    # Insert contributor table
    if len(contribution_count_limited) == 0:
        return insert_range

    table = doc.Tables.Add(insert_range, NumRows=table_rows, NumColumns=3)
    table.Cell(Row=1, Column=1).Range.Text = '#'
    table.Cell(Row=1, Column=2).Range.Text = 'Company'
    table.Cell(Row=1, Column=3).Range.Text = 'Contributions'

    for idx, item in enumerate(contribution_count_limited):
        table.Cell(Row=idx + 2, Column=1).Range.Text = '{0:,}'.format(idx + 1)
        table.Cell(Row=idx + 2, Column=2).Range.Text = item[0]
        if item[0] == source:
            table.Cell(Row=idx + 2, Column=2).Range.Bold = True
        table.Cell(Row=idx + 2, Column=3).Range.Text = '{0:,} ({1:.1%})'.format(item[1],
                                                                                item[1] / all_contribution_count)
    if add_rows_for_source:
        candidate_item = [e for e in contribution_count if e[0] == source]
        try:
            item = candidate_item[0]
            table.Cell(Row=table_rows, Column=1).Range.Text = '{0:,}'.format(item[2] + 1)
            table.Cell(Row=table_rows, Column=2).Range.Text = source
            table.Cell(Row=table_rows, Column=2).Range.Bold = True
            table.Cell(Row=table_rows, Column=3).Range.Text = '{0:,} ({1:.1%})'.format(item[1],
                                                                                       item[1] / all_contribution_count)
        except:
            print('Could not filter out contributions from {0}'.format(source))
    # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdrowalignment?view=word-pia
    table.Rows.Alignment = 1  # wdAlignRowCenter
    table.Columns(1).AutoFit()
    table.Columns(3).AutoFit()
    format_table(table)

    # Table looks better with a carriage return at the end
    end_of_table = table.Range
    end_of_table.Collapse(0)
    end_of_section = insert_text_and_format(doc, '\n', standard_style, standard_style, insert_range=end_of_table)
    return end_of_section


def insert_tdocs_for_specific_source(
        df,
        doc,
        meeting_folder,
        insert_range=None,
        source=None):
    # Source can be one of the keys specified in contributor_names (e.g. 'DT')
    if source is None:
        return insert_range
    try:
        column_name = contributor_names.contributor_columns[source]
        df_source_tdocs = df[df[column_name]]
    except:
        print('Could not find range {0}'.format(source))
        return insert_range

    insert_range = insert_doc_data_to_doc(
        df_source_tdocs,
        doc,
        meeting_folder,
        add_toc=False,
        insert_range=insert_range,
        section_title='{0} Contributions'.format(source),
        sort_by_wi=True,
        title_style=source_section_style,
        subtitle_style=source_subsection_style)
    return insert_range


def insert_doc_data_to_doc_by_wi(
        df,
        doc,
        meeting_folder,
        ais_to_skip=[],
        ais_to_output=None,
        source=None,
        reduced_full_summary=True,
        show_comments_for_full_summary=False,
        insert_cr_summary=True,
        save_to_folder=None):
    if (df is None) or (doc is None):
        return

    # Dieter could not decide on whether he really needs this or not
    show_noted_discussion_tdocs_with_crs = False
    show_noted_lss = False
    show_statistics = False
    if reduced_full_summary:
        status_to_ignore = ['Revised', 'Merged', 'Not Handled', 'Noted', 'Withdrawn']
    else:
        status_to_ignore = []

    if save_to_folder is not None:
        now = datetime.now()
        report_filename = os.path.join(
            save_to_folder,
            '{0:0>4d}.{1:0>2d}.{2:0>2d} {3:0>2d}{4:0>2d}{5:0>2d} {6}.docx'.format(
                now.year,
                now.month,
                now.day,
                now.hour,
                now.minute,
                now.second,
                meeting_folder))
        print('Saving Word report as {0}'.format(report_filename))
        doc.SaveAs(report_filename)

    # Disable spell checking, as it may cause problems if the document grows large (which it will probably do)
    doc.SpellingChecked = True
    doc.GrammarChecked = True
    doc.ShowSpellingErrors = False

    agenda_items = [e for e in df['AI'].unique().tolist() if (e is not None) and (e != '')]
    if ais_to_output is not None:
        agenda_items = [ai for ai in agenda_items if ai in ais_to_output]

    insert_range = doc.Content
    insert_range.Collapse(1)

    if insert_cr_summary:
        insert_range = insert_cr_summary_to_report(
            df,
            doc,
            insert_range=insert_range,
            source=source,
            status_to_ignore=status_to_ignore)

    # Insert specific TDocs (e.g. own TDocs)
    if source is not None:
        insert_range = insert_tdocs_for_specific_source(
            df, doc,
            meeting_folder,
            insert_range=insert_range,
            source=source)

    # Insert TDoc summary
    insert_range = insert_text_and_format(
        doc,
        'Full Contribution Summary\n',
        tdoc_list_section_style,
        standard_style,
        insert_range=insert_range)

    stats_str, df_stats = get_tdoc_statistics(df)
    insert_range = insert_text_and_format(doc, stats_str, standard_style, standard_style, insert_range=insert_range)

    if meeting_folder in server.tdoc.ai_names_cache:
        agenda_description = server.tdoc.ai_names_cache[meeting_folder]
    else:
        agenda_description = {}

    for idx, ai in enumerate(agenda_items):
        if ai in ais_to_skip:
            print('Skipping AI {0}'.format(ai))
            continue

        print('AI {0} {1} of {2}'.format(ai, idx, len(agenda_items)))
        filtered_df = df[df['AI'] == ai]

        if ai not in agenda_description:
            ai_description = filtered_df['Work Item'].unique().tolist()
            ai_description = filter_wi_list(ai_description)
        else:
            ai_description = agenda_description[ai]
        insert_range = insert_doc_data_to_doc(
            filtered_df,
            doc, meeting_folder,
            add_toc=False,
            insert_range=insert_range,
            section_title='{0} {1}'.format(ai, ai_description),
            title_style=tdoc_list_ai_section_style,
            subtitle_style=tdoc_list_ai_subsection_style,
            status_to_ignore=status_to_ignore,
            show_comments=show_comments_for_full_summary,
            show_withdrawn_crs=False,
            show_noted_discussion_tdocs_with_crs=show_noted_discussion_tdocs_with_crs,
            show_noted_lss=show_noted_lss,
            show_statistics=show_statistics)
    # Insert ToC at beginning of document after everything is finished
    insert_index_at_begin(doc)

    if save_to_folder is not None:
        doc.Save()


def diff_sources(source1, source2):
    if source1 is None:
        source1 = ''
    if source2 is None:
        source2 = ''
    if source1 == '':
        return '+ {0}'.format(source2)
    if source1 == source2:
        return ''
    elif source2.find(source1) != -1:
        raw_diff = source2.replace(source1, '')
        sources = raw_diff.split(',')
        sources = [source.strip() for source in sources if (source is not None) and (source != '')]
        sources = ', '.join(sources)
        return '+ {0}'.format(sources)
    else:
        return source2


def import_agenda(agenda_file):
    try:
        document = Document(agenda_file)
    except:
        print('Could not open file {0}'.format(agenda_file))
        return None

    all_tables = document.tables

    # First table is the topic list
    topics_table = all_tables[0]

    # Table including "Lunch" string is the agenda
    agenda_table = None
    for table in all_tables[1:-1]:
        try:
            if len(table.columns) < 5:
                # Meetings are five days, so a narrower table is not valid
                continue
            lunch_cells = [cell for row in table.rows for cell in row.cells if
                           (cell.text is not None) and ('Lunch' in cell.text)]
            if len(lunch_cells) > 0:
                # Found the agenda table
                agenda_table = table
                break
        except:
            traceback.print_exc()
            pass

    # Parse AIs
    agenda_table_parsed = [(row.cells[0].text, row.cells[1].text) for row in topics_table.rows]
    agenda_table_parsed = [(entry[0].replace('\n', '').strip(), entry[1]) for entry in agenda_table_parsed if
                           (entry[0] != '') and (entry[0] != 'AI#')]
    agenda_table_parsed = [(entry[0], entry[1].split('\n')[0]) for entry in agenda_table_parsed]
    ai_descriptions = dict(agenda_table_parsed)

    days = [cell.text.split(' ')[0] for cell in agenda_table.row_cells(0)][1:]
    hours_column = [cell.text for cell in agenda_table.column_cells(0)][1:]

    # Add room name
    last_hour = None
    repetitions = 0
    room_name = lambda x: 'Main room' if x == 0 else 'Breakout {0}'.format(x)
    for idx, current_hour in enumerate(hours_column):
        if last_hour == current_hour:
            repetitions += 1
            new_hour = '{0} ({1})'.format(current_hour, room_name(repetitions))
        else:
            repetitions = 0
            if 'Lunch' not in current_hour:
                new_hour = '{0} ({1})'.format(current_hour, room_name(repetitions))
            else:
                new_hour = current_hour
        last_hour = current_hour
        hours_column[idx] = new_hour

    return ai_descriptions


def compare_documents(tdoc_1, tdoc_2):
    if tdoc_1 is None or tdoc_1 == '' or tdoc_2 is None or tdoc_2 == '':
        print('Empty or None tdoc_input')
        return
    try:
        word_application = get_word()
        print('Comparing {0} and {1}'.format(tdoc_1, tdoc_2))
        doc_1 = open_word_document(filename=tdoc_1, set_as_active_document=False)
        doc_2 = open_word_document(filename=tdoc_2, set_as_active_document=False)

        # Call Word's compare feature
        # Destination=wdCompareDestinationNew (see https://docs.microsoft.com/en-us/office/vba/api/word.wdcomparedestination)
        comparison_document = word_application.CompareDocuments(OriginalDocument=doc_1, RevisedDocument=doc_2, Destination=2, IgnoreAllComparisonWarnings=True)
    except:
        print('Could not compare documents')
        traceback.print_exc()


def open_file(file, return_metadata=False, go_to_page=1):
    # No metadata
    if return_metadata:
        return application.word.open_file(file, go_to_page=go_to_page)

    # Case returning metadata
    return application.word.open_file(file, go_to_page=go_to_page, metadata_function=get_metadata_from_doc)


def open_files(files, return_metadata=False, go_to_page=1):
    # No metadata
    if return_metadata:
        return application.word.open_files(files, go_to_page=go_to_page)

    # Case returning metadata
    return application.word.open_files(files, go_to_page=go_to_page, metadata_function=get_metadata_from_doc)


def write_data_and_open_file(data, local_file, open_this_file=True):
    if data is None:
        return
    with open(local_file, 'wb') as output:
        output.write(data)
    if open_this_file:
        open_file(local_file)
