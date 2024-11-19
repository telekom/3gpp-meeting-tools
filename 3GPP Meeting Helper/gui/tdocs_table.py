import os
import re
import textwrap
import tkinter
import traceback
import webbrowser
from tkinter import ttk
from typing import Callable, Tuple, Any

import pandas as pd
import pyperclip

import application
import application.meeting_helper
import application.tkinter_config
import gui
import gui.common.common_elements
import gui.main_gui
import parsing.word.pywin32
import utils.local_cache
from application import powerpoint
from application.excel import open_excel_document, set_first_row_as_filter, vertically_center_all_text, save_wb, \
    set_column_width, set_wrap_text, hide_column
from gui.common.generic_table import GenericTable, treeview_sort_column, treeview_set_row_formatting
from parsing.html.revisions import revisions_file_to_dataframe
from parsing.html.tdocs_by_agenda import TdocsByAgendaData
from parsing.outlook_utils import search_subject_in_all_outlook_items
from parsing.word.pywin32 import parse_list_of_crs


class TdocsTable(GenericTable):
    current_tdocs = None
    source_width = 200
    title_width = 550

    meeting_number = '<Meeting number>'
    all_tdocs = None
    meeting_server_folder = ''

    def __init__(
            self,
            favicon,
            parent_widget: tkinter.Tk,
            meeting_name: str,
            meeting_server_folder: str,
            update_tdocs_by_agenda_fn: Callable[[], TdocsByAgendaData | None],
            get_current_meeting_name_fn: Callable[[], str],
            download_and_open_tdoc_fn=None,
            download_and_open_generic_tdoc_fn: Callable[[str], Tuple[Any, Any]] | None = None

    ):
        """
        Opens the TDoc table
        Args:
            favicon: The favicon to use
            parent_widget: The parent widget
            meeting_name: The meeting names (reader-friendly string)
            meeting_server_folder: The folder name in the 3GPP server
            update_tdocs_by_agenda_fn: A function that updates the TDocsByAgenda data in application.meeting_helper
            get_current_meeting_name_fn: Retrieves the currently selected meeting (needed for choosing between the
            generic and non-generic Tdoc retrieval functions
            download_and_open_tdoc_fn: Retrieves a TDoc from the currently-selected meeting
            download_and_open_generic_tdoc_fn: Retrieves a TDoc from any meeting (generic TDoc)
        """
        super().__init__(
            parent_widget=parent_widget,
            widget_title=f"{meeting_name} TDocs",
            favicon=favicon,
            column_names=[
                'TDoc',
                'AI', 'Type',
                'Title',
                'Source',
                'Revs',
                'Emails',
                'Send @',
                'Result'],
            row_height=60,
            display_rows=9,
            root_widget=None
        )
        # Functions to update data from the main GUI
        self.meeting_server_folder = meeting_server_folder
        self.get_current_meeting_name_fn = get_current_meeting_name_fn
        self.download_and_open_tdoc_fn = download_and_open_tdoc_fn

        # Used to check if we have the current meeting selected or not
        self.meeting_name = meeting_name
        self.update_tdocs_by_agenda_fn = update_tdocs_by_agenda_fn

        # If we have another meeting selected, a generic TDoc search is performed
        self.download_and_open_generic_tdoc_fn = download_and_open_generic_tdoc_fn

        self.tdoc_count = tkinter.StringVar()
        self.revisions_list = None
        self.revisions = None

        self.set_column('TDoc', "TDoc #", width=110)
        self.set_column('AI', width=50)
        self.set_column('Type', width=120)
        self.set_column('Title', width=TdocsTable.title_width, center=False)
        self.set_column('Source', width=TdocsTable.source_width, center=False)
        self.set_column('Revs', width=50)
        self.set_column('Emails', width=50)
        self.set_column('Send @', width=50, sort=False)
        self.set_column('Result', width=100)

        self.tree.bind("<Double-Button-1>", self.on_double_click)

        self.load_data(reload=True, reload_ais=False)
        self.reload_revisions = False
        self.insert_current_tdocs()

        # Can also do this:
        # https://stackoverflow.com/questions/33781047/tkinter-drop-down-list-of-check-boxes-combo-boxes
        self.search_text = tkinter.StringVar()
        self.search_entry = tkinter.Entry(
            self.top_frame,
            textvariable=self.search_text,
            width=25,
            font='TkDefaultFont')
        self.search_text.trace_add(['write', 'unset'], self.select_text)

        ttk.Label(self.top_frame, text="Search: ").pack(side=tkinter.LEFT)
        self.search_entry.pack(side=tkinter.LEFT)

        all_types = ['All']
        all_types.extend(list(self.current_tdocs["Type"].unique()))
        self.combo_type = ttk.Combobox(
            self.top_frame,
            values=all_types,
            state="readonly",
            width=20)
        self.combo_type.set('All')
        self.combo_type.bind("<<ComboboxSelected>>", self.select_type)

        all_ais = ['All']
        all_ais.extend(list(self.current_tdocs["AI"].unique()))
        self.combo_ai = ttk.Combobox(
            self.top_frame,
            values=all_ais,
            state="readonly",
            width=7)
        self.combo_ai.set('All')
        self.combo_ai.bind("<<ComboboxSelected>>", self.select_ai)

        all_results = ['All']
        all_results.extend(list(self.current_tdocs["Result"].unique()))
        self.combo_result = ttk.Combobox(
            self.top_frame,
            values=all_results,
            state="readonly",
            width=12)
        self.combo_result.set('All')
        self.combo_result.bind("<<ComboboxSelected>>", self.select_result)

        ttk.Label(self.top_frame, text="  By Type: ").pack(side=tkinter.LEFT)
        self.combo_type.pack(side=tkinter.LEFT)

        ttk.Label(self.top_frame, text="  By AI: ").pack(side=tkinter.LEFT)
        self.combo_ai.pack(side=tkinter.LEFT)

        ttk.Label(self.top_frame, text="  By Result: ").pack(side=tkinter.LEFT)
        self.combo_result.pack(side=tkinter.LEFT)

        ttk.Label(self.top_frame, text="  ").pack(side=tkinter.LEFT)
        ttk.Button(
            self.top_frame,
            text='Clear filters',
            command=self.clear_filters).pack(side=tkinter.LEFT)
        ttk.Button(
            self.top_frame,
            text='Reload data',
            command=self.reload_data).pack(side=tkinter.LEFT)

        ttk.Button(
            self.top_frame,
            text='Merge PPTs',
            command=self.merge_pptx_files).pack(side=tkinter.LEFT)

        ttk.Button(
            self.top_frame,
            text='Export CRs',
            command=self.export_crs).pack(side=tkinter.LEFT)

        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

        ttk.Label(self.bottom_frame, textvariable=self.tdoc_count).pack(side=tkinter.LEFT)

        # Add text wrapping
        # https: // stackoverflow.com / questions / 51131812 / wrap - text - inside - row - in -tkinter - treeview

    def retrieve_current_tdocs_by_agenda(self):
        """
        Calls retrieve_current_tdocs_by_agenda_fn and updates all_tdocs variable with the retrieved data
        """
        if self.update_tdocs_by_agenda_fn is not None:
            try:
                self.update_tdocs_by_agenda_fn()
                current_tdocs_by_agenda = application.meeting_helper.current_tdocs_by_agenda
                self.all_tdocs = current_tdocs_by_agenda.tdocs
                self.meeting_number = current_tdocs_by_agenda.meeting_number
                self.meeting_server_folder = current_tdocs_by_agenda.meeting_server_folder
                print('Loaded meeting {0}, server folder {1}'.format(self.meeting_number, self.meeting_server_folder))
            except Exception as e:
                print(f'Could not retrieve current TdocsByAgenda for Tdocs table: {e}')
                traceback.print_exc()

    def get_tdocs_by_agenda(self) -> TdocsByAgendaData | None:
        if self.update_tdocs_by_agenda_fn is not None:
            try:
                return self.update_tdocs_by_agenda_fn()
            except Exception as e:
                print(f'Could not get TdocsByAgenda, Drafts, Revisions for Tdocs table: {e}')
                traceback.print_exc()
                return None
        else:
            return None

    def download_and_open_tdoc(self, tdoc_to_open, skip_opening=False):
        # Case when we are searching documents for the currently selected meeting
        if self.selected_meeting_is_this_one:
            print(f'Opening TDoc {tdoc_to_open} of this meeting ({self.meeting_name})')
            if self.download_and_open_tdoc_fn is None:
                return None
            try:
                return self.download_and_open_tdoc_fn(
                    tdoc_to_open,
                    copy_to_clipboard=True,
                    skip_opening=skip_opening,
                    open_tdoc_details_for_global_search=False
                )
            except Exception as e:
                print(f'Could not open TDoc {tdoc_to_open} for Tdocs table: {e}')
                traceback.print_exc()
                return None

        # Case when we are searching documents for another meeting
        print(f'Opening TDoc {tdoc_to_open} of another meeting (not {self.meeting_name})')
        if self.download_and_open_generic_tdoc_fn is None:
            return None
        return self.download_and_open_generic_tdoc_fn(tdoc_to_open)

    def load_data(self, reload=False, reload_ais=True):
        if reload:
            print('Loading revision data for table')

            # Re-load TdocsByAgenda before inserting rows
            self.retrieve_current_tdocs_by_agenda()

            meeting_server_folder = self.meeting_server_folder
            print(f'Meeting server folder is {meeting_server_folder}')
            self.get_tdocs_by_agenda()

            self.revisions, self.revisions_list = revisions_file_to_dataframe(
                revisions_file=utils.local_cache.get_local_revisions_filename(meeting_server_folder),
                meeting_tdocs=self.current_tdocs,
                drafts_file=utils.local_cache.get_local_drafts_filename(meeting_server_folder))

        # Rewrite the current tdocs dataframe with the retrieved data. Resets the search filters
        self.current_tdocs = self.all_tdocs

        # Update AI Combo Box
        if reload_ais:
            all_ais = ['All']
            all_ais.extend(list(self.current_tdocs["AI"].unique()))
            self.combo_ai['values'] = all_ais

    def insert_current_tdocs(self):
        self.insert_rows(self.current_tdocs)

    def insert_rows(self, df):
        count = 0

        for idx, row in df.iterrows():
            count = count + 1
            mod = count % 2
            if mod > 0:
                tag = 'odd'
            else:
                tag = 'even'

            if self.revisions is None:
                revision_count = ''
            else:
                try:
                    rev_number = self.revisions.loc[idx, 'Revisions']
                    try:
                        rev_number_converted = int(rev_number.replace('*', ''))
                    except Exception as e:
                        print(f'Could not convert revision number to int. Set to 0: {e}')
                        rev_number_converted = 0
                    if rev_number_converted < 1:
                        revision_count = ''
                    else:
                        revision_count = rev_number
                except KeyError:
                    # Not found
                    revision_count = ''  # Zero is left empty
                    pass
                except Exception as e:
                    print(f'Could not set revision count. Sent to empty: {e}')
                    revision_count = ''  # Error is left empty
                    traceback.print_exc()

            self.tree.insert("", "end", tags=(tag,), values=(
                idx,
                row['AI'],
                row['Type'],
                textwrap.fill(row['Title'], width=70),
                textwrap.fill(row['Source'], width=25),
                revision_count,
                'Click',
                'Click',
                row['Result']))

        treeview_set_row_formatting(self.tree)
        self.tdoc_count.set('{0} documents'.format(count))

    def clear_filters(self, *args):
        self.combo_type.set('All')
        self.combo_ai.set('All')
        self.combo_result.set('All')
        self.search_text.set('')
        self.load_data(reload=False)
        self.select_ai(load_data=True)  # One will call the other(s)

    def reload_data(self, *args):
        self.load_data(reload=True)
        self.select_ai()  # One will call the other

    def merge_pptx_files(self, *args):
        print('Extracting all current TDocs and merge PowerPoint files (used to merge status report presentations)')
        print('Current Tdocs:')
        tdoc_list_to_merge = list(self.current_tdocs.index)
        print(tdoc_list_to_merge)
        all_extracted_files = []
        all_titles = []
        for tdoc_id in tdoc_list_to_merge:
            extracted_files = self.download_and_open_tdoc(tdoc_id, skip_opening=True)
            if extracted_files is not None:
                try:
                    all_extracted_files.extend(extracted_files)
                    all_titles.append(self.current_tdocs.at[tdoc_id, 'Title'])
                except Exception as e:
                    print(f'Could not iterate output from {tdoc_id}, extracted files={extracted_files}: {e}')

        all_extracted_files = [e for e in all_extracted_files if '.ppt' in e.lower()]
        print('Opened PowerPoint files:')
        print(all_extracted_files)
        powerpoint.merge_presentations(
            all_extracted_files,
            list_of_section_labels=tdoc_list_to_merge,
            headlines_for_toc=all_titles)

    def export_crs(self, *args):
        """
        From the current Tdoc list, exports the current CRs to an Excel file
        Args:
            *args:

        Returns: Nothing

        """

        # Generate a list of CR files to parse based on the information in the TdocsByAgenda file
        tdoc_list = self.current_tdocs
        tdocs_to_export = tdoc_list[tdoc_list['Type'] == 'CR']
        if len(tdocs_to_export) == 0:
            return

        # Generate list containing the TDoc number and the AI
        tdocs_to_export = zip(tdocs_to_export.index.values.tolist(), tdocs_to_export['AI'].values.tolist())
        file_path_list = []
        for tdoc_to_export in tdocs_to_export:
            try:
                tdoc_path = self.download_and_open_tdoc(tdoc_to_export[0], skip_opening=True)
            except Exception as e:
                print(f"Could not retrieve file path for {tdoc_to_export}: {e}")
                tdoc_path = None
            # Take by default the first file

            if tdoc_path is None:
                # Some files may not be available
                continue

            # Contains the first file in the TDoc's zip file, the AI and the TDoc number
            file_path_list.append((tdoc_path[0], tdoc_to_export[1], tdoc_to_export[0]))

        print("Will export {0} CRs".format(len(file_path_list)))
        # print(file_path_list)

        selected_meeting = gui.common.common_elements.tkvar_meeting.get()

        # Generate output filename for the CR summary Excel
        server_folder = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(
            selected_meeting)
        agenda_folder = utils.local_cache.get_local_agenda_folder(server_folder)
        current_dt_str = application.meeting_helper.get_now_time_str()
        excel_export_filename = os.path.join(agenda_folder, '{0} {1}.xlsx'.format(current_dt_str, 'CR_export'))

        # The actual parsing of the CRs. Returns a DataFrame object containing the CR data
        crs_df = parse_list_of_crs(file_path_list)
        crs_df = crs_df.set_index('TDoc')

        # Avoid IllegalCharacterError due to some control characters
        # https://stackoverflow.com/questions/28837057/pandas-writing-an-excel-file-containing-unicode-illegalcharactererror
        crs_df.to_excel(excel_export_filename, sheet_name="CRs", engine='xlsxwriter')

        print("Opening {0}".format(excel_export_filename))
        wb = open_excel_document(excel_export_filename)
        set_first_row_as_filter(wb)
        vertically_center_all_text(wb)
        set_wrap_text(wb)
        set_column_width('A', wb, 11)
        set_column_width('B', wb, 9)
        set_column_width('C', wb, 9)
        set_column_width('D', wb, 9)
        set_column_width('E', wb, 9)
        set_column_width('F', wb, 20)
        set_column_width('J', wb, 7)
        set_column_width('G', wb, 20)
        hide_column('H', wb)
        set_column_width('K', wb, 11)
        set_column_width('N', wb, 11)
        set_column_width('O', wb, 8)
        set_column_width('P', wb, 8)
        set_column_width('Q', wb, 8)
        set_column_width('R', wb, 75)
        set_column_width('S', wb, 75)
        set_column_width('T', wb, 75)
        set_column_width('U', wb, 11)
        save_wb(wb)

        return

    def select_ai(self, load_data=True, event=None):
        if load_data:
            self.load_data()

        tdocs_for_ai = self.current_tdocs
        selected_ai = self.combo_ai.get()
        print('Filtering by AI "{0}"'.format(selected_ai))
        if selected_ai == 'All':
            tdocs_for_ai = tdocs_for_ai
        else:
            tdocs_for_ai = tdocs_for_ai[tdocs_for_ai['AI'] == self.combo_ai.get()]

        self.current_tdocs = tdocs_for_ai

        super().clear_tree()
        self.insert_current_tdocs()

        if load_data:
            self.select_text(load_data=False)
            self.select_result(load_data=False)
            self.select_type(load_data=False)

    def select_type(self, load_data=True, event=None):
        if load_data:
            self.load_data()

        tdocs_for_type = self.current_tdocs
        selected_type = self.combo_type.get()
        print('Filtering by Type "{0}"'.format(selected_type))
        if selected_type == 'All':
            tdocs_for_type = tdocs_for_type
        else:
            tdocs_for_type = tdocs_for_type[tdocs_for_type['Type'] == self.combo_type.get()]

        self.current_tdocs = tdocs_for_type

        self.tree.delete(*self.tree.get_children())
        self.insert_current_tdocs()

        if load_data:
            self.select_ai(load_data=False)
            self.select_text(load_data=False)
            self.select_result(load_data=False)

    def select_result(self, load_data=True, event=None):
        if load_data:
            self.load_data()

        tdocs_for_result = self.current_tdocs
        selected_result = self.combo_result.get()
        print('Filtering by Result "{0}"'.format(selected_result))
        if selected_result == 'All':
            tdocs_for_result = tdocs_for_result
        else:
            tdocs_for_result = tdocs_for_result[tdocs_for_result['Result'] == self.combo_result.get()]

        self.current_tdocs = tdocs_for_result

        self.tree.delete(*self.tree.get_children())
        self.insert_current_tdocs()

        if load_data:
            self.select_text(load_data=False)
            self.select_ai(load_data=False)
            self.select_type(load_data=False)

    def select_text(self, load_data=True, *args):
        if load_data:
            self.load_data()

        # Filter based on current TDocs
        text_search = self.search_text.get()
        if text_search is None or text_search == '':
            return

        try:
            re.compile(text_search)
            is_regex = True
            print('Filtering by Regex "{0}"'.format(text_search))
        except re.error:
            is_regex = False
            print('Filtering by Text "{0}"'.format(text_search))

        text_search = text_search.lower()
        tdocs_for_text = self.current_tdocs.copy()
        tdocs_for_text['search_column'] = tdocs_for_text.index + tdocs_for_text['Title'] + tdocs_for_text['Source']
        tdocs_for_text['search_column'] = tdocs_for_text['search_column'].str.lower()
        tdocs_for_text = tdocs_for_text[tdocs_for_text['search_column'].str.contains(text_search, regex=is_regex)]
        self.current_tdocs = tdocs_for_text

        self.tree.delete(*self.tree.get_children())
        self.insert_current_tdocs()

        if load_data:
            self.select_ai(load_data=False)
            self.select_result(load_data=False)
            self.select_type(load_data=False)

    def on_double_click(self, event):
        item_id = self.tree.identify("item", event.x, event.y)
        column = int(self.tree.identify_column(event.x)[1:]) - 1  # "#1" -> 0 (one-indexed in TCL)
        item_values = self.tree.item(item_id)['values']
        try:
            actual_value = item_values[column]
        except Exception as e:
            print(f'Could not parse actual value: {e}')
            actual_value = None
        tdoc_id = item_values[0]
        print("you clicked on {0}/{1}: {2}".format(event.x, event.y, actual_value))
        if actual_value is None or actual_value == '':
            print("Empty value")
            return
        if column == 0:
            print('Opening {0}'.format(actual_value))
            self.download_and_open_tdoc(actual_value)
        if column == 5:
            print('Opening revisions for {0}'.format(tdoc_id))
            RevisionsTable(
                self.tk_top,
                self.favicon,
                tdoc_id,
                self.revisions_list,
                parent_tdocs_table=self)
        if column == 6:
            print('Opening emails for {0}'.format(tdoc_id))
            search_subject_in_all_outlook_items(tdoc_id)
        if column == 7:
            print(
                'Generating subject for email approval for {0}. Copying to clipboard and generating empty email'.format(
                    tdoc_id))
            subject = '[SA2#{3}, AI#{1}, {0}] {2}'.format(tdoc_id, item_values[1], item_values[3], self.meeting_number)
            subject = subject.replace('\n', ' ').replace('  ', ' ')
            print(subject)
            webbrowser.open('mailto:{0}?subject={1}'.format('3GPP_TSG_SA_WG2_EMEET@LIST.ETSI.ORG', subject), new=1)
            pyperclip.copy(subject)

    @property
    def selected_meeting_is_this_one(self):
        if self.get_current_meeting_name_fn is None:
            return True
        current_meeting: str = self.get_current_meeting_name_fn()
        return self.meeting_name == current_meeting


class RevisionsTable(GenericTable):

    def __init__(
            self,
            parent_widget: tkinter.Tk,
            favicon: str,
            tdoc_id,
            revisions_df,
            parent_tdocs_table):

        super().__init__(
            parent_widget=parent_widget,
            favicon=favicon,
            widget_title="Revisions for {0}".format(tdoc_id),
            root_widget=None,
            column_names=['TDoc', 'Rev.', 'Add to compare A', 'Add to compare B']
        )

        revisions = revisions_df.loc[tdoc_id, :]
        self.tdoc_id = tdoc_id
        self.parent_tdocs_table = parent_tdocs_table
        print('{0} Revisions'.format(len(revisions)))

        self.count = 0

        self.compare_a = tkinter.StringVar()
        self.compare_b = tkinter.StringVar()

        self.set_column('TDoc', "TDoc #", width=110, center=True)
        self.set_column('Rev.', width=50, center=True)
        self.set_column('Add to compare A', width=110, center=True)
        self.set_column('Add to compare B', width=110, center=True)

        self.tree.bind("<Double-Button-1>", self.on_double_click)

        ttk.Label(self.bottom_frame, textvariable=self.compare_a).pack(side=tkinter.LEFT)
        ttk.Label(self.bottom_frame, text='  vs.  ').pack(side=tkinter.LEFT)
        ttk.Label(self.bottom_frame, textvariable=self.compare_b).pack(side=tkinter.LEFT)
        ttk.Label(self.bottom_frame, text='  ').pack(side=tkinter.LEFT)

        ttk.Button(
            self.bottom_frame,
            text='Compare!',
            command=self.compare_tdocs).pack(side=tkinter.LEFT)

        # Main frame
        self.insert_rows(revisions)
        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

    def insert_rows(self, df):
        count = 0
        if df is None:
            return

        if isinstance(df, pd.Series):
            rows = [(self.tdoc_id, df)]
        else:
            rows = df.iterrows()

        for idx, row in rows:
            if count == 0:
                count = count + 1
                mod = count % 2
                if mod > 0:
                    tag = 'odd'
                else:
                    tag = 'even'
                self.tree.insert("", "end", tags=(tag,), values=(
                    idx,
                    '00',
                    'Click',
                    'Click'))

            count = count + 1
            mod = count % 2
            if mod > 0:
                tag = 'odd'
            else:
                tag = 'even'

            values = (
                idx,
                row['Revisions'],
                'Click',
                'Click')
            self.tree.insert("", "end", tags=(tag,), values=values)

        treeview_sort_column(self.tree, 'Rev.')

        self.count = count
        treeview_set_row_formatting(self.tree)

    def on_double_click(self, event):
        item_id = self.tree.identify("item", event.x, event.y)
        column = int(self.tree.identify_column(event.x)[1:]) - 1  # "#1" -> 0 (one-indexed in TCL)
        item_values = self.tree.item(item_id)['values']
        try:
            actual_value = item_values[column]
        except Exception as e:
            print(f'Could not parse actual value for column {column}: {e}')
            actual_value = None

        # Some issues with automatic conversion which we solve here
        tdoc_id = item_values[0]
        if isinstance(item_values[1], int):
            revision = 'r' + '{0:02d}'.format(item_values[1])
        else:
            revision = 'r' + item_values[1]

        if revision == 'r00':
            tdoc_to_search = tdoc_id
        else:
            tdoc_to_search = tdoc_id + revision
        print("you clicked on {0}/{1}: {2}".format(event.x, event.y, actual_value))

        match column:
            case 0:
                print('Opening {0}'.format(actual_value))
                self.parent_tdocs_table.download_and_open_tdoc(actual_value)
            case 1:
                print('Opening {0}'.format(tdoc_to_search))
                self.parent_tdocs_table.download_and_open_tdoc(tdoc_to_search)
            case 2:
                self.compare_a.set(tdoc_to_search)
            case 3:
                self.compare_b.set(tdoc_to_search)

    def compare_tdocs(self):
        compare_a = self.compare_a.get()
        compare_b = self.compare_b.get()
        print('Comparing {0} vs. {1}'.format(compare_a, compare_b))
        parsing.word.pywin32.compare_tdocs(
            entry_1=compare_a,
            entry_2=compare_b)
