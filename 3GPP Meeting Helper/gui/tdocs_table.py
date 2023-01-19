import os
import re
import textwrap
import tkinter
import traceback
import webbrowser
from tkinter import ttk

import pandas as pd
import pyperclip

import application
import gui
import server
from application import powerpoint
from application.excel import open_excel_document, set_first_row_as_filter, vertically_center_all_text, save_wb, \
    set_column_width, set_wrap_text, hide_column
from parsing.html.revisions import revisions_file_to_dataframe
from parsing.outlook_utils import search_subject_in_all_outlook_items
from parsing.word.pywin32 import parse_list_of_crs

style_name = 'mystyle.Treeview'


# See https://bugs.python.org/issue36468
def fixed_map(option):
    # Fix for setting text colour for Tkinter 8.6.9
    # From: https://core.tcl.tk/tk/info/509cafafae
    #
    # Returns the style map for 'option' with any styles starting with
    # ('!disabled', '!selected', ...) filtered out.

    # style.map() returns an empty list for missing options, so this
    # should be future-safe.
    return [elm for elm in style.map(style_name, query_opt=option) if
            elm[:2] != ('!disabled', '!selected')]


style = None


def init_style():
    global style
    if style is None:
        style = ttk.Style()
        style.map(style_name, foreground=fixed_map('foreground'),
                  background=fixed_map('background'))
        style.configure(style_name, highlightthickness=0, bd=0, rowheight=60)
        # style.configure("mystyle.Treeview.Heading", font=('Calibri', 13, 'bold'))  # Modify the font of the headings
        style.layout(style_name, [(style_name + '.treearea', {'sticky': 'nswe'})])  # Remove the borders


def set_column(tree, col, label=None, width=None, sort=True, center=True):
    if label is None:
        label = col
    if sort:
        tree.heading(col, text=label, command=lambda: treeview_sort_column(tree, col, False))
    else:
        tree.heading(col, text=label)
    if width is not None:
        tree.column(col, minwidth=0, width=width, stretch=False)
    if center:
        tree.column(col, anchor="center")


def treeview_sort_column(tree, col, reverse=False):
    l = [(tree.set(k, col), k) for k in tree.get_children('')]
    l.sort(reverse=reverse)

    # rearrange items in sorted positions
    for index, (val, k) in enumerate(l):
        tree.move(k, '', index)

    # reverse sort next time
    tree.heading(col, command=lambda: treeview_sort_column(tree, col, not reverse))


class TdocsTable:
    current_tdocs = None
    source_width = 200
    title_width = 550

    meeting_number = '<Meeting number>'
    all_tdocs = None
    meeting_server_folder = ''

    def __init__(
            self,
            parent,
            favicon,
            parent_gui_tools,
            retrieve_current_tdocs_by_agenda_fn=None,
            get_tdocs_by_agenda_for_selected_meeting_fn=None,
            download_and_open_tdoc_fn=None):

        init_style()
        self.top = tkinter.Toplevel(parent)
        self.top.title("TDoc Table for current meeting. Double-Click on TDoc # or revision # to open")
        self.top.iconbitmap(favicon)
        self.parent_gui_tools = parent_gui_tools
        self.favicon = favicon

        # Functions to update data from the main GUI
        self.retrieve_current_tdocs_by_agenda_fn = retrieve_current_tdocs_by_agenda_fn
        self.get_tdocs_by_agenda_for_selected_meeting_fn = get_tdocs_by_agenda_for_selected_meeting_fn
        self.download_and_open_tdoc_fn = download_and_open_tdoc_fn

        frame_1 = tkinter.Frame(self.top)
        frame_1.pack(anchor='w')
        frame_2 = tkinter.Frame(self.top)
        frame_2.pack()
        frame_3 = tkinter.Frame(self.top)
        frame_3.pack(anchor='w')

        self.tdoc_count = tkinter.StringVar()
        self.revisions_list = None
        self.revisions = None

        # https://stackoverflow.com/questions/50625306/what-is-the-best-way-to-show-data-in-a-table-in-tkinter
        self.tree = ttk.Treeview(
            frame_2,
            columns=('TDoc', 'AI', 'Type', 'Title', 'Source', 'Revs', 'Emails', 'Send @', 'Result'),
            show='headings',
            selectmode="browse",
            style=style_name,
            padding=[-5, -25, -5, -25])  # Left, top, right, bottom

        set_column(self.tree, 'TDoc', "TDoc #", width=110)
        set_column(self.tree, 'AI', width=50)
        set_column(self.tree, 'Type', width=120)
        set_column(self.tree, 'Title', width=TdocsTable.title_width, center=False)
        set_column(self.tree, 'Source', width=TdocsTable.source_width, center=False)
        set_column(self.tree, 'Revs', width=50)
        set_column(self.tree, 'Emails', width=50)
        set_column(self.tree, 'Send @', width=50, sort=False)
        set_column(self.tree, 'Result', width=100)

        self.tree.bind("<Double-Button-1>", self.on_double_click)

        self.load_data(reload=True, reload_ais=False)
        self.reload_revisions = False
        self.insert_current_tdocs()

        self.tree_scroll = ttk.Scrollbar(frame_2)
        self.tree_scroll.configure(command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.tree_scroll.set)
        # tree.grid(row=0, column=0)

        # Can also do this:
        # https://stackoverflow.com/questions/33781047/tkinter-drop-down-list-of-check-boxes-combo-boxes
        self.search_text = tkinter.StringVar()
        self.search_entry = tkinter.Entry(frame_1, textvariable=self.search_text, width=25, font='TkDefaultFont')
        self.search_text.trace_add(['write', 'unset'], self.select_text)

        tkinter.Label(frame_1, text="Search: ").pack(side=tkinter.LEFT)
        self.search_entry.pack(side=tkinter.LEFT)

        all_types = ['All']
        all_types.extend(list(self.current_tdocs["Type"].unique()))
        self.combo_type = ttk.Combobox(frame_1, values=all_types, state="readonly")
        self.combo_type.set('All')
        self.combo_type.bind("<<ComboboxSelected>>", self.select_type)

        all_ais = ['All']
        all_ais.extend(list(self.current_tdocs["AI"].unique()))
        self.combo_ai = ttk.Combobox(frame_1, values=all_ais, state="readonly", width=10)
        self.combo_ai.set('All')
        self.combo_ai.bind("<<ComboboxSelected>>", self.select_ai)

        all_results = ['All']
        all_results.extend(list(self.current_tdocs["Result"].unique()))
        self.combo_result = ttk.Combobox(frame_1, values=all_results, state="readonly")
        self.combo_result.set('All')
        self.combo_result.bind("<<ComboboxSelected>>", self.select_result)

        tkinter.Label(frame_1, text="  By Type: ").pack(side=tkinter.LEFT)
        self.combo_type.pack(side=tkinter.LEFT)

        tkinter.Label(frame_1, text="  By AI: ").pack(side=tkinter.LEFT)
        self.combo_ai.pack(side=tkinter.LEFT)

        tkinter.Label(frame_1, text="  By Result: ").pack(side=tkinter.LEFT)
        self.combo_result.pack(side=tkinter.LEFT)

        tkinter.Label(frame_1, text="  ").pack(side=tkinter.LEFT)
        tkinter.Button(
            frame_1,
            text='Clear filters',
            command=self.clear_filters).pack(side=tkinter.LEFT)
        tkinter.Button(
            frame_1,
            text='Reload data',
            command=self.reload_data).pack(side=tkinter.LEFT)

        tkinter.Button(
            frame_1,
            text='Merge PPTs',
            command=self.merge_pptx_files).pack(side=tkinter.LEFT)

        tkinter.Button(
            frame_1,
            text='Export CRs',
            command=self.export_crs).pack(side=tkinter.LEFT)

        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

        tkinter.Label(frame_3, textvariable=self.tdoc_count).pack(side=tkinter.LEFT)

        # Add text wrapping
        # https: // stackoverflow.com / questions / 51131812 / wrap - text - inside - row - in -tkinter - treeview

    def retrieve_current_tdocs_by_agenda(self):
        """
        Calls retrieve_current_tdocs_by_agenda_fn and updates all_tdocs variable with the retrieved data
        """
        if self.retrieve_current_tdocs_by_agenda_fn is not None:
            try:
                current_tdocs_by_agenda = self.retrieve_current_tdocs_by_agenda_fn()
                self.all_tdocs = current_tdocs_by_agenda.tdocs
                self.meeting_number = current_tdocs_by_agenda.meeting_number
                self.meeting_server_folder = current_tdocs_by_agenda.meeting_server_folder
                print('Loaded meeting {0}, server folder {1}'.format(self.meeting_number, self.meeting_server_folder))
            except:
                print('Could not retrieve current TdocsByAgenda for Tdocs table')
                traceback.print_exc()

    def get_tdocs_by_agenda_for_selected_meeting(self, meeting_server_folder):
        if self.get_tdocs_by_agenda_for_selected_meeting_fn is not None:
            try:
                return self.get_tdocs_by_agenda_for_selected_meeting_fn(
                    meeting_server_folder,
                    return_revisions_file=True,
                    return_drafts_file=True)
            except:
                print('Could not get TdocsByAgenda, Drafts, Revisions for Tdocs table')
                traceback.print_exc()
                return None
        else:
            return None

    def download_and_open_tdoc(self, actual_value, skip_opening=False):
        if self.download_and_open_tdoc_fn is not None:
            try:
                return self.download_and_open_tdoc_fn(
                    actual_value, copy_to_clipboard=True, skip_opening=skip_opening)
            except:
                print('Could not open TDoc {0} for Tdocs table'.format(actual_value))
                traceback.print_exc()
                return None
        else:
            return None

    def load_data(self, reload=False, reload_ais=True):
        if reload:
            print('Loading revision data for table')

            # Re-load TdocsByAgenda before inserting rows
            self.retrieve_current_tdocs_by_agenda()

            meeting_server_folder = self.meeting_server_folder
            tdocs_by_agenda_file, revisions_file, drafts_file = self.get_tdocs_by_agenda_for_selected_meeting(
                meeting_server_folder)

            self.revisions, self.revisions_list = revisions_file_to_dataframe(
                revisions_file,
                self.current_tdocs,
                drafts_file=drafts_file)

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
                    except:
                        rev_number_converted = 0
                    if rev_number_converted < 1:
                        revision_count = ''
                    else:
                        revision_count = rev_number
                except KeyError:
                    # Not found
                    revision_count = ''  # Zero is left empty
                    pass
                except:
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

        self.tree.tag_configure('odd', background='#E8E8E8')
        self.tree.tag_configure('even', background='#DFDFDF')
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
                except:
                    print('Could not iterate output from {0}: {1}'.format(tdoc_id, extracted_files))

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
            except:
                print("Could not retrieve file path for {0}".format(tdoc_to_export))
                tdoc_path = None
            # Take by default the first file

            if tdoc_path is None:
                # Some files may not be available
                continue

            # Contains the first file in the TDoc's zip file, the AI and the TDoc number
            file_path_list.append((tdoc_path[0], tdoc_to_export[1], tdoc_to_export[0]))

        print("Will export {0} CRs".format(len(file_path_list)))
        # print(file_path_list)

        selected_meeting = gui.main.tkvar_meeting.get()

        # Generate output filename for the CR summary Excel
        server_folder = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(
            selected_meeting)
        agenda_folder = server.common.get_local_agenda_folder(server_folder)
        current_dt_str = application.meeting_helper.get_now_time_str()
        excel_export_filename = os.path.join(agenda_folder, '{0} {1}.xlsx'.format(current_dt_str, 'CR_export'))

        # The actual parsing of the CRs. Returns a DataFrame object containing the CR data
        crs_df = parse_list_of_crs(file_path_list)
        crs_df = crs_df.set_index('TDoc')

        # Avoid IllegalCharacterError due to some control characters
        # See https://stackoverflow.com/questions/28837057/pandas-writing-an-excel-file-containing-unicode-illegalcharactererror
        crs_df.to_excel(excel_export_filename, sheet_name="CRs", engine='xlsxwriter')

        # ToDo: Some formatting of the CR metadata

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

        self.tree.delete(*self.tree.get_children())
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
        except:
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
                self.top,
                self.favicon,
                tdoc_id,
                self.revisions_list,
                self.parent_gui_tools,
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


class RevisionsTable:

    def __init__(self, parent, favicon, tdoc_id, revisions_df, parent_gui_tools, parent_tdocs_table):
        top = self.top = tkinter.Toplevel(parent)
        top.title("Revisions for {0}".format(tdoc_id))
        top.iconbitmap(favicon)
        revisions = revisions_df.loc[tdoc_id, :]
        self.tdoc_id = tdoc_id
        self.parent_gui_tools = parent_gui_tools
        self.parent_tdocs_table = parent_tdocs_table
        print('{0} Revisions'.format(len(revisions)))

        frame_1 = tkinter.Frame(top)
        frame_1.pack()
        frame_2 = tkinter.Frame(top)
        frame_2.pack(anchor='w')
        frame_3 = tkinter.Frame(top)
        frame_3.pack(anchor='w')

        self.compare_a = tkinter.StringVar()
        self.compare_b = tkinter.StringVar()

        self.tree = ttk.Treeview(
            frame_1,
            columns=('TDoc', 'Rev.', 'Add to compare A', 'Add to compare B'),
            show='headings',
            selectmode="browse",
            style=style_name,
            height=8)  # Height in rows

        set_column(self.tree, 'TDoc', "TDoc #", width=110)
        set_column(self.tree, 'Rev.', width=50)
        set_column(self.tree, 'Add to compare A', width=110)
        set_column(self.tree, 'Add to compare B', width=110)
        self.tree.bind("<Double-Button-1>", self.on_double_click)

        self.count = 0
        self.insert_rows(revisions)

        self.tree_scroll = ttk.Scrollbar(frame_1)
        self.tree_scroll.configure(command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.tree_scroll.set)
        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

        tkinter.Label(frame_2, text="{0} Documents".format(self.count)).pack(side=tkinter.LEFT)

        tkinter.Label(frame_3, textvariable=self.compare_a).pack(side=tkinter.LEFT)
        tkinter.Label(frame_3, text='  vs.  ').pack(side=tkinter.LEFT)
        tkinter.Label(frame_3, textvariable=self.compare_b).pack(side=tkinter.LEFT)
        tkinter.Label(frame_3, text='  ').pack(side=tkinter.LEFT)

        tkinter.Button(
            frame_3,
            text='Compare!',
            command=self.compare_tdocs).pack(side=tkinter.LEFT)

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

            self.tree.insert("", "end", tags=(tag,), values=(
                idx,
                row['Revisions'],
                'Click',
                'Click'))

        self.count = count

        self.tree.tag_configure('odd', background='#E8E8E8')
        self.tree.tag_configure('even', background='#DFDFDF')
        treeview_sort_column(self.tree, 'Rev.')

    def on_double_click(self, event):
        item_id = self.tree.identify("item", event.x, event.y)
        column = int(self.tree.identify_column(event.x)[1:]) - 1  # "#1" -> 0 (one-indexed in TCL)
        item_values = self.tree.item(item_id)['values']
        try:
            actual_value = item_values[column]
        except:
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
        if column == 0:
            print('Opening {0}'.format(actual_value))
            self.parent_tdocs_table.download_and_open_tdoc(actual_value)
        if column == 1:
            print('Opening {0}'.format(tdoc_to_search))
            self.parent_tdocs_table.download_and_open_tdoc(tdoc_to_search)
        if column == 2:
            self.compare_a.set(tdoc_to_search)
            try:
                self.parent_gui_tools.tkvar_tdoc_to_compare_1.set(tdoc_to_search)
            except:
                pass
        if column == 3:
            self.compare_b.set(tdoc_to_search)
            try:
                self.parent_gui_tools.tkvar_tdoc_to_compare_2.set(tdoc_to_search)
            except:
                pass

    def compare_tdocs(self):
        compare_a = self.compare_a.get()
        compare_b = self.compare_b.get()
        print('Comparing {0} vs. {1}'.format(compare_a, compare_b))
        self.parent_gui_tools.compare_tdocs(entry_1=compare_a, entry_2=compare_b)
