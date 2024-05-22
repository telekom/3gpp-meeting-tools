import os.path
import tkinter
from tkinter import ttk
from typing import List

import pyperclip

import parsing.word.pywin32 as word_parser
import server
from application.excel import open_excel_document
from application.os import open_url
from gui.common.generic_table import GenericTable, treeview_set_row_formatting, column_separator_str
from server import tdoc_search
from server.common import download_file_to_location
from server.tdoc_search import MeetingEntry, search_meeting_for_tdoc
from tdoc.utils import is_generic_tdoc
from utils.local_cache import file_exists


class MeetingsTable(GenericTable):

    def __init__(
            self,
            root_widget: tkinter.Tk,
            favicon,
            parent_widget: tkinter.Tk):
        super().__init__(
            parent_widget=parent_widget,
            widget_title="Meetings Table. Double-click start date for invitation. End date for report",
            favicon=favicon,
            column_names=['Meeting', 'Location', 'Start', 'End', 'TDoc Start', 'TDoc End', 'Documents', 'TDoc List',
                          'TDoc Excel'],
            row_height=35,
            display_rows=14,
            root_widget=root_widget
        )
        self.loaded_meeting_entries: List[MeetingEntry] | None = None
        self.chosen_meeting: MeetingEntry | None = None
        self.root_widget = root_widget

        self.meeting_count_tk_str = tkinter.StringVar()
        self.compare_text_tk_str = tkinter.StringVar()

        self.set_column('Meeting', width=200, center=True)
        self.set_column('Location', width=200, center=True)
        self.set_column('Start', width=120, center=True)
        self.set_column('End', width=120, center=True)
        self.set_column('TDoc Start', width=100, center=True)
        self.set_column('TDoc End', width=100, center=True)
        self.set_column('Documents', width=100, center=True)
        self.set_column('TDoc List', width=100, center=True)
        self.set_column('TDoc Excel', width=100, center=True)

        self.tree.bind("<Double-Button-1>", self.on_double_click)

        # Filter by group (only filter we have in this view)
        all_groups = ['All']
        all_groups.extend(tdoc_search.get_meeting_groups())
        self.combo_groups = ttk.Combobox(self.top_frame, values=all_groups, state="readonly", width=6)
        self.combo_groups.set('All')
        self.combo_groups.bind("<<ComboboxSelected>>", self.select_groups)

        # Filter by 3GPP Group/WG
        tkinter.Label(self.top_frame, text="Group: ").pack(side=tkinter.LEFT)
        self.combo_groups.pack(side=tkinter.LEFT)

        # Open/search TDoc
        tkinter.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        self.tkvar_tdoc_id = tkinter.StringVar(self.top_frame)
        self.tkvar_tdoc_id.trace_add('write', self.on_tdoc_search_change)
        self.tdoc_entry = tkinter.Entry(self.top_frame, textvariable=self.tkvar_tdoc_id, width=15, font='TkDefaultFont')
        self.button_open_tdoc = tkinter.Button(
            self.top_frame,
            text='Open TDoc',
            command=self.on_open_tdoc
        )
        self.tdoc_entry.pack(side=tkinter.LEFT)
        tkinter.Label(self.top_frame, text="  ").pack(side=tkinter.LEFT)
        self.button_open_tdoc.pack(side=tkinter.LEFT)

        # Re-download TDoc Excel if it already exists
        self.redownload_tdoc_excel_if_exists_var = tkinter.IntVar()
        self.redownload_tdoc_excel_if_exists = ttk.Checkbutton(
            self.top_frame,
            state='enabled',
            variable=self.redownload_tdoc_excel_if_exists_var)
        tkinter.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        tkinter.Label(self.top_frame, text="Re-download TDoc Excel if exists: ").pack(side=tkinter.LEFT)
        self.redownload_tdoc_excel_if_exists.pack(side=tkinter.LEFT)

        # Load meeting data
        tkinter.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        tkinter.Button(
            self.top_frame,
            text='Load meetings',
            command=self.load_meetings).pack(side=tkinter.LEFT)

        # Compare TDoc
        tkinter.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        self.tkvar_tdoc_id_2 = tkinter.StringVar(self.top_frame)
        self.tkvar_tdoc_id_2.trace_add('write', self.on_tdoc_compare_change)
        self.tdoc_entry_2 = tkinter.Entry(
            self.top_frame,
            textvariable=self.tkvar_tdoc_id_2,
            width=15,
            font='TkDefaultFont')
        self.button_compare_tdoc = tkinter.Button(
            self.top_frame,
            text='Compare TDocs',
            command=self.on_compare_tdoc,
            state=tkinter.DISABLED
        )
        self.tdoc_entry_2.pack(side=tkinter.LEFT)
        tkinter.Label(self.top_frame, text="  ").pack(side=tkinter.LEFT)
        self.button_compare_tdoc.pack(side=tkinter.LEFT)

        # Main frame
        self.load_data(initial_load=True)
        self.insert_rows()

        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

        # Bottom frame
        tkinter.Label(self.bottom_frame, textvariable=self.meeting_count_tk_str).pack(side=tkinter.LEFT)
        tkinter.Label(self.bottom_frame, text="  ").pack(side=tkinter.LEFT)
        tkinter.Label(self.bottom_frame, textvariable=self.compare_text_tk_str).pack(side=tkinter.LEFT)

        # Update text in lower frame
        self.on_tdoc_compare_change()

        # Add text wrapping
        # https: // stackoverflow.com / questions / 51131812 / wrap - text - inside - row - in -tkinter - treeview

    def load_data(self, initial_load=False):
        """
        Loads meetings from the 3GPP website

        Args:
            initial_load: Loads everything
        """
        # Load specs data
        print('Loading revision data for LATEST specs per release for table')
        if initial_load:
            tdoc_search.fully_update_cache()
            self.loaded_meeting_entries = tdoc_search.loaded_meeting_entries
        print('Finished loading meetings')

    def insert_rows(self, tdoc_override=False):
        print('Populating meetings table')

        if self.chosen_meeting is None:
            meeting_list_to_consider = self.loaded_meeting_entries
        else:
            meeting_list_to_consider = [self.chosen_meeting]

        # Filter by selected group
        selected_group = self.combo_groups.get()
        if (selected_group != 'All') and (not tdoc_override):
            meeting_list_to_consider = [m for m in meeting_list_to_consider if m.meeting_group == selected_group]

        # Sort list by date
        meeting_list_to_consider.sort(reverse=True, key=lambda m: (m.start_date, m.meeting_group))

        count = 0
        for idx, meeting in enumerate(meeting_list_to_consider):
            count = count + 1
            mod = count % 2
            if mod > 0:
                tag = 'odd'
            else:
                tag = 'even'

            if meeting.meeting_url_docs is None or meeting.meeting_url_docs == '':
                documents_str = '-'
                tdoc_list_str = '-'
                tdoc_excel_str = '-'
            else:
                documents_str = 'Documents'
                tdoc_list_str = 'Tdoc List'
                tdoc_excel_str = 'Tdoc Excel'

            # 'Meeting', 'Location', 'Start', 'End', 'TDoc Start', 'TDoc End', 'Documents'
            values = (
                meeting.meeting_name,
                meeting.meeting_location,
                meeting.start_date.strftime('%Y-%m-%d'),
                meeting.end_date.strftime('%Y-%m-%d'),
                meeting.tdoc_start,
                meeting.tdoc_end,
                documents_str,
                tdoc_list_str,
                tdoc_excel_str
            )
            self.tree.insert("", "end", tags=(tag,), values=values)

        treeview_set_row_formatting(self.tree)
        self.meeting_count_tk_str.set('{0} meetings'.format(count))

    def clear_filters(self, *args):
        self.combo_groups.set('All')

        # Refill list
        self.apply_filters()

    def load_meetings(self, *args):
        tdoc_search.update_local_html_cache(redownload_if_exists=True)
        self.load_data(initial_load=True)
        self.apply_filters()

    def apply_filters(self, tdoc_override=False):
        self.tree.delete(*self.tree.get_children())
        self.insert_rows(tdoc_override=tdoc_override)

    def select_groups(self, *args):
        self.apply_filters()

    def on_double_click(self, event):
        item_id = self.tree.identify("item", event.x, event.y)
        column = int(self.tree.identify_column(event.x)[1:]) - 1  # "#1" -> 0 (one-indexed in TCL)
        item_values = self.tree.item(item_id)['values']
        try:
            actual_value = item_values[column]
        except:
            actual_value = None

        meeting_name = item_values[0]
        meeting = [m for m in self.loaded_meeting_entries if m.meeting_name == meeting_name]
        print("you clicked on {0}/{1}: {2}".format(event.x, event.y, actual_value))

        if actual_value is None or actual_value == '':
            print("Empty value")
            return

        if column == 0:
            print(f'Clicked on meeting {meeting_name}')
            url_to_open = meeting[0].meeting_url_3gu
            open_url(url_to_open)

        if column == 2:
            print(f'Clicked on start date for meeting {meeting_name}')
            url_to_open = meeting[0].meeting_url_invitation
            open_url(url_to_open)

        if column == 3:
            print(f'Clicked on end date for meeting {meeting_name}')
            url_to_open = meeting[0].meeting_url_report
            open_url(url_to_open)

        if column == 6 and actual_value != '-':
            print(f'Clicked Documents link for meeting {meeting_name}')
            url_to_open = meeting[0].meeting_url_docs
            open_url(url_to_open)

        if column == 7 and actual_value != '-':
            print(f'Clicked TDoc List link for meeting {meeting_name}')
            url_to_open = meeting[0].meeting_tdoc_list_url
            open_url(url_to_open)

        if column == 8 and actual_value != '-':
            print(f'Clicked TDoc Excel link for meeting {meeting_name}')
            download_folder = meeting[0].local_agenda_folder_path
            local_path = os.path.join(download_folder, f'{meeting[0].meeting_name}_TDoc_List.xlsx')
            file_already_exists = file_exists(local_path)
            downloaded = False
            if not file_already_exists or self.redownload_tdoc_excel_if_exists_var.get():
                url_to_open = meeting[0].meeting_tdoc_list_excel_url
                download_file_to_location(url_to_open, local_path)
                downloaded = True
            if not downloaded:
                print('TDoc Excel list opened from cache')
            open_excel_document(local_path)

    def on_open_tdoc(self):
        tdoc_to_open = self.tdoc
        print(f'Opening {tdoc_to_open}')
        opened_docs, metadata = server.tdoc_search.search_download_and_open_tdoc(tdoc_to_open)
        if metadata is not None:
            print(f'Opened Tdoc {metadata[0].tdoc_id}, {metadata[0].url}. Copied URL to clipboard')
            pyperclip.copy(metadata[0].url)

    def on_compare_tdoc(self):
        tdoc1_to_open = self.tdoc
        tdoc2_to_open = self.original_tdoc

        print(f'Comparing {tdoc2_to_open}  (original) vs. {tdoc1_to_open}')
        opened_docs1, metadata1 = server.tdoc_search.search_download_and_open_tdoc(tdoc1_to_open, skip_open=True)
        opened_docs2, metadata2 = server.tdoc_search.search_download_and_open_tdoc(tdoc2_to_open, skip_open=True)
        doc_1 = metadata1[0].path
        doc_2 = metadata2[0].path
        print(f'Comparing {doc_2} vs. {doc_1}')
        word_parser.compare_documents(doc_2, doc_1)

    @property
    def tdoc(self) -> str | None:
        current_tdoc = self.tkvar_tdoc_id.get()
        return current_tdoc.strip() if current_tdoc is not None else None

    @property
    def original_tdoc(self) -> str | None:
        current_tdoc = self.tkvar_tdoc_id_2.get()
        return current_tdoc.strip() if current_tdoc is not None else None

    def on_tdoc_search_change(self, *args):
        self.chosen_meeting = None
        self.combo_groups.configure(state="enabled")

        # Update lower footer
        self.on_tdoc_compare_change()

        current_tdoc = self.tdoc
        if is_generic_tdoc(current_tdoc) is None:
            self.apply_filters()
            return

        meeting_for_tdoc = search_meeting_for_tdoc(current_tdoc)
        if meeting_for_tdoc is None:
            self.apply_filters()
            return

        print(f'TDoc search changed to {current_tdoc} of meeting {meeting_for_tdoc.meeting_name}')
        self.chosen_meeting = meeting_for_tdoc
        self.combo_groups.configure(state="disabled")

        self.apply_filters(tdoc_override=True)

    def on_tdoc_compare_change(self, *args):
        revised_tdoc = self.tdoc
        original_tdoc = self.original_tdoc
        revised_tdoc_is_correct = (is_generic_tdoc(revised_tdoc) is not None)
        original_tdoc_is_correct = (is_generic_tdoc(original_tdoc) is not None)

        if revised_tdoc_is_correct and original_tdoc_is_correct:
            self.compare_text_tk_str.set(f"Click to compare changes from {original_tdoc} to {revised_tdoc}")
            self.button_compare_tdoc.configure(state=tkinter.NORMAL)
        else:
            self.compare_text_tk_str.set("To compare TDocs, input two TDocs to compare")
            self.button_compare_tdoc.configure(state=tkinter.DISABLED)


