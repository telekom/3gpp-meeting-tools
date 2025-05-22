import os.path
import platform
import tkinter
from tkinter import ttk
from typing import List

import pyperclip

import server
import utils.local_cache
from application.excel import open_excel_document
from application.os import open_url, startfile
from gui.common.common_elements import tkvar_3gpp_wifi_available
from gui.common.generic_table import GenericTable, treeview_set_row_formatting, column_separator_str
from gui.common.gui_elements import TTKHoverHelpButton
from gui.common.icons import excel_icon, table_icon
from gui.tdocs_table_from_excel import TdocsTableFromExcel
from server import tdoc_search
from server.common import download_file_to_location, MeetingEntry
from server.tdoc_search import search_meeting_for_tdoc, compare_two_tdocs
from tdoc.utils import is_generic_tdoc
from utils.local_cache import file_exists
import gui.common.utils


class MeetingsTable(GenericTable):

    def __init__(
            self,
            root_widget: tkinter.Tk | None,
            favicon,
            parent_widget: tkinter.Tk | None):
        super().__init__(
            parent_widget=parent_widget,
            widget_title="Meetings Table. Double-click: location for ICS, start date for invitation, end date for report",
            favicon=favicon,
            column_names=['Meeting', 'Location', 'Start', 'End', 'TDoc Start', 'TDoc End', 'Documents',
                          'Cache', 'TDocs (3GPP)', 'TDocs (Local)'],
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
        self.set_column('Cache', width=50, center=True)
        self.set_column('TDocs (3GPP)', width=100, center=True)
        self.set_column('TDocs (Local)', width=100, center=True)

        self.tree.bind("<Double-Button-1>", self.on_double_click)

        # Filter by group (only filter we have in this view)
        all_groups = ['All']
        meeting_groups_from_3gpp_server = tdoc_search.get_meeting_groups()
        meeting_groups_from_3gpp_server.append('S3-LI')
        all_groups.extend(meeting_groups_from_3gpp_server)

        all_groups.sort()

        self.combo_groups = ttk.Combobox(
            self.top_frame,
            values=all_groups,
            state="readonly",
            width=6)
        self.combo_groups.set('All')
        self.combo_groups.bind("<<ComboboxSelected>>", self.select_groups)

        # Filter by 3GPP Group/WG
        ttk.Label(self.top_frame, text="Group: ").pack(side=tkinter.LEFT)
        self.combo_groups.pack(side=tkinter.LEFT)

        self.filter_by_now = False

        def toggle_now():
            self.filter_by_now = not self.filter_by_now
            print(f'Filtering by "now": {self.filter_by_now}')
            self.apply_filters()

        self.now_button = ttk.Button(
            self.top_frame,
            text='Now',
            command=toggle_now,
            width=4
        )
        ttk.Label(self.top_frame, text=" ").pack(side=tkinter.LEFT)
        self.now_button.pack(side=tkinter.LEFT)

        # Open/search TDoc
        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        self.tkvar_tdoc_id = tkinter.StringVar(self.top_frame)
        self.tkvar_tdoc_id.trace_add('write', self.on_tdoc_search_change)
        self.tdoc_entry = tkinter.Entry(self.top_frame, textvariable=self.tkvar_tdoc_id, width=15, font='TkDefaultFont')
        self.button_open_tdoc = ttk.Button(
            self.top_frame,
            text='Open TDoc',
            command=self.on_open_tdoc,
            state=tkinter.DISABLED
        )
        self.tdoc_entry.pack(side=tkinter.LEFT)
        ttk.Label(self.top_frame, text=" ").pack(side=tkinter.LEFT)
        self.button_open_tdoc.pack(side=tkinter.LEFT)

        gui.common.utils.bind_key_to_button(
            frame=self.tk_top,
            key_press='<Return>',
            tk_button=self.button_open_tdoc)

        # Re-download TDoc Excel if it already exists
        self.redownload_tdoc_excel_if_exists_var = tkinter.IntVar()
        self.redownload_tdoc_excel_if_exists = ttk.Checkbutton(
            self.top_frame,
            state='enabled',
            variable=self.redownload_tdoc_excel_if_exists_var)
        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        ttk.Label(self.top_frame, text="Re-DL TDoc list: ").pack(side=tkinter.LEFT)
        self.redownload_tdoc_excel_if_exists.pack(side=tkinter.LEFT)

        # Open TDoc Excel as table
        self.open_tdoc_excel_as_table_var = tkinter.IntVar()

        def toggle_excel_choice_btn(*args):
            match self.open_tdoc_excel_choice_var.get():
                case "Excel":
                    self.open_tdoc_excel_choice_var.set("Table")
                    try:
                        self.open_tdoc_excel_choice_btn.configure({'image': table_icon})
                    except AttributeError:
                        # The first time it will always fail
                        pass
                    self.open_tdoc_excel_as_table_var.set(True)
                case _:
                    self.open_tdoc_excel_choice_var.set("Excel")
                    try:
                        self.open_tdoc_excel_choice_btn.configure({'image': excel_icon})
                    except AttributeError:
                        # The first time it will always fail
                        pass
                    self.open_tdoc_excel_as_table_var.set(False)

        self.open_tdoc_excel_choice_var = tkinter.StringVar()
        toggle_excel_choice_btn()  # Init to "Excel"
        self.open_tdoc_excel_choice_btn = TTKHoverHelpButton(
            self.top_frame,
            textvariable=self.open_tdoc_excel_choice_var,
            image=excel_icon,
            command=toggle_excel_choice_btn,
            width=5,
            help_text='Click to change format in which to open TDoc list'
        )
        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        ttk.Label(self.top_frame, text="Open TDocs as ").pack(side=tkinter.LEFT)
        self.open_tdoc_excel_choice_btn.pack(side=tkinter.LEFT)

        # Load meeting data
        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        ttk.Button(
            self.top_frame,
            text='Load meetings',
            command=self.load_meetings).pack(side=tkinter.LEFT)

        # Compare TDoc
        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        self.tkvar_tdoc_id_2 = tkinter.StringVar(self.top_frame)
        self.tkvar_tdoc_id_2.trace_add('write', self.on_tdoc_compare_change)
        self.tdoc_entry_2 = tkinter.Entry(
            self.top_frame,
            textvariable=self.tkvar_tdoc_id_2,
            width=15,
            font='TkDefaultFont')
        self.button_compare_tdoc = ttk.Button(
            self.top_frame,
            text='Compare TDocs',
            command=self.on_compare_tdoc,
            state=tkinter.DISABLED
        )

        if platform.system() == 'Windows':
            # Only works in Windows
            self.tdoc_entry_2.pack(side=tkinter.LEFT)
            ttk.Label(self.top_frame, text=" ").pack(side=tkinter.LEFT)
            self.button_compare_tdoc.pack(side=tkinter.LEFT)

        # Main frame
        self.load_data(initial_load=True)
        self.insert_rows()

        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

        # Bottom frame
        ttk.Label(self.bottom_frame, textvariable=self.meeting_count_tk_str).pack(side=tkinter.LEFT)
        ttk.Label(self.bottom_frame, text=" ").pack(side=tkinter.LEFT)
        ttk.Label(self.bottom_frame, textvariable=self.compare_text_tk_str).pack(side=tkinter.LEFT)

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
            tdoc_search.fully_update_cache(redownload_if_exists=False)
            self.loaded_meeting_entries = tdoc_search.loaded_meeting_entries
        print('Finished loading meetings')

    def insert_rows(self, tdoc_override=False):
        print('Populating meetings table')

        if self.chosen_meeting is None:
            meeting_list_to_consider = self.loaded_meeting_entries
        else:
            meeting_list_to_consider = [self.chosen_meeting]

        # If filtering by "now"
        if self.filter_by_now:
            meeting_list_to_consider = [m for m in meeting_list_to_consider if m.meeting_is_now]

        # Filter by selected group
        selected_group = self.combo_groups.get()
        if (selected_group != 'All') and (not tdoc_override):
            if selected_group == 'S3-LI':
                meeting_list_to_consider = [m for m in meeting_list_to_consider if
                                            m.meeting_group == 'S3' and m.is_li]
            elif selected_group == 'S3':
                meeting_list_to_consider = [m for m in meeting_list_to_consider if
                                            m.meeting_group == 'S3' and not m.is_li]
            else:
                meeting_list_to_consider = [m for m in meeting_list_to_consider if
                                            m.meeting_group == selected_group]

        # Sort list by date
        meeting_list_to_consider.sort(reverse=True, key=lambda m: (m.start_date, m.meeting_group))

        count = 0
        previous_row: None | MeetingEntry = None
        for idx, meeting in enumerate(meeting_list_to_consider):
            count = count + 1
            mod = count % 2
            if mod > 0:
                tag = 'odd'
            else:
                tag = 'even'

            if meeting.meeting_url_docs is None or meeting.meeting_url_docs == '':
                documents_str = '-'
                cache_str = '-'
                tdoc_list_str = 'Link'
                tdoc_excel_str = '-'
            else:
                documents_str = 'Link'
                cache_str = 'Open'
                tdoc_list_str = 'Link'
                tdoc_excel_str = 'Open'

            # Overwrite for case of co-located meetings
            if ((previous_row is not None) and
                    (previous_row.meeting_location == meeting.meeting_location) and
                    (previous_row.start_date == meeting.start_date) and
                    (previous_row.end_date == meeting.end_date)):
                location_str = '"'
                start_date_str = '"'
                end_date_str = '"'
            else:
                location_str = meeting.meeting_location
                start_date_str = meeting.start_date.strftime('%Y-%m-%d')
                end_date_str = meeting.end_date.strftime('%Y-%m-%d')

            # 'Meeting', 'Location', 'Start', 'End', 'TDoc Start', 'TDoc End', 'Documents'
            values = (
                meeting.meeting_name,
                location_str,
                start_date_str,
                end_date_str,
                meeting.tdoc_start,
                meeting.tdoc_end,
                documents_str,
                cache_str,
                tdoc_list_str,
                tdoc_excel_str
            )
            self.tree.insert(
                "",
                "end",
                tags=(tag,),
                values=values
            )
            previous_row = meeting

        treeview_set_row_formatting(self.tree)
        self.meeting_count_tk_str.set('{0} meetings'.format(count))

    def clear_filters(self, *args):
        self.combo_groups.set('All')

        # Refill list
        self.apply_filters()

    def load_meetings(self, *args):
        tdoc_search.fully_update_cache(redownload_if_exists=True)
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
        except Exception as e:
            print(f'Could not process TreeView double-click: {e}')
            actual_value = None

        meeting_name = item_values[0]
        meeting = [m for m in self.loaded_meeting_entries if m.meeting_name == meeting_name]
        print("you clicked on {0}/{1}: {2}".format(event.x, event.y, actual_value))
        try:
            print(f"Selected meeting: {meeting[0].meeting_number} ({meeting[0].meeting_name}), URL: {meeting[0].meeting_folder_url}")
        except Exception as e:
            print(f"Could not retrieve meeting for {meeting_name}, {e}")

        if actual_value is None or actual_value == '':
            print("Empty value")
            return

        if column == 0:
            print(f'Clicked on meeting {meeting_name}')
            url_to_open = meeting[0].meeting_url_3gu
            open_url(url_to_open)

        if column == 1:
            print(f'Clicked on meeting {meeting_name} location')
            url_to_open = meeting[0].meeting_calendar_ics_url
            # Using generic folder because meeting folder may not yet exist
            download_folder = utils.local_cache.get_meeting_list_folder()
            local_path = os.path.join(download_folder, f'{meeting[0].meeting_name}.ics')
            download_file_to_location(url_to_open, local_path, force_download=True)
            if utils.local_cache.file_exists(local_path):
                startfile(local_path)
                print(f'Opened ICS file {local_path}')
            else:
                print(f'Could not open ICS file {local_path}')

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
            print(f'Clicked Cache link for meeting {meeting_name}')
            path_to_open = meeting[0].local_folder_path
            utils.local_cache.create_folder_if_needed(path_to_open, create_dir=True)
            open_url(path_to_open)

        if column == 8 and actual_value != '-':
            print(f'Clicked TDoc List link for meeting {meeting_name}')
            url_to_open = meeting[0].meeting_tdoc_list_url
            open_url(url_to_open)

        if column == 9 and actual_value != '-':
            print(f'Clicked TDoc Excel link for meeting {meeting_name}')
            download_folder = meeting[0].local_agenda_folder_path
            if download_folder is None:
                print(f'Meeting folder name not yet known. Cannot save local file')
                return
            local_path = os.path.join(download_folder, f'{meeting[0].meeting_name}_TDoc_List.xlsx')
            file_already_exists = file_exists(local_path)
            downloaded = False
            if not file_already_exists or self.redownload_tdoc_excel_if_exists_var.get():
                url_to_open = meeting[0].meeting_tdoc_list_excel_url
                download_file_to_location(url_to_open, local_path, force_download=True)
                downloaded = True
            if not downloaded:
                print('TDoc Excel list from cache')
            if not self.open_tdoc_excel_as_table_var.get():
                print(f'Opening Excel {local_path}')
                open_excel_document(local_path)
            else:
                # Open TDoc table from Excel
                print(f'Opening TDoc table based on {local_path}')
                TdocsTableFromExcel(
                    favicon=self.favicon,
                    parent_widget=self.tk_top,
                    meeting=meeting[0],
                    root_widget=self.root_widget,
                    tdoc_excel_path=local_path)

    def on_open_tdoc(self):
        tdoc_to_open = self.tdoc
        print(f'Opening {tdoc_to_open}')
        opened_docs_folder, metadata = server.tdoc_search.search_download_and_open_tdoc(
            tdoc_to_open,
            tkvar_3gpp_wifi_available=tkvar_3gpp_wifi_available)
        if metadata is not None:
            print(f'Opened Tdoc {metadata[0].tdoc_id}, {metadata[0].url}. Copied URL to clipboard')
            pyperclip.copy(metadata[0].url)

    def on_compare_tdoc(self):
        tdoc1_to_open = self.tdoc
        tdoc2_to_open = self.original_tdoc

        compare_two_tdocs(tdoc1_to_open, tdoc2_to_open)

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
            # Disable button to search if TDoc is not valid
            self.button_open_tdoc.configure(state=tkinter.DISABLED)
            self.apply_filters()
            return

        # Enable button to search if TDoc is valid
        self.button_open_tdoc.configure(state=tkinter.NORMAL)
        meeting_for_tdoc = search_meeting_for_tdoc(current_tdoc, return_last_meeting_if_tdoc_is_new=True)
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
