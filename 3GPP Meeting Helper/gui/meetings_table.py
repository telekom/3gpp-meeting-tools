import datetime
import os.path
import platform
import tkinter
from tkinter import ttk
from typing import List, Final

import pandas
import pandas as pd
import pyperclip
from pandas.core.interchange.dataframe_protocol import DataFrame

import application.excel
import gui.common.utils
import server
import utils.local_cache
from application.excel import open_excel_document
from application.os import open_url, startfile
from config.meetings import MeetingConfig
from gui.common.common_elements import tkvar_3gpp_wifi_available
from gui.common.generic_table import GenericTable, treeview_set_row_formatting, column_separator_str
from gui.common.gui_elements import TTKHoverHelpButton
from gui.common.icons import refresh_icon, search_icon, compare_icon, table_icon
from gui.tdocs_table_from_excel import TdocsTableFromExcel
from server import tdoc_search
from server.common.MeetingEntry import MeetingEntry
from server.common.server_utils import download_file_to_location
from server.meeting import batch_download_meeting_tdocs_excel
from server.tdoc_search import search_meeting_for_tdoc, compare_two_tdocs
from tdoc.utils import is_generic_tdoc


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
            column_names=['Meeting', 'Location', 'Start', 'End', 'TDoc Start', 'TDoc End', 'TDocs Excel',
                          'TDocs Table'],
            row_height=35,
            display_rows=14,
            root_widget=root_widget
        )
        self.loaded_meeting_entries: List[MeetingEntry] | None = None
        self.chosen_meeting: MeetingEntry | None = None
        self.current_meeting_list: List[MeetingEntry] = []
        self.root_widget = root_widget

        # Start by loading data
        self.load_data(initial_load=True)

        self.meeting_count_tk_str = tkinter.StringVar()
        self.compare_text_tk_str = tkinter.StringVar()

        self.set_column('Meeting', width=200, center=True)
        self.set_column('Location', width=200, center=True)
        self.set_column('Start', width=120, center=True)
        self.set_column('End', width=120, center=True)
        self.set_column('TDoc Start', width=100, center=True)
        self.set_column('TDoc End', width=100, center=True)
        self.set_column('TDocs Excel', width=100, center=True)
        self.set_column('TDocs Table', width=100, center=True)

        self.tree.bind("<Double-Button-1>", self.on_double_click)

        # Filter by group (only filter we have in this view)
        all_groups_str: Final[str] = 'All Groups'
        all_groups = [all_groups_str]
        meeting_groups_from_3gpp_server = tdoc_search.get_meeting_groups()
        meeting_groups_from_3gpp_server.append('S3-LI')
        all_groups.extend(meeting_groups_from_3gpp_server)
        all_groups.sort()

        self.combo_groups = ttk.Combobox(
            self.top_frame,
            values=all_groups,
            state="readonly",
            width=10)
        if MeetingConfig.meeting_list_group_filter is not None and MeetingConfig.meeting_list_group_filter in all_groups:
            print(f'Setting group filter to "{MeetingConfig.meeting_list_group_filter}"')
            self.combo_groups.set(MeetingConfig.meeting_list_group_filter)
        else:
            print(f'Setting group filter to "{all_groups_str}"')
            self.combo_groups.set(all_groups_str)
        self.combo_groups.bind("<<ComboboxSelected>>", self.select_rows)
        self.combo_groups.pack(side=tkinter.LEFT)
        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)

        # Filter by year (starting date)
        all_years_str: Final[str] = 'All Years'
        all_years = [all_years_str]
        meeting_years = tdoc_search.get_meeting_years()
        all_years.extend(meeting_years)

        self.combo_years = ttk.Combobox(
            self.top_frame,
            values=all_years,
            state="readonly",
            width=10)
        current_year = datetime.datetime.now().year
        if current_year in all_years:
            print(f'Set year filter to current year "{current_year}"')
            self.combo_years.set(current_year)
        else:
            print(f'Year "{current_year}" not in {all_years}. Using "{all_years_str}"')
            self.combo_years.set(all_years_str)
        self.combo_years.bind("<<ComboboxSelected>>", self.select_rows)
        self.combo_years.pack(side=tkinter.LEFT)

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
        self.button_open_tdoc = TTKHoverHelpButton(
            self.top_frame,
            help_text='Search for TDoc',
            command=self.on_open_tdoc,
            state=tkinter.DISABLED,
            image=search_icon
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

        # Compare TDoc
        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        self.tkvar_tdoc_id_2 = tkinter.StringVar(self.top_frame)
        self.tkvar_tdoc_id_2.trace_add('write', self.on_tdoc_compare_change)
        self.tdoc_entry_2 = tkinter.Entry(
            self.top_frame,
            textvariable=self.tkvar_tdoc_id_2,
            width=15,
            font='TkDefaultFont')
        self.button_compare_tdoc = TTKHoverHelpButton(
            self.top_frame,
            help_text='Show changes from right TDoc to left TDoc',
            image=compare_icon,
            command=self.on_compare_tdoc,
            state=tkinter.DISABLED
        )

        if platform.system() == 'Windows':
            # Only works in Windows
            self.button_compare_tdoc.pack(side=tkinter.LEFT)
            ttk.Label(self.top_frame, text=" ").pack(side=tkinter.LEFT)
            self.tdoc_entry_2.pack(side=tkinter.LEFT)

        # Load meeting data
        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        TTKHoverHelpButton(
            self.top_frame,
            help_text='(Re-)load meeting list from 3GPP server',
            command=self.load_meetings,
            image=refresh_icon
        ).pack(side=tkinter.LEFT)

        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        self.redownload_tdoc_excel_if_exists.pack(side=tkinter.LEFT)
        ttk.Label(self.top_frame, text="Re-DL TDoc list").pack(side=tkinter.LEFT)

        # Merge TDoc lists
        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        TTKHoverHelpButton(
            self.top_frame,
            help_text='Show list of all TDocs in shown meetings',
            command=self.load_data_for_several_meetings,
            image=table_icon
        ).pack(side=tkinter.LEFT)

        # Main frame
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

    def meeting_list_to_consider(self, tdoc_override=False) -> List[MeetingEntry]:
        try:
            if self.chosen_meeting is None:
                if self.loaded_meeting_entries is None:
                    meeting_list_to_consider = []
                else:
                    meeting_list_to_consider = self.loaded_meeting_entries
            else:
                meeting_list_to_consider = [self.chosen_meeting]
        except Exception as e:
            print(f'Could not retrieve current meeting list: {e}')
            meeting_list_to_consider = []

        def meeting_matches_filter(m: MeetingEntry) -> bool:
            filter_match = True

            # If filtering by "now"
            if self.filter_by_now:
                filter_match = filter_match and m.meeting_is_now

            # Filter by selected year
            selected_year = self.combo_years.get()
            if (not selected_year.startswith('All')) and (not tdoc_override):
                filter_match = filter_match and m.starts_in_given_year(int(selected_year))

            # Filter by selected group
            selected_group = self.combo_groups.get()
            if (not selected_group.startswith('All')) and (not tdoc_override):
                if selected_group == 'S3-LI':
                    filter_match = filter_match and (m.meeting_group == 'S3' and m.is_li)
                elif selected_group == 'S3':
                    filter_match = filter_match and (m.meeting_group == 'S3' and not m.is_li)
                else:
                    filter_match = filter_match and (m.meeting_group == selected_group)

            return filter_match

        meeting_list_to_consider = [m for m in meeting_list_to_consider if meeting_matches_filter(m)]

        # Sort list by date
        meeting_list_to_consider.sort(reverse=True, key=lambda m: (m.start_date, m.meeting_group))
        return meeting_list_to_consider

    def insert_rows(self, tdoc_override=False):
        print('Populating meetings table')

        count = 0
        previous_row: None | MeetingEntry = None
        meetings_to_list = self.meeting_list_to_consider(tdoc_override)
        for idx, meeting in enumerate(meetings_to_list):
            count = count + 1
            mod = count % 2
            if mod > 0:
                tag = 'odd'
            else:
                tag = 'even'

            if meeting.meeting_url_docs is None or meeting.meeting_url_docs == '':
                tdoc_excel_str = '-'
                tdoc_table_str = '-'
            else:
                tdoc_excel_str = 'Open'
                tdoc_table_str = 'Open'

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
                tdoc_excel_str,
                tdoc_table_str
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
        self.current_meeting_list = meetings_to_list

    def clear_filters(self, *args):
        self.combo_groups.set('All Groups')
        self.combo_years.set('All Years')

        # Refill list
        self.apply_filters()

    def load_meetings(self, *args):
        tdoc_search.fully_update_cache(redownload_if_exists=True)
        self.load_data(initial_load=True)
        self.apply_filters()

    def load_data_for_several_meetings(self, *args):
        groups_set = list(set([e.meeting_group for e in self.current_meeting_list]))
        groups_set.sort()
        groups_set_str = ', '.join(groups_set)
        print(f'Merging TDoc list from {len(self.current_meeting_list)} meetings from groups: {groups_set_str}')

        # Download TDoc Excel files if necessary
        batch_download_meeting_tdocs_excel(self.current_meeting_list)

        def apply_meeting_data_to_df(
                group_name: str,
                meeting_name: str,
                start_date:datetime,
                docs_folder:str,
                df_in: DataFrame)->DataFrame:
            df_out = df_in.copy()
            df_out['WG'] = group_name
            df_out['Meeting'] = meeting_name
            df_out['Start date'] = start_date
            df_out['Start date'] = pd.to_datetime(df_out['Start date']).dt.date

            def generate_cell_hyperlink(cell_value, cell_tdoc_idx):
                return f'=HYPERLINK("{cell_value}{cell_tdoc_idx}.zip","{cell_tdoc_idx}")'

            df_out['Hyperlink'] = docs_folder
            df_out['Hyperlink'] = df_out.apply(lambda row: generate_cell_hyperlink(row['Hyperlink'], row.name), axis=1)
            df_out = df_out.set_index(['Hyperlink'])
            df_out.index.names = ['TDoc']
            return df_out

        merged_df = pandas.concat(
            [apply_meeting_data_to_df(
                e.meeting_group,
                e.meeting_name,
                e.start_date,
                e.meeting_url_docs,
                e.tdoc_data_from_excel.tdocs_df) for e in
                       self.current_meeting_list if e.tdoc_data_from_excel is not None],
            axis=0, join='outer',
            ignore_index=False, keys=None, levels=None, names=None, verify_integrity=False, copy=True)
        print(f'Total of {len(merged_df)} TDocs')
        export_path = os.path.join(utils.local_cache.get_cache_folder(), 'Export')
        utils.local_cache.create_folder_if_needed(folder_name=export_path, create_dir=True)
        now = datetime.datetime.now()
        file_name = f'{now.year}.{now.month}.{now.day} {now.hour}{now.minute}{now.second} TDoc export.xlsx'
        excel_export = os.path.join(export_path, file_name)
        merged_df.to_excel(
            excel_export,
            freeze_panes=(1, 1))
        wb = application.excel.open_excel_document(excel_export)
        application.excel.set_first_row_as_filter(wb)
        application.excel.set_column_width('A', wb, 10) # TDoc
        application.excel.set_column_width('B', wb, 36)  # Title
        application.excel.hide_column('D', wb)  # Contact
        application.excel.hide_column('E', wb) # Contact ID
        application.excel.hide_column('G', wb) # For
        application.excel.hide_column('J', wb)  # Agenda item sort order
        application.excel.set_column_width('H', wb, 30)  # Abstract
        application.excel.set_column_width('I', wb, 17)  # Secretary Remarks
        application.excel.set_column_width('L', wb, 24)  # Agenda item description
        application.excel.hide_column('M', wb) # TDoc sort order within agenda item
        application.excel.set_column_width('N', wb, 20)  # TDoc Status
        application.excel.hide_column('O', wb)  # Reservation date
        application.excel.hide_column('P', wb)  # Uploaded
        application.excel.set_column_width('Q', wb, 13.5)  # Is revision of
        application.excel.set_column_width('V', wb, 16)  # Related WIs
        application.excel.set_column_width('R', wb, 13.5)  # Revised to
        application.excel.set_column_width('AL', wb, 12.5)  # Meeting
        application.excel.set_column_width('AM', wb, 10)  # Meeting
        application.excel.set_wrap_text(wb)
        application.excel.vertically_center_all_text(wb)
        application.excel.apply_tdoc_status_conditional_formatting_formula('N', wb)
        print(f'Exported TDocs to {excel_export}')
        # os.startfile(excel_export)

    def apply_filters(self, tdoc_override=False):
        self.tree.delete(*self.tree.get_children())
        self.insert_rows(tdoc_override=tdoc_override)

    def select_rows(self, *args):
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
        meeting_list = [m for m in self.loaded_meeting_entries if m.meeting_name == meeting_name]
        print("you clicked on {0}/{1}: {2}".format(event.x, event.y, actual_value))
        try:
            meeting: MeetingEntry = meeting_list[0]
            print(
                f"Selected meeting: {meeting.meeting_number} ({meeting.meeting_name}), URL: {meeting.meeting_folder_url}")
        except Exception as e:
            print(f"Could not retrieve meeting for {meeting_name}, {e}")
            return

        if actual_value is None or actual_value == '':
            print("Empty value")
            return

        local_path = meeting.tdoc_excel_local_path

        if column == 0:  # Meeting
            print(f'Clicked on meeting {meeting_name}')
            url_to_open = meeting.meeting_url_3gu
            open_url(url_to_open)

        if column == 1:  # Location
            print(f'Clicked on meeting {meeting_name} location')
            url_to_open = meeting.meeting_calendar_ics_url
            # Using generic folder because meeting folder may not yet exist
            download_folder = utils.local_cache.get_meeting_list_folder()
            local_path = os.path.join(download_folder, f'{meeting.meeting_name}.ics')
            download_file_to_location(url_to_open, local_path, force_download=True)
            if utils.local_cache.file_exists(local_path):
                startfile(local_path)
                print(f'Opened ICS file {local_path}')
            else:
                print(f'Could not open ICS file {local_path}')

        if column == 2:  # Start
            print(f'Clicked on start date for meeting {meeting_name}')
            url_to_open = meeting.meeting_url_invitation
            open_url(url_to_open)

        if column == 3:  # End
            print(f'Clicked on end date for meeting {meeting_name}')
            url_to_open = meeting.meeting_url_report
            open_url(url_to_open)

        if (column == 6 or column == 7) and actual_value != '-':  # TDocs Excel
            print(f'Clicked TDoc Excel link for meeting {meeting_name}')
            file_already_exists = meeting.tdoc_excel_exists_in_local_folder
            if file_already_exists is None:
                print(f'Meeting folder name not yet known. Cannot save local file')
                return

            downloaded = False
            if not file_already_exists or self.redownload_tdoc_excel_if_exists_var.get():
                url_to_open = meeting.meeting_tdoc_list_excel_url
                download_file_to_location(url_to_open, local_path, force_download=True)
                downloaded = True
            if not downloaded:
                print('TDoc Excel list from cache')

            if column == 6:
                # Open TDoc Excel
                print(f'Opening Excel {local_path}')
                open_excel_document(local_path)
            else:
                # Open TDoc table from Excel
                print(f'Opening TDoc table based on {local_path}')
                TdocsTableFromExcel(
                    favicon=self.favicon,
                    parent_widget=self.tk_top,
                    meeting=meeting,
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

        compare_two_tdocs(
            tdoc1_to_open,
            tdoc2_to_open,
            tkvar_3gpp_wifi_available=tkvar_3gpp_wifi_available
        )

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
