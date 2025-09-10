import tkinter
from functools import reduce
from tkinter import ttk
from typing import List
from urllib.parse import parse_qs, urlparse

from application.os import open_url
from gui.common.generic_table import GenericTable, treeview_set_row_formatting, column_separator_str
from gui.common.gui_elements import TTKHoverHelpButton
from gui.common.icons import refresh_icon
from server import tdoc_search
from server.common.MeetingEntry import MeetingEntry, WorkItem
from server.meeting import batch_download_meeting_tdocs_excel


class WorkItemsTable(GenericTable):

    def __init__(
            self,
            root_widget: tkinter.Tk | None,
            favicon,
            parent_widget: tkinter.Tk | None):
        super().__init__(
            parent_widget=parent_widget,
            widget_title="3GPP Work Items from meeting TDoc list. Double-click: WI code or name for 3GPP page",
            favicon=favicon,
            column_names=['WI Code', 'Acronym'],
            row_height=35,
            display_rows=14,
            root_widget=root_widget
        )
        self.loaded_meeting_entries: List[MeetingEntry] | None = None
        self.chosen_meeting: MeetingEntry | None = None
        self.root_widget = root_widget
        self.wi_list: list[WorkItem] = []
        self.finished_loading = False
        self.redownload_meeting_list_var = tkinter.IntVar()

        # Start by loading data
        self.load_data(initial_load=True)

        self.meeting_count_tk_str = tkinter.StringVar()
        self.compare_text_tk_str = tkinter.StringVar()

        self.set_column('WI Code', width=200, center=True)
        self.set_column('Acronym', width=200, center=True)

        self.tree.bind("<Double-Button-1>", self.on_double_click)

        # Filter by group (only filter we have in this view)
        all_groups = ['All Groups']
        meeting_groups_from_3gpp_server = tdoc_search.get_meeting_groups()
        meeting_groups_from_3gpp_server.append('S3-LI')
        all_groups.extend(meeting_groups_from_3gpp_server)
        all_groups.sort()

        self.combo_groups = ttk.Combobox(
            self.top_frame,
            values=all_groups,
            state="readonly",
            width=10)
        self.combo_groups.set('All Groups')
        self.combo_groups.bind("<<ComboboxSelected>>", self.select_rows)
        self.combo_groups.pack(side=tkinter.LEFT)
        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)

        # Filter by year (starting date)
        all_years = ['All Years']
        meeting_years = tdoc_search.get_meeting_years()
        all_years.extend(meeting_years)

        self.combo_years = ttk.Combobox(
            self.top_frame,
            values=all_years,
            state="readonly",
            width=10)
        self.combo_years.set('All Years')
        self.combo_years.bind("<<ComboboxSelected>>", self.select_rows)
        self.combo_years.pack(side=tkinter.LEFT)

        # Open/search TDoc
        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)

        # Re-download TDoc Excel if it already exists
        self.redownload_meeting_list = ttk.Checkbutton(
            self.top_frame,
            state='enabled',
            variable=self.redownload_meeting_list_var)

        # Load meeting data
        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        TTKHoverHelpButton(
            self.top_frame,
            help_text='(Re-)load work item list for selected meeting filter',
            command=self.apply_filters,
            image=refresh_icon
        ).pack(side=tkinter.LEFT)

        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        self.redownload_meeting_list.pack(side=tkinter.LEFT)
        ttk.Label(self.top_frame, text="Re-DL meeting list").pack(side=tkinter.LEFT)

        # Main frame
        self.insert_rows()

        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

        # Bottom frame
        ttk.Label(self.bottom_frame, textvariable=self.meeting_count_tk_str).pack(side=tkinter.LEFT)
        ttk.Label(self.bottom_frame, text=" ").pack(side=tkinter.LEFT)
        ttk.Label(self.bottom_frame, textvariable=self.compare_text_tk_str).pack(side=tkinter.LEFT)

        # Add text wrapping
        # https: // stackoverflow.com / questions / 51131812 / wrap - text - inside - row - in -tkinter - treeview

        self.finished_loading = True

    def load_data(self, initial_load=False):
        """
        Loads meetings from the 3GPP website

        Args:
            initial_load: Loads everything
        """
        # Load specs data
        print('Loading revision data for LATEST specs per release for table')
        if initial_load or self.redownload_meeting_list_var.get():
            tdoc_search.fully_update_cache(redownload_if_exists=self.redownload_meeting_list_var.get())
            self.loaded_meeting_entries = tdoc_search.loaded_meeting_entries
        print('Finished loading meetings')

    def insert_rows(self, tdoc_override=False):
        if not self.finished_loading:
            print(f'Initial load: not populating table')
            return

        selected_year = self.combo_years.get()
        selected_group = self.combo_groups.get()
        print(f'Populating WI table: WG {selected_group} for year {selected_year}')

        def meeting_matches_filter(m:MeetingEntry)->bool:
            filter_match = True

            # Filter by selected year
            if (not selected_year.startswith('All')) and (not tdoc_override):
                filter_match = filter_match and m.starts_in_given_year(int(selected_year))

            # Filter by selected group
            if (not selected_group.startswith('All')) and (not tdoc_override):
                if selected_group == 'S3-LI':
                    filter_match = filter_match and (m.meeting_group == 'S3' and m.is_li)
                elif selected_group == 'S3':
                    filter_match = filter_match and (m.meeting_group == 'S3' and not m.is_li)
                else:
                    filter_match = filter_match and (m.meeting_group == selected_group)

            return filter_match

        selected_meetings = [m for m in self.loaded_meeting_entries if
                             meeting_matches_filter(m) ]
        print(f'{len(selected_meetings)} meetings selected')

        # Download meetings if necessary
        batch_download_meeting_tdocs_excel(selected_meetings)

        list_of_lists = [m.tdoc_data_from_excel_with_cache_overwrite.work_items for m in selected_meetings]
        wi_list = [item for sublist in list_of_lists for item in sublist]
        wi_list = list(set(wi_list))
        wi_list = sorted(wi_list, key=lambda x:x.acronym)
        self.wi_list = wi_list

        count = 0
        for wi_data in wi_list:
            count = count + 1
            mod = count % 2
            if mod > 0:
                tag = 'odd'
            else:
                tag = 'even'

            wi_code = wi_data.acronym
            work_item_id = wi_data.work_item_id

            # 'Meeting', 'Location', 'Start', 'End', 'TDoc Start', 'TDoc End', 'Documents'
            values = (
                wi_code,
                work_item_id,
            )
            self.tree.insert(
                "",
                "end",
                tags=(tag,),
                values=values
            )

        treeview_set_row_formatting(self.tree)
        self.meeting_count_tk_str.set('{0} work items'.format(count))

    def clear_filters(self, *args):
        self.combo_groups.set('All Groups')
        self.combo_years.set('All Years')

        # Refill list
        # self.apply_filters()

    def load_meetings(self, *args):
        tdoc_search.fully_update_cache(redownload_if_exists=True)
        self.load_data(initial_load=True)
        self.apply_filters()

    def apply_filters(self, tdoc_override=False):
        self.tree.delete(*self.tree.get_children())
        self.insert_rows(tdoc_override=tdoc_override)

    def select_rows(self, *args):
        # self.apply_filters()
        pass

    def on_double_click(self, event):
        item_id = self.tree.identify("item", event.x, event.y)
        column = int(self.tree.identify_column(event.x)[1:]) - 1  # "#1" -> 0 (one-indexed in TCL)
        item_values = self.tree.item(item_id)['values']
        try:
            actual_value = item_values[column]
        except Exception as e:
            print(f'Could not process TreeView double-click: {e}')
            actual_value = None

        wi_acronym = item_values[0]
        wis = [e for e in self.wi_list if e is not None and e.acronym==wi_acronym]
        if wis is not None and len(wis)>0:
            wi = wis[0]
        else:
            wi = None

        print("you clicked on {0}/{1}: {2}".format(event.x, event.y, actual_value))

        if column == 0 or column == 1: # WI
            print(f'Clicked on WI {wi_acronym}: {wi.url}')
            url_to_open = wi.url
            open_url(url_to_open)


