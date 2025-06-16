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
from server.common.MeetingEntry import MeetingEntry
from server.common.server_enums import WorkingGroup


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
        self.wi_dict = {}

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
        self.redownload_meeting_list_var = tkinter.IntVar()
        self.redownload_meeting_list = ttk.Checkbutton(
            self.top_frame,
            state='enabled',
            variable=self.redownload_meeting_list_var)

        # Load meeting data
        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        TTKHoverHelpButton(
            self.top_frame,
            help_text='(Re-)load work item list from selected meetings',
            command=self.load_meetings,
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
        print('Populating WI table')

        def meeting_matches_filter(m:MeetingEntry)->bool:
            filter_match = True

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

        selected_meetings = [m for m in self.loaded_meeting_entries if
                             m.starts_in_given_year(2025) and
                             m.working_group_enum == WorkingGroup.S2 and
                             meeting_matches_filter(m) ]
        selected_meetings = selected_meetings[0:2]
        list_of_dicts = [m.tdoc_data_from_excel.wi_hyperlinks for m in selected_meetings]
        wi_dict:dict[str,str] = reduce(lambda acc, current_dict: {**acc, **current_dict}, list_of_dicts, {})
        wi_list = wi_dict.items()
        self.wi_dict = wi_dict

        count = 0
        for idx, wi_data in enumerate(wi_list):
            count = count + 1
            mod = count % 2
            if mod > 0:
                tag = 'odd'
            else:
                tag = 'even'

            wi_code = wi_data[0]
            wi_url = wi_data[1]

            # e.g. "https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId=1060084"
            work_item_id = parse_qs(urlparse(wi_url).query).get('workitemId', [None])[0]

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
        self.apply_filters()

    def load_meetings(self, *args):
        tdoc_search.fully_update_cache(redownload_if_exists=True)
        self.load_data(initial_load=True)
        self.apply_filters()

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

        wi_acronym = item_values[0]
        wi_url = self.wi_dict[wi_acronym]

        print("you clicked on {0}/{1}: {2}".format(event.x, event.y, actual_value))

        if column == 0 or column == 1: # WI
            print(f'Clicked on WI {wi_acronym}')
            url_to_open = wi_url
            open_url(url_to_open)


