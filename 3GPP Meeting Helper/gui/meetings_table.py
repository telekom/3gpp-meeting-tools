import tkinter
from tkinter import ttk
from typing import List

import pandas as pd

from application.os import open_url_and_copy_to_clipboard
from gui.generic_table import GenericTable, set_column, treeview_set_row_formatting
from parsing.html.specs import cleanup_spec_name
from server import tdoc_search
from server.specs import version_to_file_version, get_url_for_spec_page, get_specs_folder, get_url_for_crs_page
from server.tdoc_search import MeetingEntry


class MeetingsTable(GenericTable):

    def __init__(self, parent, favicon, parent_gui_tools):
        super().__init__(
            parent,
            "Meetings Table",
            favicon,
            ['Meeting', 'Location', 'Start', 'End', 'TDoc Start', 'TDoc End', 'Documents']
        )
        self.loaded_meeting_entries = None
        self.parent_gui_tools = parent_gui_tools

        self.meeting_count = tkinter.StringVar()

        set_column(self.tree, 'Meeting', width=200, center=False)
        set_column(self.tree, 'Location', width=200, center=False)
        set_column(self.tree, 'Start', width=120, center=True)
        set_column(self.tree, 'End', width=120, center=True)
        set_column(self.tree, 'TDoc Start', width=100, center=True)
        set_column(self.tree, 'TDoc End', width=100, center=True)
        set_column(self.tree, 'Documents', width=100, center=True)

        self.tree.bind("<Double-Button-1>", self.on_double_click)

        # Filter by group (only filter we have in this view)
        all_groups = ['All']
        all_groups.extend(tdoc_search.get_meeting_groups())
        self.combo_groups = ttk.Combobox(self.top_frame, values=all_groups, state="readonly", width=6)
        self.combo_groups.set('All')
        self.combo_groups.bind("<<ComboboxSelected>>", self.select_groups)

        tkinter.Label(self.top_frame, text="Group: ").pack(side=tkinter.LEFT)
        self.combo_groups.pack(side=tkinter.LEFT)

        tkinter.Label(self.top_frame, text="     ").pack(side=tkinter.LEFT)
        tkinter.Button(
            self.top_frame,
            text='Load meetings',
            command=self.load_meetings).pack(side=tkinter.LEFT)

        # Main frame
        self.load_data(initial_load=True)
        self.insert_current_meetings()

        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

        # Bottom frame
        tkinter.Label(self.bottom_frame, textvariable=self.meeting_count).pack(side=tkinter.LEFT)

        # Add text wrapping
        # https: // stackoverflow.com / questions / 51131812 / wrap - text - inside - row - in -tkinter - treeview

    def load_data(self, initial_load=False):
        """
        Loads specifications frm the 3GPP website

        Args:
            initial_load: Loads everything
        """
        # Load specs data
        print('Loading revision data for LATEST specs per release for table')
        if initial_load:
            tdoc_search.fully_update_cache()
            self.loaded_meeting_entries = tdoc_search.loaded_meeting_entries
        print('Finished loading meetings')

    def insert_current_meetings(self):
        self.insert_rows(self.loaded_meeting_entries)

    def insert_rows(self, meeting_list: List[MeetingEntry]):
        print('Populating meetings table')

        meeting_list_to_consider = meeting_list

        # Filter by selected group
        selected_group = self.combo_groups.get()
        if selected_group != 'All':
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

            # 'Meeting', 'Location', 'Start', 'End', 'TDoc Start', 'TDoc End', 'Documents'
            self.tree.insert("", "end", tags=(tag,), values=(
                meeting.meeting_name,
                meeting.meeting_location,
                meeting.start_date.strftime('%Y-%m-%d'),
                meeting.end_date.strftime('%Y-%m-%d'),
                meeting.tdoc_start,
                meeting.tdoc_end,
                'Documents'
            ))

        treeview_set_row_formatting(self.tree)
        self.meeting_count.set('{0} meetings'.format(count))

    def clear_filters(self, *args):
        self.combo_groups.set('All')

        # Refill list
        self.apply_filters()

    def load_meetings(self, *args):
        tdoc_search.update_local_html_cache(redownload_if_exists=True)
        self.load_data(initial_load=True)
        self.apply_filters()

    def apply_filters(self):
        self.tree.delete(*self.tree.get_children())
        self.insert_current_meetings()

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

        spec_id = cleanup_spec_name(item_values[0], clean_type=True, clean_dots=False)
        print("you clicked on {0}/{1}: {2}".format(event.x, event.y, actual_value))

        if actual_value is None or actual_value == '':
            print("Empty value")
            return

        if column == 0:
            print('Clicked spec ID {0}. Opening 3GPP spec page'.format(actual_value))
            url_to_open = get_url_for_spec_page(spec_id)
            open_url_and_copy_to_clipboard(url_to_open)
        if column == 1:
            print('Clicked title for spec ID {0}: {1}. Opening 3GPP spec page'.format(spec_id, actual_value))
            url_to_open = get_url_for_spec_page(spec_id)
            open_url_and_copy_to_clipboard(url_to_open)
        if column == 2:
            print('Clicked versions for spec ID {0}: {1}'.format(spec_id, actual_value))
            print('No specs table here!')
        if column == 3:
            print('Clicked local folder for spec ID {0}'.format(spec_id))
            url_to_open = get_specs_folder(spec_id=spec_id)
            open_url_and_copy_to_clipboard(url_to_open)
        if column == 5:
            print('Clicked CRs link for spec ID {0}'.format(spec_id))
            url_to_open = get_url_for_crs_page(spec_id)
            open_url_and_copy_to_clipboard(url_to_open)


def get_url_for_version_text(spec_entries: pd.DataFrame, version_text: str) -> str:
    """
    Returns the URL for the matching version. It is assumed that all rows in the DataFrame are for a single Spec.
    Args:
        spec_entries: DataFrame containing entries for a given specification
        version_text: The version to be retrieved, e.g. 16.0.0

    Returns:
        The URL of the given specification/version.
    """
    file_version = version_to_file_version(version_text)

    # Because of using '[[]]', it is sure that the returned object is a DataFrame and not a Series
    entry_to_load = spec_entries.loc[spec_entries.version == file_version, ['spec_url']]
    entry_to_load = entry_to_load.iloc[0]
    return entry_to_load.spec_url

