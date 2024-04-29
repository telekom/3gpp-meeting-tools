import os.path
import tkinter
from tkinter import ttk
from typing import List

import server
from application.excel import open_excel_document
from application.os import open_url
from gui.generic_table import GenericTable, set_column, treeview_set_row_formatting
from server import tdoc_search, wi_search
from server.common import download_file_to_location
from server.wi_search import WiEntry, wgs_list, download_wi_list
from utils.local_cache import file_exists


class WorkItemsTable(GenericTable):

    def __init__(self, parent, favicon, parent_gui_tools):
        super().__init__(
            parent,
            "Work Items Table. Double-click UID for WI page, on Code for WI CRs",
            favicon,
            ['UID', 'Code', 'Title', 'Release', 'Lead body']
        )
        self.loaded_work_item_entries: List[WiEntry] | None = None
        self.parent_gui_tools = parent_gui_tools

        self.wi_count = tkinter.StringVar()

        set_column(self.tree, 'UID', width=90, center=True)
        set_column(self.tree, 'Code', width=230, center=True)
        set_column(self.tree, 'Title', width=575, center=True)
        set_column(self.tree, 'Release', width=90, center=True)
        set_column(self.tree, 'Lead body', width=100, center=True)

        self.tree.bind("<Double-Button-1>", self.on_double_click)
        column_separator_str = "     "

        # Filter by group (only filter we have in this view)
        all_groups = ['All']
        all_groups.extend(wgs_list)
        self.combo_groups = ttk.Combobox(self.top_frame, values=all_groups, state="readonly", width=6)
        self.combo_groups.set('All')
        self.combo_groups.bind("<<ComboboxSelected>>", self.select_groups)

        # Filter by 3GPP Group/WG
        tkinter.Label(self.top_frame, text="Group: ").pack(side=tkinter.LEFT)
        self.combo_groups.pack(side=tkinter.LEFT)

        # Redownload WI list if it already exists
        self.redownload_wi_list_if_exists_var = tkinter.IntVar()
        self.redownload_wi_list_if_exists = ttk.Checkbutton(
            self.top_frame,
            state='enabled',
            variable=self.redownload_wi_list_if_exists_var)
        tkinter.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        tkinter.Label(self.top_frame, text="Re-download WI list if exists: ").pack(side=tkinter.LEFT)
        self.redownload_wi_list_if_exists.pack(side=tkinter.LEFT)

        # Load meeting data
        tkinter.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        tkinter.Button(
            self.top_frame,
            text='Load WIs',
            command=self.load_data).pack(side=tkinter.LEFT)

        # Main frame
        self.load_data()
        self.insert_rows()

        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

        # Bottom frame
        tkinter.Label(self.bottom_frame, textvariable=self.wi_count).pack(side=tkinter.LEFT)

        # Add text wrapping
        # https: // stackoverflow.com / questions / 51131812 / wrap - text - inside - row - in -tkinter - treeview

    def load_data(self):
        """
        Loads Work Items from the 3GPP website
        """
        # Load specs data
        if self.redownload_wi_list_if_exists_var.get():
            print('Loading 3GPP WI data with re-download')
            download_wi_list(re_download_if_exists=True)
        else:
            print('Loading 3GPP WI data using cache')
            download_wi_list(re_download_if_exists=False)

        wi_search.load_wi_entries()
        self.loaded_work_item_entries = wi_search.loaded_wi_entries
        print('Finished loading WIs')
        self.insert_rows()

    def insert_rows(self):
        wi_list_to_consider = self.loaded_work_item_entries
        print(f'Populating WI table from {len(wi_list_to_consider)} WIs')

        # Filter by selected group
        selected_group = self.combo_groups.get()
        if selected_group != 'All':
            wi_list_to_consider = [m for m in wi_list_to_consider if m.lead_body == selected_group]

        # Sort list by date
        wi_list_to_consider.sort(reverse=True, key=lambda m: m.uid)

        count = 0
        for idx, wi in enumerate(wi_list_to_consider):
            count = count + 1
            mod = count % 2
            if mod > 0:
                tag = 'odd'
            else:
                tag = 'even'

            # 'Meeting', 'Location', 'Start', 'End', 'TDoc Start', 'TDoc End', 'Documents'
            self.tree.insert("", "end", tags=(tag,), values=(
                wi.uid,
                wi.code,
                wi.title,
                wi.release,
                wi.lead_body
            ))

        treeview_set_row_formatting(self.tree)
        self.wi_count.set('{0} WIs'.format(count))

    def clear_filters(self, *args):
        self.combo_groups.set('All')

        # Refill list
        self.apply_filters()

    def apply_filters(self):
        self.tree.delete(*self.tree.get_children())
        self.insert_rows()

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

        print("You clicked on {0}/{1}: {2}".format(event.x, event.y, actual_value))
        uid = str(item_values[0])
        wi = [m for m in self.loaded_work_item_entries if uid in m.uid]
        print(f'UID of the clicked row is {uid}. Found {len(wi)} matching WIs')

        if actual_value is None or actual_value == '':
            print("Empty value")
            return

        if column == 0:
            print(f'Clicked on WI {uid}')
            url_to_open = wi[0].wid_page_url
            open_url(url_to_open)

        if column == 1:
            print(f'Clicked on CR list for WI {uid}')
            url_to_open = wi[0].cr_list_url
            open_url(url_to_open)

    def on_open_tdoc(self):
        tdoc_to_open = self.tkvar_tdoc_id.get()
        print(f'Opening {tdoc_to_open}')
        server.tdoc_search.search_download_and_open_tdoc(tdoc_to_open)
