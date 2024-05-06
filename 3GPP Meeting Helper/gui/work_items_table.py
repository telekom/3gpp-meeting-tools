import re
import textwrap
import tkinter
from tkinter import ttk
from typing import List, Set

from application.os import open_url
from gui.generic_table import GenericTable, set_column, treeview_set_row_formatting
from server import wi_search
from server.connection import get_html
from server.tdoc_search import search_download_and_open_tdoc
from server.wi_search import WiEntry, wgs_list, download_wi_list
from tdoc.utils import tdoc_generic_regex

# To avoid matches for things like "utf-8"
tdoc_id_match_regex = re.compile(r'contributionUid=(' + tdoc_generic_regex.pattern + r')')


class WorkItemsTable(GenericTable):

    def __init__(self, parent, favicon, parent_gui_tools):
        super().__init__(
            parent,
            "Work Items Table. Double-click UID for WI page",
            favicon,
            ['UID', 'Code', 'Title', 'Release', 'Lead body', 'WID', 'Specs', 'CRs']
        )
        self.release_list: List[str] | None = [f'Rel-{rel_number}' for rel_number in range(5, 19, 1)]
        self.loaded_work_item_entries: List[WiEntry] | None = None
        self.filtered_work_item_entries: List[WiEntry] | None = None
        self.parent_gui_tools = parent_gui_tools

        self.wi_count = tkinter.StringVar()

        set_column(self.tree, 'UID', width=90, center=True)
        set_column(self.tree, 'Code', width=230, center=False)
        set_column(self.tree, 'Title', width=575, center=False)
        set_column(self.tree, 'Release', width=90, center=True)
        set_column(self.tree, 'Lead body', width=100, center=True)
        set_column(self.tree, 'WID', width=50, center=True)
        set_column(self.tree, 'Specs', width=50, center=True)
        set_column(self.tree, 'CRs', width=50, center=True)

        self.tree.bind("<Double-Button-1>", self.on_double_click)
        column_separator_str = "     "

        # Filter by group
        all_groups = ['All']
        all_groups.extend(wgs_list)
        self.combo_groups = ttk.Combobox(self.top_frame, values=all_groups, state="readonly", width=6)
        self.combo_groups.set('All')
        self.combo_groups.bind("<<ComboboxSelected>>", self.select_groups)

        # Filter by 3GPP Group/WG
        tkinter.Label(self.top_frame, text="Group: ").pack(side=tkinter.LEFT)
        self.combo_groups.pack(side=tkinter.LEFT)

        # Filter by release
        all_releases = ['All']
        all_releases.extend(self.release_list)
        self.combo_releases = ttk.Combobox(self.top_frame, values=all_releases, state="readonly", width=6)
        self.combo_releases.set('All')
        self.combo_releases.bind("<<ComboboxSelected>>", self.select_releases)

        # Filter by 3GPP Release
        tkinter.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        tkinter.Label(self.top_frame, text="Release: ").pack(side=tkinter.LEFT)
        self.combo_releases.pack(side=tkinter.LEFT)

        # Open/search TDoc
        tkinter.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        self.tkvar_wi_name = tkinter.StringVar(self.top_frame)
        self.tkvar_wi_name.trace_add('write', self.on_wi_search_change)
        self.wi_entry = tkinter.Entry(self.top_frame, textvariable=self.tkvar_wi_name, width=15, font='TkDefaultFont')

        self.wi_entry.pack(side=tkinter.LEFT)
        tkinter.Label(self.top_frame, text="  ").pack(side=tkinter.LEFT)

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

        def _sort_rel_str(rel_str) -> int:
            return_value = 0
            if 'R99' in rel_str:
                return_value = 1
            rel_match = re.match(r'Rel[\-‑]([\d]+)', rel_str)
            if rel_match is not None:
                return_value = int(rel_match.group(1))
            # print(f'{rel_str}->{return_value}')
            return return_value

        self.release_list = ['All']
        release_list_from_items = list({wi.release for wi in self.loaded_work_item_entries})
        release_list_from_items.sort(key=_sort_rel_str, reverse=True)
        self.release_list.extend(release_list_from_items)
        self.combo_releases['values'] = self.release_list

        print('Finished loading WIs')
        self.apply_filters()
        self.insert_rows()

    def insert_rows(self):
        wi_list_to_consider = self.filtered_work_item_entries
        print(f'Populating WI table from {len(wi_list_to_consider)} WIs')

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
                textwrap.fill(wi.title, width=75),
                wi.release,
                textwrap.fill(wi.lead_body, width=10),
                'Click',
                'Click',
                'Click'
            ))

        treeview_set_row_formatting(self.tree)
        self.wi_count.set('{0} WIs'.format(count))

    def clear_filters(self, *args):
        self.combo_groups.set('All')

        # Refill list
        self.apply_filters()

    def apply_filters(self):
        self.tree.delete(*self.tree.get_children())

        # Filter by selected group
        wi_list_to_consider = self.loaded_work_item_entries
        selected_group = self.combo_groups.get()
        if selected_group != 'All':
            print(f'Filtering by group {selected_group}')
            wi_list_to_consider = [m for m in wi_list_to_consider if selected_group in m.lead_body]

        # Filter by selected release
        selected_release = self.combo_releases.get()
        if selected_release != 'All':
            print(f'Filtering by release {selected_release}')
            wi_list_to_consider = [m for m in wi_list_to_consider if selected_release in m.release]

        # Filter by search string
        wi_search_str = self.tkvar_wi_name.get()
        if wi_search_str is not None and wi_search_str != '':
            print(f'Filtering WIs with code/title/UID {wi_search_str}')
            wi_search_str = wi_search_str.lower()
            wi_list_to_consider = [m for m in wi_list_to_consider if
                                   wi_search_str in m.code.lower() or
                                   wi_search_str in m.title.lower() or
                                   wi_search_str in m.uid.lower()]

        self.filtered_work_item_entries = wi_list_to_consider
        self.insert_rows()

    def select_groups(self, *args):
        self.apply_filters()

    def select_releases(self, *args):
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

        if column == 5:
            print(f'Clicked on WID {uid}. Will download latest WID version from {wi[0].wid_page_url}')
            url_to_open = wi[0].wid_page_url
            html_bytes = get_html(url_to_open)
            if html_bytes is None:
                print(f'Could not retrieve HTML for WID {uid}')
                return
            html_str = html_bytes.decode("utf-8")
            tdoc_match = tdoc_id_match_regex.search(html_str)
            if tdoc_match is None:
                print(f'Could not find WID in HTML for WID {uid}')
                return
            tdoc_id = tdoc_match.group(1)
            print(f'Last WID version is {tdoc_id}')
            search_download_and_open_tdoc(tdoc_id)

        if column == 6:
            print(f'Clicked on Spec list for WI {uid}')
            url_to_open = wi[0].spec_list_url
            open_url(url_to_open)

        if column == 7:
            print(f'Clicked on CR list for WI {uid}')
            url_to_open = wi[0].cr_list_url
            open_url(url_to_open)

    def on_wi_search_change(self, *args):
        self.apply_filters()
