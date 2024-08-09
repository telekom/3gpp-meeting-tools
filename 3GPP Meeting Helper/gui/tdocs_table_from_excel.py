import numbers
import os
import textwrap
import tkinter
from tkinter import ttk

import numpy as np
import pandas as pd
import pyperclip
from pandas import DataFrame

import server
from application.excel import open_excel_document, set_autofilter_values
from application.os import open_url
from gui.common.generic_table import GenericTable, treeview_set_row_formatting
from server.tdoc_search import MeetingEntry, batch_search_and_download_tdocs
from tdoc.utils import are_generic_tdocs


class TdocsTableFromExcel(GenericTable):
    source_width = 200
    title_width = 550

    def __init__(
            self,
            favicon,
            parent_widget: tkinter.Tk,
            meeting: MeetingEntry,
            tdoc_excel_path: str,
            root_widget: tkinter.Tk | None = None,
    ):
        """
        Opens the TDoc table
        Args:
            favicon: The favicon to use
            parent_widget: The parent widget
            meeting: The meeting
            tdoc_excel_path: Path to the Excel file from the 3GPP server containing the TDoc list
        """
        self.tdoc_excel_path = tdoc_excel_path

        # Load and cleanup DataFrame
        self.tdocs_df: DataFrame = pd.read_excel(io=self.tdoc_excel_path, index_col=0)
        self.tdocs_df = self.tdocs_df.fillna(value='')
        self.tdocs_df['Secretary Remarks'] = self.tdocs_df['Secretary Remarks'].str.replace('<br/><br/>', '. ')
        self.meeting = meeting

        def agenda_sort_item(input_str):
            if input_str is None or input_str == np.nan or input_str == '':
                return 0
            input_split = [int(i) for i in input_str.split('.')]
            out_value = input_split[0] * 1000
            if len(input_split) > 1:
                out_value = out_value + input_split[1] * 10
            if len(input_split) > 2:
                out_value = out_value + input_split[2]
            return out_value

        self.tdocs_df['Sort Order'] = self.tdocs_df['Agenda item'].map(agenda_sort_item)

        self.tdocs_df = self.tdocs_df.sort_values(by=[
            'Sort Order',
            self.tdocs_df.index.name])
        self.tdocs_current_df = self.tdocs_df

        # Fill in drop-down filters
        self.release_list = ['All']
        ai_items = self.tdocs_df['Release'].unique().tolist()
        ai_items.sort()
        self.release_list.extend(ai_items)
        self.ai_list = ['All']
        self.ai_list.extend(self.tdocs_df['Agenda item'].unique().tolist())
        self.type_list = ['All']
        type_items = self.tdocs_df['Type'].unique().tolist()
        type_items.sort()
        self.type_list.extend(type_items)

        # Document counter
        self.tdoc_count = tkinter.StringVar()

        super().__init__(
            parent_widget=parent_widget,
            widget_title=f"{meeting.meeting_name} ({meeting.meeting_location}). Double-Click on TDoc # or secretary remarks # to open",
            favicon=favicon,
            column_names=[
                'TDoc',
                'AI', 'Type',
                'Title',
                'Source',
                'Details',
                'Secretary Remarks'],
            row_height=60,
            display_rows=9,
            root_widget=root_widget
        )

        self.set_column('TDoc', "TDoc #", width=110)
        self.set_column('AI', width=50)
        self.set_column('Type', width=120)
        self.set_column('Title', width=TdocsTableFromExcel.title_width, center=False)
        self.set_column('Source', width=TdocsTableFromExcel.source_width, center=False)
        self.set_column('Details', width=75)
        self.set_column('Secretary Remarks', width=400, center=False)

        self.search_text = tkinter.StringVar()
        self.search_entry = tkinter.Entry(
            self.top_frame,
            textvariable=self.search_text,
            width=25,
            font='TkDefaultFont')
        self.search_text.trace_add(['write', 'unset'], self.select_rows)
        ttk.Label(self.top_frame, text="Search: ").pack(side=tkinter.LEFT)
        self.search_entry.pack(
            side=tkinter.LEFT,
            pady=10)

        # Drop-down lists
        self.combo_release = ttk.Combobox(
            self.top_frame,
            values=self.release_list,
            state="readonly",
            width=10)
        self.combo_release.set('All')
        self.combo_release.bind("<<ComboboxSelected>>", self.select_rows)

        self.combo_ai = ttk.Combobox(
            self.top_frame,
            values=self.ai_list,
            state="readonly",
            width=10)
        self.combo_ai.set('All')
        self.combo_ai.bind("<<ComboboxSelected>>", self.select_rows)

        self.combo_type = ttk.Combobox(
            self.top_frame,
            values=self.type_list,
            state="readonly",
            width=10)
        self.combo_type.set('All')
        self.combo_type.bind("<<ComboboxSelected>>", self.select_rows)

        ttk.Label(self.top_frame, text="  AI: ").pack(side=tkinter.LEFT)
        self.combo_ai.pack(side=tkinter.LEFT)

        ttk.Label(self.top_frame, text="  Release: ").pack(side=tkinter.LEFT)
        self.combo_release.pack(side=tkinter.LEFT)

        ttk.Label(self.top_frame, text="  Type: ").pack(side=tkinter.LEFT)
        self.combo_type.pack(side=tkinter.LEFT)

        self.tree.bind("<Double-Button-1>", self.on_double_click)

        self.open_excel_btn = ttk.Button(
            self.top_frame,
            text='Open Excel',
            command=lambda: open_excel_document(self.tdoc_excel_path),
            width=9
        )
        self.open_excel_btn.pack(side=tkinter.LEFT)

        self.open_excel_btn = ttk.Button(
            self.top_frame,
            text='Filter Excel',
            command=self.open_and_filter_excel,
            width=8
        )
        self.open_excel_btn.pack(side=tkinter.LEFT)

        self.download_btn = ttk.Button(
            self.top_frame,
            text='Download',
            command=self.download_tdocs,
            width=8
        )
        self.download_btn.pack(side=tkinter.LEFT)

        self.cache_btn = ttk.Button(
            self.top_frame,
            text='Cache',
            command=lambda: os.startfile(meeting.local_folder_path),
            width=5
        )
        self.cache_btn.pack(side=tkinter.LEFT)

        self.insert_rows()

        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

        ttk.Label(self.bottom_frame, textvariable=self.tdoc_count).pack(side=tkinter.LEFT)

    def download_tdocs(self):
        tdoc_list = self.tdocs_current_df.index.tolist()
        if len(tdoc_list) > 0:
            batch_search_and_download_tdocs(tdoc_list)

    def open_and_filter_excel(self):
        wb = open_excel_document(self.tdoc_excel_path)
        tdoc_list = self.tdocs_current_df.index.tolist()
        if len(tdoc_list) > 0:
            print(f'Filtering TDoc list for {len(tdoc_list)} TDocs shown')
            set_autofilter_values(wb=wb, value_list=tdoc_list)

    def on_double_click(self, event):
        item_id = self.tree.identify("item", event.x, event.y)
        column = int(self.tree.identify_column(event.x)[1:]) - 1  # "#1" -> 0 (one-indexed in TCL)
        item_values = self.tree.item(item_id)['values']
        try:
            actual_value = item_values[column]
        except Exception as e:
            print(f'Could not process TreeView double-click: {e}')
            actual_value = None

        if actual_value is None or actual_value == '':
            return

        tdoc_id = item_values[0]

        match column:
            case 0 | 6:
                print(f'Clicked on TDoc {actual_value}. Row: {tdoc_id}')
                tdocs_to_open = are_generic_tdocs(actual_value)
                for tdoc_to_open in tdocs_to_open:
                    print(f'Opening {tdoc_to_open.tdoc}')
                    opened_docs, metadata = server.tdoc_search.search_download_and_open_tdoc(tdoc_to_open.tdoc)
                    if metadata is not None:
                        print(f'Opened Tdoc {metadata[0].tdoc_id}, {metadata[0].url}. Copied URL to clipboard')
                        pyperclip.copy(metadata[0].url)
            case 5:

                TdocDetailsFromExcel(
                    favicon=self.favicon,
                    parent_widget=self.tk_top,
                    root_widget=self.root_widget,
                    tdoc_str=tdoc_id,
                    tdoc_row=self.tdocs_df.loc[tdoc_id])

    def select_rows(self, *args):
        filter_str = self.search_text.get()
        filtered_df = self.tdocs_df

        if filter_str is not None and filter_str != '':
            print(f'Filtering by "{filter_str}"')

            # Search in TDoc ID and title
            filtered_df = filtered_df[
                filtered_df.index.str.contains(filter_str, case=False) |
                filtered_df["Title"].str.contains(filter_str, case=False) |
                filtered_df["Related WIs"].str.contains(filter_str, case=False) |
                filtered_df["Source"].str.contains(filter_str, case=False) |
                filtered_df["Secretary Remarks"].str.contains(filter_str, case=False)]

        ai_filter = self.combo_ai.get()
        if ai_filter != 'All':
            print(f'Filtering by AI: "{ai_filter}"')
            filtered_df = filtered_df[filtered_df["Agenda item"] == ai_filter]

        type_filter = self.combo_type.get()
        if type_filter != 'All':
            print(f'Filtering by Type: "{type_filter}"')
            filtered_df = filtered_df[filtered_df["Type"] == type_filter]

        rel_filter = self.combo_release.get()
        if rel_filter != 'All':
            print(f'Filtering by Release: "{rel_filter}"')
            filtered_df = filtered_df[filtered_df["Release"] == rel_filter]

        self.tdocs_current_df = filtered_df
        self.tree.delete(*self.tree.get_children())
        self.insert_rows()

    def insert_rows(self):
        print('Populating TDocs table')
        count = 0
        for idx, row in self.tdocs_current_df.iterrows():
            count = count + 1
            mod = count % 2
            if mod > 0:
                tag = 'odd'
            else:
                tag = 'even'

            self.tree.insert("", "end", tags=(tag,), values=(
                idx,
                row['Agenda item'],
                row['Type'],
                textwrap.fill(row['Title'], width=70),
                textwrap.fill(row['Source'], width=25),
                'Click',
                textwrap.fill(row['Secretary Remarks'], width=50)))

            treeview_set_row_formatting(self.tree)
            self.tdoc_count.set('{0} documents'.format(count))


class TdocDetailsFromExcel(GenericTable):
    def __init__(
            self,
            favicon,
            parent_widget: tkinter.Tk,
            tdoc_str: str,
            tdoc_row,
            root_widget: tkinter.Tk | None = None,
    ):
        self.tdoc_id = tdoc_str
        self.tdoc_row = tdoc_row

        super().__init__(
            parent_widget=parent_widget,
            widget_title=f"{tdoc_str}",
            favicon=favicon,
            column_names=[
                'Info',
                'Content'],
            row_height=30,
            display_rows=9,
            root_widget=root_widget
        )

        self.set_column('Info', width=250, center=False)
        self.set_column('Content', width=1500, center=False)

        self.tree.bind("<Double-Button-1>", self.on_double_click)
        self.insert_rows()

        (ttk.Label(
            self.top_frame,
            text=textwrap.fill(
                f"Abstract: {self.tdoc_row['Abstract']}",
                width=240))
        .pack(
            side=tkinter.LEFT,
            pady=10)
        )

        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

    def insert_rows(self):
        print('Populating TDocs table')
        count = 0
        for row_name in [
            'TDoc',
            'Title',
            'Source',
            'Contact',
            'Type',
            'For',
            'Secretary Remarks',
            'Agenda item',
            'Agenda item description',
            'TDoc Status',
            'Is revision of',
            'Revised to',
            'Release',
            'Spec',
            'Version',
            'Related WIs',
            'CR',
            'CR revision',
            'CR category',
            'TSG CR Pack',
            'UICC',
            'ME',
            'RAN',
            'CN',
            'Clauses Affected',
            'Reply to',
            'To',
            'Cc',
            'Original LS',
            'Reply in'
        ]:

            count = count + 1
            mod = count % 2
            if mod > 0:
                tag = 'odd'
            else:
                tag = 'even'

            match row_name:
                case 'TDoc':
                    row_value = self.tdoc_id
                case 'CR' | 'CR revision':
                    row_value = self.tdoc_row[row_name]
                    if isinstance(row_value, numbers.Number):
                        row_value = f'{row_value:0.0f}'
                case _:
                    row_value = self.tdoc_row[row_name]

            self.tree.insert("", "end", tags=(tag,), values=(
                row_name,
                textwrap.fill(str(row_value), width=250)))

            treeview_set_row_formatting(self.tree)

    def on_double_click(self, event):
        item_id = self.tree.identify("item", event.x, event.y)
        column = int(self.tree.identify_column(event.x)[1:]) - 1  # "#1" -> 0 (one-indexed in TCL)

        # Trigger only if the descriptions are clicked
        if column != 1:
            return

        item_values = self.tree.item(item_id)['values']
        try:
            actual_value = item_values[column]
        except Exception as e:
            print(f'Could not process TreeView double-click: {e}')
            actual_value = None

        if actual_value is None or actual_value == '':
            return

        row_name = item_values[0]
        match row_name:
            case 'TDoc' | 'Secretary Remarks':
                print(f'Clicked on TDoc {actual_value}')
                tdocs_to_open = are_generic_tdocs(actual_value)
                for tdoc_to_open in tdocs_to_open:
                    print(f'Opening {tdoc_to_open.tdoc}')
                    opened_docs, metadata = server.tdoc_search.search_download_and_open_tdoc(tdoc_to_open.tdoc)
                    if metadata is not None:
                        print(f'Opened Tdoc {metadata[0].tdoc_id}, {metadata[0].url}. Copied URL to clipboard')
                        pyperclip.copy(metadata[0].url)
            case "Contact":
                person_id = self.tdoc_row["Contact ID"]
                url_to_open = f'https://webapp.etsi.org/teldir/ListPersDetails.asp?PersId={person_id}'
                open_url(url_to_open)
