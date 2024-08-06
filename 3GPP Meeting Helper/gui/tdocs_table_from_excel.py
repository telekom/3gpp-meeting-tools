import textwrap
import tkinter
from tkinter import ttk

import numpy as np
import pandas as pd
import pyperclip
from pandas import DataFrame

import server
from gui.common.generic_table import GenericTable, treeview_set_row_formatting
from server.tdoc_search import MeetingEntry


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
        self.tdocs_df: DataFrame = pd.read_excel(io=self.tdoc_excel_path, index_col=0)
        self.tdocs_df = self.tdocs_df.fillna(value='')

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
        self.tdoc_count = tkinter.StringVar()

        super().__init__(
            parent_widget=parent_widget,
            widget_title=f"{meeting.meeting_name} TDocs. Double-Click on TDoc # or revision # to open",
            favicon=favicon,
            column_names=[
                'TDoc',
                'AI', 'Type',
                'Title',
                'Source',
                'Details',
                'Rev. of',
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
        self.set_column('Rev. of', width=100)
        self.set_column('Secretary Remarks', width=300)

        self.tree.bind("<Double-Button-1>", self.on_double_click)

        self.insert_rows()

        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

        ttk.Label(self.bottom_frame, textvariable=self.tdoc_count).pack(side=tkinter.LEFT)

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

        match column:
            case 0 | 6:
                print(f'Clicked on TDoc {actual_value}. Row: {item_values[0]}')
                tdoc_to_open = actual_value
                print(f'Opening {tdoc_to_open}')
                opened_docs, metadata = server.tdoc_search.search_download_and_open_tdoc(tdoc_to_open)
                if metadata is not None:
                    print(f'Opened Tdoc {metadata[0].tdoc_id}, {metadata[0].url}. Copied URL to clipboard')
                    pyperclip.copy(metadata[0].url)

    def insert_rows(self):
        print('Populating TDocs table')
        count = 0
        for idx, row in self.tdocs_df.iterrows():
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
                row['Is revision of'],
                textwrap.fill(row['Secretary Remarks'], width=35)))

            treeview_set_row_formatting(self.tree)
            self.tdoc_count.set('{0} documents'.format(count))
