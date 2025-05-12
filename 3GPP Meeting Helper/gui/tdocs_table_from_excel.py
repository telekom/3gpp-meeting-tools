import numbers
import os.path
import textwrap
import tkinter
from tkinter import ttk

import numpy as np
import pandas as pd
import pyperclip
from pandas import DataFrame

import server
import utils.local_cache
from application.excel import open_excel_document, set_autofilter_values, export_columns_to_markdown, clear_autofilter
from application.meeting_helper import tdoc_tags, open_sa2_session_plan_update_url
from application.os import open_url, startfile
from config.markdown import MarkdownConfig
from gui.common.common_elements import tkvar_3gpp_wifi_available
from gui.common.generic_table import GenericTable, treeview_set_row_formatting
from gui.common.generic_table import cloud_icon, cloud_download_icon
from server.common import WorkingGroup, get_document_or_folder_url, DocumentType, ServerType, get_tdoc_details_url, \
    MeetingEntry
from server.tdoc_search import batch_search_and_download_tdocs, search_meeting_for_tdoc
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
        self.tdocs_df: DataFrame = pd.read_excel(
            io=self.tdoc_excel_path,
            index_col=0)
        print(f'Imported meeting Tdocs for {meeting.meeting_name}: {self.tdocs_df.columns.values}')

        self.tdocs_df = self.tdocs_df.fillna(value='')
        self.tdocs_df['Secretary Remarks'] = self.tdocs_df['Secretary Remarks'].str.replace('<br/><br/>', '. ')
        self.meeting = meeting
        self.tdoc_tags = tdoc_tags
        self.tkvar_3gpp_wifi_available = tkvar_3gpp_wifi_available

        # Process tags
        self.tdoc_tag_list_str = ['All']
        if len(self.tdoc_tags) != 0:
            all_tags = list(set([s.tag for s in self.tdoc_tags]))
            all_tags.sort()
            self.tdoc_tag_list_str.extend(all_tags)

        self.tdocs_df['Tag'] = ''
        for tag in self.tdoc_tags:
            self.tdocs_df.loc[self.tdocs_df['Agenda item'] == tag.agenda_item, 'Tag'] = tag.tag

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

        try:
            self.tdocs_df['Sort Order'] = self.tdocs_df['Agenda item'].map(agenda_sort_item)
        except AttributeError as e:
            print(f"Could not order by Agenda Item: {e}")
        try:
            self.tdocs_df = self.tdocs_df.sort_values(by=[
                'Sort Order',
                self.tdocs_df.index.name])
        except KeyError as e:
            print(f"Could not order by Sort Order {e}")

        self.tdocs_current_df = self.tdocs_df

        # Fill in drop-down filters
        self.release_list = ['All']
        ai_items = self.tdocs_df['Release'].unique().tolist()
        ai_items.sort()
        self.release_list.extend(ai_items)
        self.ai_list = ['All']
        self.ai_list.extend(self.tdocs_df['Agenda item'].unique().tolist())

        self.tdoc_status_list = ['All']
        self.tdoc_status_list.extend(self.tdocs_df['TDoc Status'].unique().tolist())

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
            root_widget=root_widget,
            treeview_show=('headings', 'tree')
        )

        self.tree.column('#0', width=80, anchor='w')
        self.set_column('TDoc', "TDoc #", width=110)
        self.set_column('AI', width=52)
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
            width=9)
        self.combo_ai.set('All')
        self.combo_ai.bind("<<ComboboxSelected>>", self.select_rows)

        self.combo_status = ttk.Combobox(
            self.top_frame,
            values=self.tdoc_status_list,
            state="readonly",
            width=10)
        self.combo_status.set('All')
        self.combo_status.bind("<<ComboboxSelected>>", self.select_rows)

        self.combo_type = ttk.Combobox(
            self.top_frame,
            values=self.type_list,
            state="readonly",
            width=10)
        self.combo_type.set('All')
        self.combo_type.bind("<<ComboboxSelected>>", self.select_rows)

        self.combo_tag = ttk.Combobox(
            self.top_frame,
            values=self.tdoc_tag_list_str,
            state="readonly",
            width=10)
        self.combo_tag.set('All')
        self.combo_tag.bind("<<ComboboxSelected>>", self.select_rows)

        ttk.Label(self.top_frame, text="  AI: ").pack(side=tkinter.LEFT)
        self.combo_ai.pack(side=tkinter.LEFT)

        ttk.Label(self.top_frame, text="  Status: ").pack(side=tkinter.LEFT)
        self.combo_status.pack(side=tkinter.LEFT)

        ttk.Label(self.top_frame, text="  Rel.: ").pack(side=tkinter.LEFT)
        self.combo_release.pack(side=tkinter.LEFT)

        ttk.Label(self.top_frame, text="  Type: ").pack(side=tkinter.LEFT)
        self.combo_type.pack(side=tkinter.LEFT)

        ttk.Label(self.top_frame, text="  Tag: ").pack(side=tkinter.LEFT)
        self.combo_tag.pack(side=tkinter.LEFT)

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

        self.excel_to_markdown_btn = ttk.Button(
            self.top_frame,
            text='Excel2Markdown',
            command=self.current_excel_rows_to_clipboard,
            width=13
        )
        self.excel_to_markdown_btn.pack(side=tkinter.LEFT)

        self.download_btn = ttk.Button(
            self.top_frame,
            text='Download',
            command=self.download_tdocs,
            width=8
        )
        self.download_btn.pack(side=tkinter.LEFT)

        self.cache_btn = ttk.Button(
            self.top_frame,
            text='Local',
            command=lambda: startfile(meeting.local_folder_path),
            width=5
        )
        self.cache_btn.pack(side=tkinter.LEFT)

        self.markdown_export_per_ai_btn = ttk.Button(
            self.top_frame,
            text='Markdown/AI',
            command=self.export_ais_to_markdown,
            width=11
        )
        self.markdown_export_per_ai_btn.pack(side=tkinter.LEFT)

        # SA2-specific buttons
        if self.meeting.working_group_enum == WorkingGroup.S2 and self.meeting.meeting_is_now:
            ttk.Button(
                self.top_frame,
                text='Session Updates',
                command=lambda: startfile(
                    open_sa2_session_plan_update_url),
                width=13
            ).pack(side=tkinter.LEFT)

            def open_sa2_tdocsbyagenda():
                if tkvar_3gpp_wifi_available.get() and self.meeting.meeting_is_now:
                    server_type = ServerType.PRIVATE
                elif self.meeting.meeting_is_now:
                    server_type = ServerType.SYNC
                else:
                    server_type = ServerType.PUBLIC

                candidate_folders = get_document_or_folder_url(
                    server_type=server_type,
                    document_type=DocumentType.TDOCS_BY_AGENDA,
                    meeting_folder_in_server=self.meeting.meeting_folder,
                    working_group=WorkingGroup.S2
                )
                if len(candidate_folders) < 1:
                    return
                startfile(f'{candidate_folders[0]}TdocsByAgenda.htm')

            ttk.Button(
                self.top_frame,
                text='TDocsByAgenda',
                command=open_sa2_tdocsbyagenda,
                width=13
            ).pack(side=tkinter.LEFT)

        self.insert_rows()

        self.tree.pack(fill='both', expand=True, side='left')
        self.tree_scroll.pack(side=tkinter.RIGHT, fill='y')

        ttk.Label(self.bottom_frame, textvariable=self.tdoc_count).pack(side=tkinter.LEFT)

    def download_tdocs(self):
        tdoc_list = self.tdocs_current_df.index.tolist()
        if len(tdoc_list) > 0:
            batch_search_and_download_tdocs(tdoc_list)

        # Re-load Tdoc list to allow for icon changes
        self.insert_rows()

    def open_and_filter_excel(self):
        wb = open_excel_document(self.tdoc_excel_path)
        tdoc_list = self.tdocs_current_df.index.tolist()
        clear_autofilter(wb=wb)
        if len(tdoc_list) > 0:
            print(f'Filtering TDoc list for {len(tdoc_list)} TDocs shown')
            set_autofilter_values(wb=wb, value_list=tdoc_list)

    def current_excel_rows_to_clipboard(self):
        wb = open_excel_document(self.tdoc_excel_path)
        export_columns_to_markdown(wb, MarkdownConfig.columns_for_3gu_tdoc_export)

    def export_ais_to_markdown(self):
        # We want to take only the following:
        #   - LS IN regardless of status
        #   - LS OUT: only approved or agreed
        #   - CR: only approved or agreed
        #   - pCR: only approved or agreed

        is_cr = self.tdocs_df['Type'] == 'CR'
        is_pcr = self.tdocs_df['Type'] == 'pCR'
        is_ls_in = self.tdocs_df['Type'] == 'LS in'
        is_ls_out = self.tdocs_df['Type'] == 'LS out'

        is_approved_or_agreed = ((self.tdocs_df['TDoc Status'] == 'agreed') |
                                 (self.tdocs_df['TDoc Status'] == 'approved') |
                                 (self.tdocs_df['TDoc Status'] == 'endorsed'))

        pcrs_to_show = self.tdocs_df[(is_pcr & is_approved_or_agreed)]
        crs_to_show = self.tdocs_df[(is_cr & is_approved_or_agreed)]
        ls_to_show = self.tdocs_df[(is_ls_out & is_approved_or_agreed) | is_ls_in]

        company_contributions_filter = self.tdocs_df['Source'].str.contains(MarkdownConfig.company_name_regex_for_report)
        company_contributions = self.tdocs_df[company_contributions_filter]

        ls_out_to_show = self.tdocs_df[(is_ls_out & is_approved_or_agreed)]

        print(f'Will export:')
        print(f'  - {len(pcrs_to_show)} pCRs')
        print(f'  - {len(crs_to_show)} CRs')
        print(f'  - {len(ls_to_show)} LS IN/OUT')
        print(f'  - {len(ls_out_to_show)} LS OUT')
        print(f'  - {len(company_contributions)} Company contributions matching {MarkdownConfig.company_name_regex_for_report}')

        local_folder = self.meeting.local_export_folder_path
        wb = open_excel_document(self.tdoc_excel_path)
        clear_autofilter(wb=wb)

        ai_summary = dict()

        # Export LS OUT
        index_list = list(ls_out_to_show.index)
        ai_name = 'LS OUT'
        if len(index_list) > 0:
            print(f'{ai_name}: {len(index_list)} LS IN/OUT to export')
            set_autofilter_values(
                wb=wb,
                value_list=index_list,
                sort_by_sort_order_within_agenda_item=True)
            markdown_output = export_columns_to_markdown(
                wb,
                MarkdownConfig.columns_for_3gu_tdoc_export_ls_out,
                copy_output_to_clipboard=False)
            ai_summary[ai_name] = f'Following LS OUT were sent:\n\n{markdown_output}'
        else:
            print(f'{ai_name}: {len(index_list)} LS IN/OUT')

        # Export Company Contributions
        index_list = list(company_contributions.index)
        ai_name = 'Company'
        if len(index_list) > 0:
            print(
                f'{len(index_list)} Company contributions matching {MarkdownConfig.company_name_regex_for_report} to export')
            set_autofilter_values(
                wb=wb,
                value_list=index_list,
                sort_by_sort_order_within_agenda_item=True)
            markdown_output = export_columns_to_markdown(
                wb,
                MarkdownConfig.columns_for_3gu_tdoc_export_contributor,
                copy_output_to_clipboard=False)
            ai_summary[ai_name] = f'Following Company Contributions:\n\n{markdown_output}'
        else:
            print(f'{len(index_list)} Company contributions matching {MarkdownConfig.company_name_regex_for_report}')

        # Export LSs
        for ai_name, group in ls_to_show.groupby('Agenda item'):
            index_list = list(group.index)

            if len(index_list) > 0:
                print(f'{ai_name}: {len(index_list)} LS IN/OUT to export')
                set_autofilter_values(
                    wb=wb,
                    value_list=index_list,
                    sort_by_sort_order_within_agenda_item=True)
                markdown_output = export_columns_to_markdown(
                    wb,
                    MarkdownConfig.columns_for_3gu_tdoc_export_ls,
                    copy_output_to_clipboard=False)
                ai_summary[ai_name] = f'Following LS were received and/or answered:\n\n{markdown_output}'
            else:
                print(f'{ai_name}: {len(index_list)} LS IN/OUT')

        # Export pCRs
        for ai_name, group in pcrs_to_show.groupby('Agenda item'):
            index_list = list(group.index)

            if len(index_list) > 0:
                print(f'{ai_name}: {len(index_list)} pCRs to export')
                set_autofilter_values(
                    wb=wb,
                    value_list=index_list,
                    sort_by_sort_order_within_agenda_item=True)
                markdown_output = export_columns_to_markdown(
                    wb,
                    MarkdownConfig.columns_for_3gu_tdoc_export_pcr,
                    copy_output_to_clipboard=False)

                summary_text = f'Following pCRs were agreed:\n\n{markdown_output}'
                if ai_name in ai_summary:
                    ai_summary[ai_name] = f'{ai_summary[ai_name]}\n\n{summary_text}'
                else:
                    ai_summary[ai_name] = f'{summary_text}'

            else:
                print(f'{ai_name}: {len(index_list)} pCRs')

        # Export CRs
        for ai_name, group in crs_to_show.groupby('Agenda item'):
            index_list = list(group.index)

            if len(index_list) > 0:
                print(f'{ai_name}: {len(index_list)} CRs to export')
                set_autofilter_values(
                    wb=wb,
                    value_list=index_list,
                    sort_by_sort_order_within_agenda_item=True)
                markdown_output = export_columns_to_markdown(
                    wb,
                    MarkdownConfig.columns_for_3gu_tdoc_export_cr,
                    copy_output_to_clipboard=False)

                summary_text = f'Following CRs were agreed:\n\n{markdown_output}'
                if ai_name in ai_summary:
                    ai_summary[ai_name] = f'{ai_summary[ai_name]}\n\n{summary_text}'
                else:
                    ai_summary[ai_name] = f'{summary_text}'

            else:
                print(f'{ai_name}: {len(index_list)} CRs')

        for ai_name, summary_text in ai_summary.items():
            meeting_name_for_export = self.meeting.meeting_name.replace('3GPP','')
            summary_text = f'<!--- [{meeting_name_for_export}]({self.meeting.meeting_folder_url}) --->\n\n{summary_text}'
            with open(os.path.join(local_folder, f'{ai_name}.md'), 'w') as f:
                f.write(summary_text)

        print(f'Completed Markdown export for {self.meeting.meeting_name}')

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
                    opened_docs, metadata = server.tdoc_search.search_download_and_open_tdoc(
                        tdoc_to_open.tdoc,
                        tkvar_3gpp_wifi_available=tkvar_3gpp_wifi_available,
                        tdoc_meeting=self.meeting
                    )

                    # Re-load Tdoc list to allow for icon changes
                    self.insert_rows()

                    if metadata is not None:
                        try:
                            if isinstance(metadata, list):
                                metadata = metadata[0]
                            print(f'Opened Tdoc {metadata.tdoc_id}, {metadata.url}. Copied URL to clipboard')
                            pyperclip.copy(metadata.url)
                        except Exception as e:
                            print(f'Could not copy TDoc URL to clipboard: {e}')
                self.select_rows()
            case 5:

                TdocDetailsFromExcel(
                    favicon=self.favicon,
                    parent_widget=self.tk_top,
                    root_widget=self.root_widget,
                    tdoc_str=tdoc_id,
                    tdoc_row=self.tdocs_df.loc[tdoc_id],
                    meeting=self.meeting)

    def select_rows(self, *args):
        filter_str = self.search_text.get()
        filtered_df = self.tdocs_df

        if filter_str is not None and filter_str != '':
            print(f'Filtering by "{filter_str}"')

            try:
                filter_str_float = float(filter_str)
            except ValueError:
                filter_str_float = float('nan')

            # Search in TDoc ID and title
            filtered_df = filtered_df[
                filtered_df.index.str.contains(filter_str, case=False) |
                filtered_df["Title"].str.contains(filter_str, case=False) |
                filtered_df["Related WIs"].str.contains(filter_str, case=False) |
                filtered_df["Source"].str.contains(filter_str, case=False) |
                filtered_df["Spec"].str.contains(filter_str, case=False) |
                (filtered_df["CR"] == filter_str_float) |
                filtered_df["Secretary Remarks"].str.contains(filter_str, case=False)]

        ai_filter = self.combo_ai.get()
        if ai_filter != 'All':
            print(f'Filtering by AI: "{ai_filter}"')
            filtered_df = filtered_df[filtered_df["Agenda item"] == ai_filter]

        status_filter = self.combo_status.get()
        if status_filter != 'All':
            print(f'Filtering by TDoc status: "{status_filter}"')
            filtered_df = filtered_df[filtered_df["TDoc Status"] == status_filter]

        type_filter = self.combo_type.get()
        if type_filter != 'All':
            print(f'Filtering by Type: "{type_filter}"')
            filtered_df = filtered_df[filtered_df["Type"] == type_filter]

        rel_filter = self.combo_release.get()
        if rel_filter != 'All':
            print(f'Filtering by Release: "{rel_filter}"')
            filtered_df = filtered_df[filtered_df["Release"] == rel_filter]

        tag_filter = self.combo_tag.get()
        if tag_filter != 'All':
            print(f'Filtering by Tag: "{tag_filter}"')
            filtered_df = filtered_df[filtered_df["Tag"] == tag_filter]

        self.tdocs_current_df = filtered_df
        self.insert_rows()

    def insert_rows(self):
        print('(Re-)Populating TDocs table')
        self.tree.delete(*self.tree.get_children())
        count = 0

        for tdoc_id, row in self.tdocs_current_df.iterrows():
            count = count + 1
            mod = count % 2
            if mod > 0:
                tag = 'odd'
            else:
                tag = 'even'

            local_file = self.meeting.get_tdoc_local_path(str(tdoc_id))

            # Icon to show if a TDoc was downloaded
            if utils.local_cache.file_exists(local_file):
                row_icon = cloud_download_icon
            else:
                row_icon = cloud_icon

            # Text including spec number
            tdoc_title = row['Title']
            try:
                tdoc_spec = row['Spec']
                try:
                    cr_number = row['CR']
                    if cr_number is None or cr_number == '':
                        cr_number = ''
                    else:
                        cr_number = f"CR{cr_number:.0f}"

                    try:
                        tdoc_release = row['Release']  # Rel-18
                        cr_category = f"Cat-{row['CR category']}"  # F
                    except Exception as e:
                        print(f'Could not retrieve Rel/CR category: {e}')
                        tdoc_release = ''
                        cr_category = ''
                except Exception as e:
                    print(f'Could not retrieve CR number: {e}')
                    cr_number = ''
                    tdoc_release = ''
                    cr_category = ''

                if tdoc_spec is not None and tdoc_spec != '':
                    tdoc_title = f'{tdoc_spec}{cr_number}: {tdoc_title}'

                if tdoc_release is not None and tdoc_release != '':
                    if cr_category is not None and cr_category != '' and cr_number != '':
                        tdoc_title = f'{tdoc_title} ({tdoc_release}, {cr_category})'
                    else:
                        tdoc_title = f'{tdoc_title} ({tdoc_release})'
            except Exception as e:
                print(f'Could not retrieve TDoc title: {e}')

            self.tree.insert(
                "",
                "end",
                tags=(tag,),
                values=(
                    tdoc_id,
                    row['Agenda item'],
                    row['Type'],
                    textwrap.fill(tdoc_title, width=70),
                    textwrap.fill(row['Source'], width=25),
                    'Click',
                    textwrap.fill(row['Secretary Remarks'], width=50)
                ),
                image=row_icon,
            )

            treeview_set_row_formatting(self.tree)
            self.tdoc_count.set('{0} documents'.format(count))


class TdocDetailsFromExcel(GenericTable):
    def __init__(
            self,
            favicon,
            parent_widget: tkinter.Tk,
            tdoc_str: str,
            tdoc_row,
            meeting: MeetingEntry,
            root_widget: tkinter.Tk | None = None,
    ):
        self.tdoc_id = tdoc_str
        self.tdoc_row = tdoc_row
        self.tkvar_3gpp_wifi_available = tkvar_3gpp_wifi_available
        self.meeting = meeting

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
            '3GU Link',
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
            'Reply in',
            'URL'
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
                case 'URL':
                    try:
                        tdoc_meeting = search_meeting_for_tdoc(self.tdoc_id, return_last_meeting_if_tdoc_is_new=True)
                        row_value = tdoc_meeting.get_tdoc_url(self.tdoc_id)
                    except Exception as e:
                        row_value = ''
                        print(f'Could not generate TDoc URL for {self.tdoc_id}: {e}')
                case '3GU Link':
                    row_value = 'Click!'
                case _:
                    try:
                        row_value = self.tdoc_row[row_name]
                    except KeyError:
                        row_value = ''
                        print(f'Could not read column {row_name}')

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
                    opened_docs, metadata = server.tdoc_search.search_download_and_open_tdoc(
                        tdoc_to_open.tdoc,
                        tkvar_3gpp_wifi_available=tkvar_3gpp_wifi_available)
                    if metadata is not None:
                        print(f'Opened Tdoc {metadata[0].tdoc_id}, {metadata[0].url}. Copied URL to clipboard')
                        pyperclip.copy(metadata[0].url)
            case "Contact":
                person_id = self.tdoc_row["Contact ID"]
                url_to_open = f'https://webapp.etsi.org/teldir/ListPersDetails.asp?PersId={person_id}'
                open_url(url_to_open)
            case '3GU Link':
                url_to_open = get_tdoc_details_url(self.tdoc_id)
                open_url(url_to_open)
            case 'URL':
                url_value = actual_value
                pyperclip.copy(url_value)
                print(f'Copied URL for {self.tdoc_id} to clipboard: {url_value}')
