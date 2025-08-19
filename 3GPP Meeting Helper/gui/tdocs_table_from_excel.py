import datetime
import json
import numbers
import os.path
import shutil
import textwrap
import tkinter
from pathlib import Path
from tkinter import ttk
from typing import List, Any, NamedTuple

import numpy as np
import pyperclip
from pandas import DataFrame
from pypdf import PdfWriter
from pypdf.generic import PAGE_FIT

import server
import utils.caching.common
import utils.local_cache
from application.common import ActionAfter, ExportType
from application.excel import open_excel_document, set_autofilter_values, export_columns_to_markdown, clear_autofilter, \
    export_columns_to_markdown_dataframe
from application.meeting_helper import tdoc_tags, open_sa2_drafts_url
from application.os import open_url, startfile
from application.word import export_document
from config.markdown import MarkdownConfig
from gui.common.common_elements import tkvar_3gpp_wifi_available
from gui.common.generic_table import GenericTable, treeview_set_row_formatting, column_separator_str
from gui.common.gui_elements import TTKHoverHelpButton
from gui.common.icons import cloud_icon, cloud_download_icon, folder_icon, share_icon, excel_icon, website_icon, \
    filter_icon, note_icon, ftp_icon, markdown_icon, share_markdown_icon
from server.common.MeetingEntry import MeetingEntry
from server.common.server_utils import get_document_or_folder_url, get_tdoc_details_url, \
    DownloadedTdocDocument
from server.common.server_utils import ServerType, DocumentType, WorkingGroup
from server.tdoc_search import batch_search_and_download_tdocs, search_meeting_for_tdoc
from tdoc.utils import are_generic_tdocs
from utils.local_cache import create_folder_if_needed
from utils.utils import invert_dict_defaultdict


def df_boolean_index_for_wi(in_df: DataFrame, wi_str:str):
    filter_idx = ((in_df["Related WIs"] == wi_str) |
                  in_df["Related WIs"].str.startswith(f'{wi_str}, ') |
                  in_df["Related WIs"].str.contains(f', {wi_str},') |
                  in_df["Related WIs"].str.endswith(f', {wi_str}'))
    return filter_idx


def get_markdown_for_tdocs(
        filtered_df: DataFrame,
        column_list,
        the_meeting:MeetingEntry,
        sort_by_agenda_sort_order=True):
    if len(filtered_df.index) == 0:
        return ''

    filtered_df = filtered_df.copy()
    filtered_df['TDoc'] = filtered_df.index
    filtered_df['TDoc'] = filtered_df['TDoc'].apply(lambda x: f'[{x}]({the_meeting.get_tdoc_url(x)})')
    if sort_by_agenda_sort_order:
        filtered_df = filtered_df.sort_values(by=['TDoc sort order within agenda item'])
    filtered_df = filtered_df[column_list]
    return filtered_df.to_markdown(index=False)


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


        cached_data = meeting.tdoc_data_from_excel
        self.tdocs_df = cached_data.tdocs_df
        self.wi_hyperlinks = cached_data.wi_hyperlinks

        print(f'Imported meeting Tdocs for {meeting.meeting_name}: {self.tdocs_df.columns.values}')

        self.tdocs_df = self.tdocs_df.fillna(value='')
        self.tdocs_df['Secretary Remarks'] = self.tdocs_df['Secretary Remarks'].str.replace('<br/><br/>', '. ')
        self.meeting = meeting
        self.tdoc_tags = tdoc_tags
        self.tkvar_3gpp_wifi_available = tkvar_3gpp_wifi_available

        # Process tags
        self.tdoc_tag_list_str = ['All Tags']
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
        self.release_list = ['All Releases']
        ai_items = self.tdocs_df['Release'].unique().tolist()
        ai_items.sort()
        self.release_list.extend(ai_items)
        self.ai_list = ['All AIs']
        self.ai_list.extend(self.tdocs_df['Agenda item'].unique().tolist())

        self.tdoc_status_list = ['All Status']
        self.tdoc_status_list.extend(self.tdocs_df['TDoc Status'].unique().tolist())

        self.type_list = ['All Types']
        type_items = self.tdocs_df['Type'].unique().tolist()
        type_items.sort()
        self.type_list.extend(type_items)

        self.wi_list = ['All WIs']
        wis_items_clean = self.get_wis_in_meeting()
        self.wi_list.extend(wis_items_clean)

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
        self.combo_release.set('All Releases')
        self.combo_release.bind("<<ComboboxSelected>>", self.select_rows)

        self.combo_ai = ttk.Combobox(
            self.top_frame,
            values=self.ai_list,
            state="readonly",
            width=9)
        self.combo_ai.set('All AIs')
        self.combo_ai.bind("<<ComboboxSelected>>", self.select_rows)

        self.combo_status = ttk.Combobox(
            self.top_frame,
            values=self.tdoc_status_list,
            state="readonly",
            width=10)
        self.combo_status.set('All Status')
        self.combo_status.bind("<<ComboboxSelected>>", self.select_rows)

        self.combo_type = ttk.Combobox(
            self.top_frame,
            values=self.type_list,
            state="readonly",
            width=10)
        self.combo_type.set('All Types')
        self.combo_type.bind("<<ComboboxSelected>>", self.select_rows)

        self.combo_wis = ttk.Combobox(
            self.top_frame,
            values=self.wi_list,
            state="readonly",
            width=10)
        self.combo_wis.set('All WIs')
        self.combo_wis.bind("<<ComboboxSelected>>", self.select_rows)

        self.combo_tag = ttk.Combobox(
            self.top_frame,
            values=self.tdoc_tag_list_str,
            state="readonly",
            width=10)
        self.combo_tag.set('All Tags')
        self.combo_tag.bind("<<ComboboxSelected>>", self.select_rows)

        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        self.combo_ai.pack(side=tkinter.LEFT)

        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        self.combo_status.pack(side=tkinter.LEFT)

        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        self.combo_release.pack(side=tkinter.LEFT)

        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        self.combo_type.pack(side=tkinter.LEFT)

        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        self.combo_wis.pack(side=tkinter.LEFT)

        ttk.Label(self.top_frame, text=column_separator_str).pack(side=tkinter.LEFT)
        self.combo_tag.pack(side=tkinter.LEFT)

        self.tree.bind("<Double-Button-1>", self.on_double_click)

        self.open_excel_btn = TTKHoverHelpButton(
            self.top_frame,
            image=excel_icon,
            command=lambda: open_excel_document(self.tdoc_excel_path),
            width=5,
            help_text="Open 3GU's TDoc Excel"
        )
        self.open_excel_btn.pack(side=tkinter.LEFT)

        self.open_excel_btn = TTKHoverHelpButton(
            self.top_frame,
            image=filter_icon,
            command=self.open_and_filter_excel,
            width=8,
            help_text="Filter 3GU's Excel with the same filters as the GUI"
        )
        self.open_excel_btn.pack(side=tkinter.LEFT)

        self.excel_to_markdown_btn = TTKHoverHelpButton(
            self.top_frame,
            image=markdown_icon,
            command=self.current_excel_rows_to_clipboard,
            width=9,
            help_text="Convert currently visible rows in Excel to Markdown and copy result to clipboard"
        )
        self.excel_to_markdown_btn.pack(side=tkinter.LEFT)

        self.download_btn = TTKHoverHelpButton(
            self.top_frame,
            image=cloud_download_icon,
            command=self.download_tdocs,
            width=8,
            help_text="Batch download currently shown TDocs in the table"
        )
        self.download_btn.pack(side=tkinter.LEFT)

        self.cache_btn = TTKHoverHelpButton(
            self.top_frame,
            image=folder_icon,
            command=lambda: startfile(meeting.local_folder_path),
            width=5,
            help_text="Open local folder for meeting"
        )
        self.cache_btn.pack(side=tkinter.LEFT)

        self.markdown_export_per_ai_btn = TTKHoverHelpButton(
            self.top_frame,
            image=share_markdown_icon,
            command=self.export_ais_to_markdown,
            width=6,
            help_text="Export this meeting's results per topic to Markdown for meeting report"
        )
        self.markdown_export_per_ai_btn.pack(side=tkinter.LEFT)

        # Export TDocs in table to export folder
        self.share_btn = TTKHoverHelpButton(
            self.top_frame,
            image=share_icon,
            command=self.export_tdocs_to_folder,
            help_text="Export TDocs currently visible in the Excel to export folder in specified format"
        )
        self.share_btn.pack(side=tkinter.LEFT)

        # Export format
        self.combo_export_format = ttk.Combobox(
            self.top_frame,
            values=['Original', 'PDF', 'HTML'],
            state="readonly",
            width=7)
        self.combo_export_format.set('Original')
        self.combo_export_format.pack(side=tkinter.LEFT)

        self.open_meeting_btn = TTKHoverHelpButton(
            self.top_frame,
            image=website_icon,
            command=lambda: open_url(self.meeting.meeting_url_3gu),
            width=5,
            help_text='Open meeting in 3GU'
        )
        self.open_meeting_btn.pack(side=tkinter.LEFT)

        def open_meeting_url():
            if self.meeting.meeting_is_now and tkvar_3gpp_wifi_available.get():
                ftp_folder_to_open = self.meeting.local_server_url
                print(f'Opening local WiFi FTP server page: {ftp_folder_to_open}')
            elif self.meeting.meeting_is_now:
                ftp_folder_to_open = self.meeting.sync_server_url
                print(f'Opening FTP SYNC server page: {ftp_folder_to_open}')
            else:
                ftp_folder_to_open = self.meeting.meeting_folder_url
                print(f'Opening FTP server page: {ftp_folder_to_open}')
            open_url(ftp_folder_to_open)

        self.open_meeting_ftp_btn = TTKHoverHelpButton(
            self.top_frame,
            image=ftp_icon,
            command=open_meeting_url,
            width=5,
            help_text='Open meeting folder in 3GPP FTP'
        )
        self.open_meeting_ftp_btn.pack(side=tkinter.LEFT)

        # SA2-specific buttons
        if self.meeting.working_group_enum == WorkingGroup.S2 and self.meeting.meeting_is_now:
            TTKHoverHelpButton(
                self.top_frame,
                help_text='Open Drafts folder for this meeting',
                image=note_icon,
                command=lambda: startfile(
                    open_sa2_drafts_url),
                width=10
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

    def download_tdocs(self, tdoc_list: List[str] | None = None) -> List[Any]:
        if tdoc_list is None:
            tdoc_list = self.tdocs_current_df.index.tolist()
        downloaded_files = []
        if len(tdoc_list) > 0:
            downloaded_files = batch_search_and_download_tdocs(
                tdoc_list,
                tdoc_meeting=self.meeting)

        # Re-load Tdoc list to allow for icon changes
        self.insert_rows()
        return downloaded_files

    def open_and_filter_excel(self):
        wb = open_excel_document(self.tdoc_excel_path)
        tdoc_list = self.tdocs_current_df.index.tolist()
        clear_autofilter(wb=wb)
        if len(tdoc_list) > 0:
            print(f'Filtering TDoc list for {len(tdoc_list)} TDocs shown')
            set_autofilter_values(wb=wb, value_list=tdoc_list)

    def export_tdocs_to_folder(self):
        print(f'Retrieving TDocs visible in Excel')
        wb = open_excel_document(self.tdoc_excel_path)
        tdoc_list = export_columns_to_markdown_dataframe(wb)
        # tdoc_list = self.tdocs_current_df.index.tolist()
        print(f'Exporting {len(tdoc_list)} TDocs: {tdoc_list}')

        if len(tdoc_list) < 1:
            return

        time_now = datetime.datetime.now()
        export_root_folder = utils.local_cache.get_export_folder()
        export_id = f'{time_now.year:04d}.{time_now.month:02d}.{time_now.day:02d} {time_now.hour:02d}{time_now.minute:02d}{time_now.second:02d}'
        export_folder = os.path.join(export_root_folder, export_id)

        print(f'Exporting files to {export_folder}')
        create_folder_if_needed(folder_name=export_folder, create_dir=True)

        # First we need to download the TDocs
        downloaded_files = self.download_tdocs(tdoc_list=tdoc_list)
        files_to_export = [e[1] for e in downloaded_files if e is not None and isinstance(e, tuple) and e[1] is not None]
        files_to_export = [item for sublist in files_to_export for item in sublist]
        files_to_export = [(e.tdoc_id, e.path) for e in files_to_export if isinstance(e, DownloadedTdocDocument)]

        print(f'Exporting files to {export_folder}')
        exported_files = []
        for (tdoc, file_to_export) in files_to_export:
            print(f'  {tdoc} in {file_to_export}')
            output_file = f'{tdoc}_{os.path.basename(file_to_export)}'
            output_path = os.path.join(export_folder, output_file)
            shutil.copy(file_to_export, output_path)
            exported_files.append(output_path)

        os.startfile(export_folder)

        class PdfBookmark(NamedTuple):
            page_begin: int
            page_end: int
            tdoc_id: str

        pdf_bookmarks: List[PdfBookmark] = []
        export_type = ExportType.NONE
        match self.combo_export_format.get():
            case 'HTML':
                export_type = ExportType.HTML
            case 'PDF':
                export_type = ExportType.PDF


        if export_type != ExportType.NONE:
            print('Exporting files to specified format')
            exported_pdfs = export_document(
                exported_files,
                export_format=export_type,
                do_after=ActionAfter.CLOSE_AND_DELETE_FILE)
            folder_path = Path(export_folder)

            if export_type != ExportType.PDF:
                # TDoc merge and prompt generation only for PDF format
                return

            all_exported_files :List[Path] = list([f.resolve() for f in folder_path.glob("*.pdf")])
            merger = PdfWriter()
            current_page = 0
            last_bookmark_page = 0
            last_bookmark = None
            tdoc_id = None
            old_tdoc_id = None
            for exported_file in all_exported_files:
                old_tdoc_id = tdoc_id
                tdoc_id = exported_file.name.split('_')[0]
                merger.append(exported_file.absolute())
                last_page = merger.pages[-1].page_number  # Access the last appended file's page count
                if last_bookmark is not None or last_bookmark != tdoc_id:
                    merger.add_outline_item(
                        title=tdoc_id,
                        page_number=current_page,
                        bold=True,
                        fit=PAGE_FIT
                    )

                    if last_bookmark is not None:
                        # Page number one-indexed
                        pdf_bookmarks.append(PdfBookmark(
                            page_begin = last_bookmark_page +1,
                            page_end = current_page,
                            tdoc_id=old_tdoc_id
                        ))
                    last_bookmark = tdoc_id
                    last_bookmark_page = current_page

                print(f'Merged {exported_file.name} as {tdoc_id}, bookmark on page {current_page}')
                current_page = last_page + 1

            # The last bookmark is otherwise lost
            if last_bookmark is not None:
                # Page number one-indexed
                pdf_bookmarks.append(PdfBookmark(
                    page_begin=last_bookmark_page + 1,
                    page_end=current_page,
                    tdoc_id=tdoc_id
                ))

            merge_file = os.path.join(export_folder, f"{export_id}.pdf")
            merger.write(merge_file)
            merger.close()
            os.startfile(merge_file)

            def get_text_for_prompt(e:PdfBookmark) -> str:
                el_tdoc_id = e.tdoc_id
                try:
                    el_title = self.tdocs_current_df.loc[e.tdoc_id, 'Title']
                    el_sources = self.tdocs_current_df.loc[e.tdoc_id, 'Source']

                    el_str = f'  - {el_tdoc_id}, titled {el_title}, is sourced by {el_sources} and spans pages {e.page_begin} to {e.page_end}\n'
                    return el_str
                except KeyError as e:
                    print(f'Could not generate prompt line for TDoc {el_tdoc_id}')

            with open(os.path.join(export_folder, f"{export_id}_bookmarks.txt"), 'w') as f:
                f.write('The attached PDF contains a collection of documents\n')
                lines_to_write = [get_text_for_prompt(e) for e in pdf_bookmarks]
                f.writelines(lines_to_write)

    def current_excel_rows_to_clipboard(self):
        wb = open_excel_document(self.tdoc_excel_path)
        export_columns_to_markdown(wb, MarkdownConfig.columns_for_3gu_tdoc_export)

    def get_wis_in_meeting(self, tdocs_input:DataFrame=None) -> List[str]:
        if tdocs_input is None:
            tdocs_input = self.tdocs_df

        wis_items: List[str] = tdocs_input['Related WIs'].unique().tolist()
        wis_items_clean: List[str] = []
        for wi_item in wis_items:
            if wi_item is None or wi_item == '':
                continue
            wi_item = [e.strip() for e in wi_item.split(',')]
            wis_items_clean.extend(wi_item)
        wis_items_clean = list(set(wis_items_clean))
        return sorted(wis_items_clean, key=str.lower)

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
        is_sid_new = self.tdocs_df['Type'] == 'SID new'

        is_approved_or_agreed = ((self.tdocs_df['TDoc Status'] == 'agreed') |
                                 (self.tdocs_df['TDoc Status'] == 'approved') |
                                 (self.tdocs_df['TDoc Status'] == 'endorsed'))

        pcrs_to_show = self.tdocs_df[(is_pcr & is_approved_or_agreed)]
        crs_to_show = self.tdocs_df[(is_cr & is_approved_or_agreed)]
        ls_to_show = self.tdocs_df[(is_ls_out & is_approved_or_agreed) | is_ls_in]

        company_contributions_filter = self.tdocs_df['Source'].str.contains(MarkdownConfig.company_name_regex_for_report)
        company_contributions = self.tdocs_df[company_contributions_filter]

        ls_out_to_show = self.tdocs_df[(is_ls_out & is_approved_or_agreed)]

        sid_new_to_show = self.tdocs_df[(is_sid_new & is_approved_or_agreed)]

        local_folder = self.meeting.local_export_folder_path
        meeting_name_for_export = self.meeting.meeting_name.replace('3GPP', '')

        wis_in_ai = {}
        for ai_name, group in self.tdocs_df.groupby('Agenda item'):
            wis_in_ai[ai_name] = self.get_wis_in_meeting(tdocs_input=group)
        ais_in_wi = invert_dict_defaultdict(wis_in_ai)

        print(f'Following WIs in the AIs:')
        for k, v in wis_in_ai.items():
            print(f'  - {k}: {v}')

        print(f'Following AIs in the WIs:')
        for k, v in ais_in_wi.items():
            print(f'  - {k}: {v}')

        with open(os.path.join(local_folder, f'WI-to-AI.json'), 'w') as f:
            f.write(json.dumps(wis_in_ai, indent=4))

        with open(os.path.join(local_folder, f'AI-to-WI.json'), 'w') as f:
            f.write(json.dumps(ais_in_wi, indent=4))

        print(f'Starting export per WI')

        class WiTdocs(NamedTuple):
            ls_df: DataFrame
            pcr_df: DataFrame
            cr_df: DataFrame

        wi_data: dict[str, WiTdocs] = dict()
        for wi in self.get_wis_in_meeting():
            print(f'Exporting contributions for WI {wi}')
            wi_filter = df_boolean_index_for_wi(self.tdocs_df, wi)

            full_text = ''

            wi_ls_df = self.tdocs_df[wi_filter & ((is_ls_out & is_approved_or_agreed) | is_ls_in)]
            markdown_ls = get_markdown_for_tdocs(
                wi_ls_df,
                MarkdownConfig.columns_for_3gu_tdoc_export_ls,
                self.meeting
            )

            if markdown_ls != '':
                full_text = f'{full_text}\n\nFollowing LS were received and/or answered:\n\n{markdown_ls}'

            wi_pcr_df = self.tdocs_df[wi_filter & (is_pcr & is_approved_or_agreed)]
            markdown_pcr = get_markdown_for_tdocs(
                wi_pcr_df,
                MarkdownConfig.columns_for_3gu_tdoc_export_pcr,
                self.meeting
            )

            if markdown_pcr != '':
                full_text = f'{full_text}\n\nFollowing pCRs were agreed:\n\n{markdown_pcr}'

            wi_cr_df = self.tdocs_df[wi_filter & (is_cr & is_approved_or_agreed)]
            markdown_cr = get_markdown_for_tdocs(
                wi_cr_df,
                MarkdownConfig.columns_for_3gu_tdoc_export_cr,
                self.meeting
            )

            wi_data[wi] = WiTdocs(
                ls_df=wi_ls_df,
                pcr_df=wi_pcr_df,
                cr_df=wi_cr_df
            )

            if markdown_cr != '':
                full_text = f'{full_text}\n\nFollowing CRs were agreed:\n\n{markdown_cr}'

            if full_text != '':
                full_text = f'<!--- [{meeting_name_for_export}]({self.meeting.meeting_folder_url}) --->{full_text}'
                with open(os.path.join(local_folder, f'{wi}.md'), 'w') as f:
                    f.write(full_text)

        ai_summary: dict[str,str] = dict()

        print(f'Will export per AI:')
        print(f'  - {len(pcrs_to_show)} pCRs')
        print(f'  - {len(crs_to_show)} CRs')
        print(f'  - {len(ls_to_show)} LS IN/OUT')
        print(f'  - {len(ls_out_to_show)} LS OUT')
        print(f'  - {len(company_contributions)} Company contributions matching {MarkdownConfig.company_name_regex_for_report}')
        print(f'  - {len(sid_new_to_show)} SID new')

        # Export SID new
        index_list = list(sid_new_to_show.index)
        ai_name = 'SID new'
        if len(index_list) > 0:
            print(f'{ai_name}: {len(index_list)} SID new to export')
            markdown_output = get_markdown_for_tdocs(
                sid_new_to_show,
                MarkdownConfig.columns_for_3gu_tdoc_export,
                self.meeting
            )

            ai_summary[ai_name] = f'Following new SIDs were agreed:\n\n{markdown_output}'
        else:
            print(f'{ai_name}: {len(index_list)} SID new')

        # Export LS OUT
        index_list = list(ls_out_to_show.index)
        ai_name = 'LS OUT'
        if len(index_list) > 0:
            print(f'{ai_name}: {len(index_list)} LS IN/OUT to export')
            markdown_output = get_markdown_for_tdocs(
                ls_out_to_show,
                MarkdownConfig.columns_for_3gu_tdoc_export_ls_out,
                self.meeting
            )

            ai_summary[ai_name] = f'Following LS OUT were sent:\n\n{markdown_output}'
        else:
            print(f'{ai_name}: {len(index_list)} LS IN/OUT')

        # Export Company Contributions
        index_list = list(company_contributions.index)
        ai_name = 'Company'
        if len(index_list) > 0:
            print(f'{len(index_list)} Company contributions matching '
                  f'{MarkdownConfig.company_name_regex_for_report} to export')
            markdown_output = get_markdown_for_tdocs(
                company_contributions,
                MarkdownConfig.columns_for_3gu_tdoc_export_contributor,
                self.meeting
            )

            ai_summary[ai_name] = f'Following Company Contributions:\n\n{markdown_output}'
        else:
            print(f'{len(index_list)} Company contributions matching {MarkdownConfig.company_name_regex_for_report}')

        # Export LSs per Agenda Item
        for ai_name, group in ls_to_show.groupby('Agenda item'):
            index_list = list(group.index)

            if len(index_list) > 0:
                print(f'{ai_name}: {len(index_list)} LS IN/OUT to export')
                markdown_output = get_markdown_for_tdocs(
                    group,
                    MarkdownConfig.columns_for_3gu_tdoc_export_ls,
                    self.meeting
                )

                ai_summary[ai_name] = f'Following LS were received and/or answered:\n\n{markdown_output}'
            else:
                print(f'{ai_name}: {len(index_list)} LS IN/OUT')

        # Export pCRs per Agenda Item
        for ai_name, group in pcrs_to_show.groupby('Agenda item'):
            index_list = list(group.index)

            if len(index_list) > 0:
                print(f'{ai_name}: {len(index_list)} pCRs to export')
                markdown_output = get_markdown_for_tdocs(
                    group,
                    MarkdownConfig.columns_for_3gu_tdoc_export_pcr,
                    self.meeting
                )

                summary_text = f'Following pCRs were agreed:\n\n{markdown_output}'
                if ai_name in ai_summary:
                    ai_summary[ai_name] = f'{ai_summary[ai_name]}\n\n{summary_text}'
                else:
                    ai_summary[ai_name] = f'{summary_text}'

            else:
                print(f'{ai_name}: {len(index_list)} pCRs')

        # Export CRs per Agenda Item
        for ai_name, group in crs_to_show.groupby('Agenda item'):
            index_list = list(group.index)

            if len(index_list) > 0:
                print(f'{ai_name}: {len(index_list)} CRs to export')
                markdown_output = get_markdown_for_tdocs(
                    group,
                    MarkdownConfig.columns_for_3gu_tdoc_export_cr,
                    self.meeting
                )

                summary_text = f'Following CRs were agreed:\n\n{markdown_output}'
                if ai_name in ai_summary:
                    ai_summary[ai_name] = f'{ai_summary[ai_name]}\n\n{summary_text}'
                else:
                    ai_summary[ai_name] = f'{summary_text}'

            else:
                print(f'{ai_name}: {len(index_list)} CRs')

        for ai_name, summary_text in ai_summary.items():
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
                    opened_docs_folder, metadata = server.tdoc_search.search_download_and_open_tdoc(
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
                    meeting=self.meeting,
                    wi_hyperlinks=self.wi_hyperlinks
                )

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
        if not ai_filter.startswith('All'):
            print(f'Filtering by AI: "{ai_filter}"')
            filtered_df = filtered_df[filtered_df["Agenda item"] == ai_filter]

        status_filter = self.combo_status.get()
        if not status_filter.startswith('All'):
            print(f'Filtering by TDoc status: "{status_filter}"')
            filtered_df = filtered_df[filtered_df["TDoc Status"] == status_filter]

        type_filter = self.combo_type.get()
        if not type_filter.startswith('All'):
            print(f'Filtering by Type: "{type_filter}"')
            filtered_df = filtered_df[filtered_df["Type"] == type_filter]

        rel_filter = self.combo_release.get()
        if not rel_filter.startswith('All'):
            print(f'Filtering by Release: "{rel_filter}"')
            filtered_df = filtered_df[filtered_df["Release"] == rel_filter]

        wi_filter = self.combo_wis.get()
        if not wi_filter.startswith('All'):
            print(f'Filtering by WI: "{wi_filter}"')
            filtered_df = filtered_df[df_boolean_index_for_wi(filtered_df, wi_filter)            ]

        tag_filter = self.combo_tag.get()
        if not tag_filter.startswith('All'):
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
            wi_hyperlinks: dict[str, str],
            root_widget: tkinter.Tk | None = None,
    ):
        self.tdoc_id = tdoc_str
        self.tdoc_row = tdoc_row
        self.tkvar_3gpp_wifi_available = tkvar_3gpp_wifi_available
        self.meeting = meeting
        self.wi_hyperlinks = wi_hyperlinks

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
                    opened_docs_folder, metadata = server.tdoc_search.search_download_and_open_tdoc(
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
            case 'Related WIs':
                related_wis = [wi.strip() for wi in actual_value.split(',') if wi is not None and wi!='']
                for related_wi in related_wis:
                    try:
                        the_url = self.wi_hyperlinks[related_wi]
                        open_url(the_url)
                        print(f'Opened URL WI {related_wi}: {the_url}')
                    except KeyError:
                        pass


