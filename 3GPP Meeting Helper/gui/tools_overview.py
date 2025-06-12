import datetime
import os
import os.path
import threading
import tkinter
import traceback
from tkinter import ttk
from typing import List, Callable

import pandas
import pythoncom

import application.excel
import application.meeting_helper
import application.word
import gui.common.utils
import gui.main_gui
import gui.meetings_table
import gui.specs_table
import gui.tdocs_table
import gui.work_items_table
import parsing.excel as excel_parser
import parsing.html.common
import parsing.html.tdocs_by_agenda
import parsing.outlook
import parsing.word.pywin32
import parsing.word.pywin32 as word_parser
import server.agenda
import server.chairnotes
import server.common.server_utils
import server.tdoc
import utils.local_cache
from gui.common.tkinter_widget import TkWidget
from parsing.html.chairnotes import chairnotes_file_to_dataframe
from parsing.html.revisions import extract_tdoc_revisions_from_html
from utils.threading import do_something_on_thread


class ToolsDialog(TkWidget):
    export_text = 'Export TDocs by Agenda of meeting to Excel + Add comments found in Agenda folder (saved in Agenda folder)'
    export_year_text = 'Export Tdocs for given year (saved in current meeting Agenda folder)'
    outlook_text_1 = "Order & summarize Outlook emails for current email approval/e-meeting"
    outlook_text_2 = "Order emails for current meeting's email approval/e-meeting"
    outlook_attachment_text = "Download SA2 email attachments (excluded are email approval emails)"
    word_report_text = 'Word report for AIs (empty=all)'
    bulk_tdoc_open_text = "Cache TDocs"
    bulk_ai_open_text = "Cache AIs (empty=all)"

    def __init__(
            self,
            parent: tkinter.Tk,
            favicon: str,
            selected_meeting_fn: Callable[[], str]
    ):
        super().__init__(
            parent,
            None,
            "Extended tools",
            favicon
        )

        self.top_level_widget: tkinter.Tk = parent
        self.selected_meeting_fn = selected_meeting_fn

        columnspan = 4

        # Set the window to the front and wait until it is closed
        # https://stackoverflow.com/questions/1892339/how-to-make-a-tkinter-window-jump-to-the-front
        # top.attributes("-topmost", True)

        # Row 1: Export TDocs by agenda to Excel
        self.export_button = ttk.Button(self.tk_top, text=ToolsDialog.export_text,
                                            command=self.export_tdocs_by_agenda_to_excel)
        self.export_button.grid(row=1, column=0, columnspan=columnspan, sticky="EW")

        self.tkvar_tdoc = tkinter.StringVar(self.tk_top)
        self.tkvar_original_tdocs = tkinter.StringVar(self.tk_top)
        self.tkvar_final_tdocs = tkinter.StringVar(self.tk_top)

        self.tkvar_original_tdocs.set('Press button to analyze')
        self.tkvar_final_tdocs.set('Press button to analyze')

        # Row 2: Outlook tools (email approval emails)
        self.outlook_button_text = tkinter.StringVar(self.tk_top)
        self.outlook_button_text.set(ToolsDialog.outlook_text_2)
        self.outlook_generate_summary = tkinter.BooleanVar(self.tk_top)
        self.outlook_generate_summary.set(False)

        def change_outlook_button_label():
            if self.outlook_generate_summary.get():
                self.outlook_button_text.set(ToolsDialog.outlook_text_1)
            else:
                self.outlook_button_text.set(ToolsDialog.outlook_text_2)

        self.email_attachments_generate_summary_checkbox = ttk.Checkbutton(self.tk_top,
                                                                               text="Cache emails & Excel summary",
                                                                               variable=self.outlook_generate_summary,
                                                                               command=change_outlook_button_label)
        self.email_attachments_generate_summary_checkbox.grid(row=2, column=0, sticky="EW")
        self.email_approval_button = ttk.Button(self.tk_top, textvariable=self.outlook_button_text,
                                                    command=self.outlook_email_approval)
        self.email_approval_button.grid(row=2, column=1, columnspan=2, sticky="EW")
        self.download_chairnotes = ttk.Button(self.tk_top, text="Process Chairman's Notes (be patient)",
                                                  command=self.process_chairnotes)
        self.download_chairnotes.grid(row=2, column=3, columnspan=1, sticky="EW")

        # Row 3: Outlook tools (download email attachments)
        self.email_attachments_button = ttk.Button(self.tk_top, text=ToolsDialog.outlook_attachment_text,
                                                       command=self.outlook_email_attachments)
        self.email_attachments_button.grid(row=3, column=0, columnspan=columnspan, sticky="EW")

        # Row 5: Export tdocs from a given year
        self.tkvar_year = tkinter.StringVar(self.tk_top)
        self.year_entry = tkinter.Entry(self.tk_top, textvariable=self.tkvar_year, width=25, font='TkDefaultFont')
        self.year_entry.insert(0, str(datetime.datetime.now().year))
        self.year_entry.grid(row=5, column=0, padx=10, pady=10)
        self.year_entry.config(state='normal')
        self.tdoc_report_button = ttk.Button(self.tk_top, text=ToolsDialog.export_year_text,
                                                 command=self.export_year_tdocs_by_agenda_to_excel)
        self.tdoc_report_button.grid(row=5, column=1, columnspan=3, sticky="EW")

        # Row 6A: Generate Word report
        self.tdoc_word_report_button = ttk.Button(
            self.tk_top,
            text=ToolsDialog.word_report_text,
            command=self.generate_word_report)
        self.tdoc_word_report_button.grid(row=6, column=0, columnspan=1, sticky="EW")

        self.tkvar_ai_list_word_report = tkinter.StringVar(self.tk_top)
        self.ai_list_entry_word_report = tkinter.Entry(
            self.tk_top,
            textvariable=self.tkvar_ai_list_word_report,
            width=30,
            font='TkDefaultFont')
        self.ai_list_entry_word_report.insert(0, '')
        self.ai_list_entry_word_report.grid(row=6, column=1, columnspan=1, padx=10, pady=10, sticky="EW")

        # Row 6B: Bulk AI cache
        self.ai_bulk_open_button = ttk.Button(
            self.tk_top,
            text=ToolsDialog.bulk_ai_open_text,
            command=self.bulk_cache_ais)
        self.ai_bulk_open_button.grid(row=6, column=2, columnspan=1, sticky="EW")

        self.tkvar_ai_list = tkinter.StringVar(self.tk_top)
        self.ai_list_entry = tkinter.Entry(
            self.tk_top,
            textvariable=self.tkvar_ai_list,
            width=30,
            font='TkDefaultFont')
        self.ai_list_entry.insert(0, '')
        self.ai_list_entry.grid(row=6, column=3, columnspan=1, padx=10, pady=10, sticky="EW")

        # Row 8: Replace Author names in active document
        self.replace_author_names_button = ttk.Button(
            self.tk_top,
            text="Replace Active Doc's review author",
            command=self.replace_document_revisions_author)
        self.replace_author_names_button.grid(row=8, column=0, columnspan=1, sticky="EW")

        self.original_author_name = tkinter.StringVar(self.tk_top)
        self.original_author_name_entry = tkinter.Entry(
            self.tk_top,
            textvariable=self.original_author_name,
            width=30,
            font='TkDefaultFont')
        self.original_author_name_entry.insert(0, '')
        self.original_author_name_entry.grid(row=8, column=1, columnspan=1, padx=10, pady=10, sticky="EW")

        self.final_author_name = tkinter.StringVar(self.tk_top)
        self.final_author_name_entry = tkinter.Entry(
            self.tk_top,
            textvariable=self.final_author_name,
            width=30,
            font='TkDefaultFont')
        self.final_author_name_entry.insert(0, '')
        self.final_author_name_entry.grid(row=8, column=2, columnspan=1, padx=10, pady=10, sticky="EW")

        # Info after analyzing TDoc
        current_row = 9

        def set_original_tdocs(*args):
            self.original_tdocs_textbox.delete('1.0', tkinter.END)
            self.original_tdocs_textbox.insert(tkinter.END, self.tkvar_original_tdocs.get())

        self.tkvar_original_tdocs.trace('w', set_original_tdocs)

        def set_final_tdocs(*args):
            self.final_tdocs_textbox.delete('1.0', tkinter.END)
            self.final_tdocs_textbox.insert(tkinter.END, self.tkvar_final_tdocs.get())

        self.tkvar_final_tdocs.trace('w', set_final_tdocs)

        # Configure column row widths
        self.tk_top.grid_columnconfigure(0, weight=1)
        self.tk_top.grid_columnconfigure(1, weight=1)
        self.tk_top.grid_columnconfigure(2, weight=1)
        self.tk_top.grid_columnconfigure(3, weight=1)

    def generate_word_report(self):
        current_tdocs_by_agenda = gui.main.open_tdocs_by_agenda(open_this_file=False)
        if current_tdocs_by_agenda is None:
            print('Could not generate report. No current TDocsByAgenda found')
            return
        selected_meeting = self.selected_meeting_fn()
        server_folder = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(
            selected_meeting)
        ais_for_report = self.tkvar_ai_list_word_report.get()
        if (ais_for_report is not None) and (ais_for_report != ''):
            try:
                ais_to_output = [ai.strip() for ai in ais_for_report.split(',')]
            except:
                print('Could not parse AI list for Word report: {0}'.format(ais_for_report))
                ais_to_output = []
        else:
            ais_to_output = []

        if len(ais_to_output) == 0:
            ais_to_output = None

        meeting_folder = utils.local_cache.get_local_agenda_folder(server_folder)

        df_tdocs = current_tdocs_by_agenda.tdocs
        email_approval_tdocs = df_tdocs[(df_tdocs['Result'] == 'For e-mail approval')]
        n_email_approval = len(email_approval_tdocs)
        print('TDocsByAgenda: {0} TDocs marked as "For e-mail approval"'.format(n_email_approval))

        # Force Agenda cache for the section titles
        server.agenda.get_last_agenda(server_folder)

        doc = application.word.open_word_document()
        word_parser.insert_doc_data_to_doc_by_wi(
            df_tdocs,
            doc,
            server_folder,
            source=application.meeting_helper.company_report_name,
            ais_to_output=ais_to_output,
            save_to_folder=meeting_folder)
        print('Finished exporting Word report')

    # ToDo: change output_meeting so that it can be a string!!!
    def export_tdocs_by_agenda_to_excel(
            self,
            selected_meeting=None,
            output_meeting=None,
            filename='TdocsByAgenda',
            close_file=False,
            use_thread=True,
            current_dt_str=None,
            process_comments=True,
            destination_folder=None,
            add_pivot_table=True):
        try:
            if selected_meeting is None:
                selected_meeting = self.selected_meeting_fn()
            if output_meeting is None:
                output_meeting = self.selected_meeting_fn()

            input_meeting_folder = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(
                selected_meeting)
            output_meeting_folder = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(
                output_meeting)
            inbox_active = False

            input_local_agenda_file = server.agenda.download_agenda_file(selected_meeting, inbox_active)
            if input_local_agenda_file is None:
                print('Could not find agenda file for {0}'.format(selected_meeting))
                return

            if destination_folder is None:
                (head, tail) = os.path.split(
                    utils.local_cache.get_tdocs_by_agenda_filename(output_meeting_folder))
            else:
                head = destination_folder
            if current_dt_str is None:
                current_dt_str = application.meeting_helper.get_now_time_str()
            excel_export = os.path.join(head, '{0} {1}.xlsx'.format(current_dt_str, filename))
            print('Exporting ' + input_local_agenda_file + ' to ' + excel_export)

            # Disable buttons
            self.export_button.config(text='Exporting... DO NOT interrupt Excel until COMPLETELY finished!',
                                      state='disabled')
            self.year_entry.config(state='readonly')
            self.tdoc_report_button.config(text='Exporting... DO NOT interrupt Excel until COMPLETELY finished!',
                                           state='disabled')

            if use_thread:
                t = threading.Thread(target=lambda: self.export_and_open_excel(
                    input_local_agenda_file,
                    excel_export,
                    input_meeting_folder,
                    selected_meeting,
                    close_file,
                    process_comments=process_comments,
                    add_pivot_summary=add_pivot_table))
                t.start()
            else:
                self.export_and_open_excel(
                    input_local_agenda_file,
                    excel_export,
                    input_meeting_folder,
                    selected_meeting,
                    close_file,
                    process_comments=process_comments,
                    add_pivot_summary=add_pivot_table)
        except:
            print('Could not export TDocs by agenda data')
            traceback.print_exc()

    def export_and_open_excel(
            self,
            local_agenda_file,
            excel_export,
            meeting_folder,
            sheet_name,
            close_file=False,
            process_comments=True,
            add_pivot_summary=True):
        try:
            tdocs_by_agenda = parsing.html.tdocs_by_agenda.TdocsByAgendaData(
                gui.main_gui.get_tdocs_by_agenda_file_or_url(local_agenda_file),
                meeting_server_folder=meeting_folder
            )

            # Do not export to Excel the last columns (just a lot of True/False columns for each vendor)
            tdocs_df = tdocs_by_agenda.tdocs.iloc[:, 0:19]

            tdocs = tdocs_df.index.tolist()
            server_urls = [(tdoc, server.tdoc.get_remote_filename_for_tdoc(meeting_folder, tdoc, use_private_server=False)) for tdoc in
                           tdocs]
            tdocs_df.loc[:, parsing.excel.session_comments_column] = ''

            # Get TDoc comments from the comments files
            agenda_folder = os.path.dirname(os.path.abspath(local_agenda_file))
            if process_comments:
                parsed_comments = parsing.excel.get_comments_from_dir_format(agenda_folder, merge_comments=True)
            else:
                parsed_comments = None
            fg_color = {}
            text_color = {}
            if parsed_comments is not None:
                comments = []

                # Generate meta-comments for revision of and merge of for easier review
                for tdoc_idx in tdocs:
                    row = tdocs_df.loc[tdoc_idx, :]
                    merge_of = row['Merge of']
                    revision_of = row['Revision of']
                    session_comments = row[parsing.excel.session_comments_column]

                    tdoc_parent_list = []
                    if revision_of != '':
                        tdoc_parent_list.append(revision_of)
                    if merge_of != '':
                        merge_of_parent_list = [e.strip() for e in merge_of.split(',') if e.strip() != '']
                        tdoc_parent_list.extend(merge_of_parent_list)

                    comments_for_this_tdoc = [(parent_tdoc, parsed_comments[parent_tdoc]) for parent_tdoc in
                                              tdoc_parent_list if
                                              parent_tdoc in parsed_comments and len(parsed_comments[parent_tdoc]) > 0]
                    # Only if there are no comments and there are revisions to do
                    if len(comments_for_this_tdoc) > 0:
                        if len(comments_for_this_tdoc) > 0:
                            list_of_sublists = [tup[1] for tup in comments_for_this_tdoc]
                            fg_colors = [e[2] for sublist in list_of_sublists for e in sublist]
                            text_colors = [e[3] for sublist in list_of_sublists for e in sublist]
                            fg_color = parsing.excel.get_reddest_color(fg_colors)
                            text_color = parsing.excel.get_reddest_color(text_colors)
                        else:
                            fg_color = None
                            text_color = None
                        merged_texts = []
                        for comment in comments_for_this_tdoc:
                            comment_tdoc = comment[0]
                            comment_list = comment[1]
                            comment_list = [parsing.excel.get_comment_full_text(e[0], '{0}'.format(e[1])) for e in
                                            comment_list]
                            merged_comment_list = '\n'.join(comment_list)
                            merged_comment = parsing.excel.get_comment_full_text(comment_tdoc, '{{\n{0}\n}}'.format(
                                merged_comment_list))
                            merged_texts.append(merged_comment)

                        if len(merged_texts) > 0:
                            merged_comment_text = '\n'.join([e for e in merged_texts])
                            # 'None' generatese no tag for the comment
                            parent_comment = (None, merged_comment_text, fg_color, text_color)
                        else:
                            parent_comment = None
                    else:
                        parent_comment = None
                    # Store comments
                    if tdoc_idx not in parsed_comments:
                        parsed_comments[tdoc_idx] = []
                    if parent_comment is not None:
                        parsed_comments[tdoc_idx].append(parent_comment)
                        comments.append(parent_comment)
                # Apply comments
                for idx, comment_list in parsed_comments.items():
                    comment_list_txt = [parsing.excel.get_comment_full_text(comment[0], comment[1], ) for comment in
                                        comment_list]
                    full_comment = '\n'.join(comment_list_txt)
                    try:
                        tdocs_df.at[idx, parsing.excel.session_comments_column] = full_comment
                    except Exception as e:
                        print(f'Did not find TDoc entry for comment {idx}. Skipping: {e}')
                fg_color, text_color = parsing.excel.get_colors_from_comments(parsed_comments)

            # Need Pandas 0.24 for this .See https://stackoverflow.com/questions/42589835/adding-a-pandas-dataframe-to-existing-excel-file
            # or just use https://github.com/pandas-dev/pandas/issues/3441#issuecomment-24898286
            # Needs Pandas >= 0.24.0
            # Note that xlsxwriter does NOT support append mode
            if os.path.isfile(excel_export):
                write_mode = 'a'
            else:
                write_mode = 'w'
            with pandas.ExcelWriter(excel_export, engine='openpyxl', mode=write_mode) as writer:
                tdocs_df.to_excel(writer, sheet_name=sheet_name)

            parsing.excel.apply_comments_coloring_and_hyperlinks(excel_export, fg_color, text_color, server_urls)

            # Need to reinitialize COM on each thread
            # https://stackoverflow.com/questions/26745617/win32com-client-dispatch-cherrypy-coinitialize-has-not-been-called
            pythoncom.CoInitialize()
            wb = application.excel.open_excel_document(excel_export, sheet_name=sheet_name)

            application.excel.set_first_row_as_filter(wb)
            excel_parser.adjust_tdocs_by_agenda_column_width(wb)
            excel_parser.set_tdoc_colors(wb, server_urls)
            application.excel.hide_columns(wb, ['H', 'K:S'])
            application.excel.vertically_center_all_text(wb)
            if add_pivot_summary:
                excel_parser.generate_pivot_chart_from_tdocs_by_agenda(wb)
            else:
                print('Skipping generation of pivot table')
            application.excel.save_wb(wb)
            if close_file:
                application.excel.close_wb(wb)

            print('Non-parsed vendors: {0}'.format('\n'.join(tdocs_by_agenda.others_cosigners)))
        except:
            print('Could not export TDoc list + comments to Excel')
            traceback.print_exc()
        finally:
            self.export_button.config(text=ToolsDialog.export_text, state='normal')
            self.tdoc_report_button.config(text=ToolsDialog.export_year_text, state='normal')
            self.year_entry.config(state='normal')

    def outlook_email_approval(self):
        current_text = self.outlook_button_text.get()
        self.outlook_button_text.set('Processing... DO NOT interrupt Outlook until COMPLETELY finished!')
        self.email_approval_button.config(state='disabled')
        self.email_attachments_generate_summary_checkbox.config(state='disabled')
        t = threading.Thread(target=lambda: self.on_outlook_email_approval(current_text))
        t.start()

    def on_outlook_email_approval(self, current_text):
        try:
            # Need to reinitialize COM on each thread
            # https://stackoverflow.com/questions/26745617/win32com-client-dispatch-cherrypy-coinitialize-has-not-been-called
            pythoncom.CoInitialize()
            selected_meeting = self.selected_meeting_fn()
            generate_summary = self.outlook_generate_summary.get()
            parsing.outlook.process_email_approval(selected_meeting, generate_summary)
        finally:
            self.outlook_button_text.set(current_text)
            self.email_approval_button.config(state='normal')
            self.email_attachments_generate_summary_checkbox.config(state='normal')

    def outlook_email_attachments(self):
        self.email_attachments_button.config(text='Processing... DO NOT interrupt Outlook until COMPLETELY finished!',
                                             state='disabled')
        t = threading.Thread(target=lambda: self.on_outlook_email_attachments())
        t.start()

    def on_outlook_email_attachments(self):
        try:
            # Need to reinitialize COM on each thread
            # https://stackoverflow.com/questions/26745617/win32com-client-dispatch-cherrypy-coinitialize-has-not-been-called
            pythoncom.CoInitialize()
            parsing.outlook.process_email_attachments()
        finally:
            self.email_attachments_button.config(text=ToolsDialog.outlook_attachment_text, state='normal')

    def export_year_tdocs_by_agenda_to_excel(self):
        try:
            year_to_search = int(self.tkvar_year.get())
            meetings_to_check = application.meeting_helper.sa2_meeting_data.get_meetings_for_given_year(year_to_search)
            output_meeting = self.selected_meeting_fn()

            t = threading.Thread(
                target=lambda: self.on_export_year_tdocs_by_agenda_to_excel(meetings_to_check, output_meeting,
                                                                            'Report {0}'.format(year_to_search)))
            t.start()
        except:
            print('Could not generate yearly report')
            traceback.print_exc()

    def on_export_year_tdocs_by_agenda_to_excel(self, meetings_to_check, output_meeting, filename):
        last_folder = meetings_to_check.meeting_names[-1]
        current_dt_str = application.meeting_helper.get_now_time_str()

        for meeting_to_check in meetings_to_check.meeting_names:
            print(meeting_to_check)
            close_file = meeting_to_check != last_folder
            self.export_tdocs_by_agenda_to_excel(
                selected_meeting=meeting_to_check,
                output_meeting=output_meeting,
                filename=filename,
                close_file=close_file,
                use_thread=False,
                current_dt_str=current_dt_str,
                process_comments=False,
                destination_folder=utils.local_cache.get_cache_folder(),
                add_pivot_table=False)

    def bulk_cache_ais(self):
        try:
            tdocs = application.meeting_helper.current_tdocs_by_agenda.tdocs
            ais = self.tkvar_ai_list.get()
            if ais is None:
                return
            if ais == '':
                ais = None
            else:
                ais = [ai.strip() for ai in ais.replace(',', '').split(' ') if (ai is not None) and (ai != '')]
            tdocs_to_cache: List[str] = []
            if ais is None:
                # More efficient download by just checking the Docs folder in the server's meeting folder
                # Add revisions and drafts folder only for the "no AIs case" (at least for now)
                meeting_server_folder = application.meeting_helper.current_tdocs_by_agenda.meeting_server_folder

                docs_file = server.tdoc.download_docs_file(meeting_server_folder)
                revisions_file, revisions_folder_url = server.tdoc.download_revisions_file(meeting_server_folder)
                drafts_file = server.tdoc.download_drafts_file(meeting_server_folder)

                docs_list = extract_tdoc_revisions_from_html(
                    docs_file,
                    is_draft=False,
                    is_path=True,
                    ignore_revision=True)
                revisions_list = extract_tdoc_revisions_from_html(revisions_file, is_draft=False, is_path=True)
                drafts_list = extract_tdoc_revisions_from_html(drafts_file, is_draft=True, is_path=True)

                docs_list_full = [e.tdoc for e in docs_list]
                revisions_list_full = ['{0}r{1}'.format(e.tdoc, e.revision) for e in revisions_list]
                drafts_list_full = ['{0}r{1}'.format(e.tdoc, e.revision) for e in drafts_list]

                tdocs_to_cache.extend(docs_list_full)
                tdocs_to_cache.extend(revisions_list_full)
                drafts_list_full.extend(drafts_list_full)
            else:
                for ai in ais:
                    ai_tdocs = tdocs.index[tdocs['AI'] == ai].tolist()
                    if len(ai_tdocs) > 0:
                        tdocs_to_cache.extend(ai_tdocs)
                    print('{0} items in AI {1}, total items: {2}'.format(len(ai_tdocs), ai, len(tdocs.index)))

            # Temporarily disable
            download_from_inbox = gui.main_gui.tkvar_3gpp_wifi_available.get()
            meeting_folder_name = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(
                self.selected_meeting_fn())

            do_something_on_thread(
                task=lambda: server.tdoc.cache_tdocs(
                    tdoc_list=tdocs_to_cache,
                    download_from_private_server=download_from_inbox,
                    meeting_folder_name=meeting_folder_name),
                before_starting=lambda: self.ai_bulk_open_button.config(state=tkinter.DISABLED),
                after_task=lambda: self.ai_bulk_open_button.config(state=tkinter.NORMAL),
                on_error_log='General error performing bulk AI caching'
            )

        except Exception as e:
            print(f'General error performing bulk AI caching: {e}')
            traceback.print_exc()

    def replace_document_revisions_author(self):
        original_author_name_to_replace = self.original_author_name.get()
        final_author_name = self.final_author_name.get()
        print("Will change Author name '{0}' to '{1}' for changes in active Word document".format(
            original_author_name_to_replace,
            final_author_name))
        application.word.get_reviews_for_active_document(
            search_author=original_author_name_to_replace,
            replace_author=final_author_name)

    def process_chairnotes(self):
        selected_meeting = self.selected_meeting_fn()
        meeting_folder = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(
            selected_meeting)
        local_file = server.chairnotes.download_chairnotes_file(meeting_folder)
        latest_chairnotes_df = chairnotes_file_to_dataframe(local_file)
        print(latest_chairnotes_df)
        # ToDo
        #  - Download most recent files based on DataFrame
        #  - Covert each file to docx so that the docx library can parse it (pywin32 is very slow in comparison)
        #  - Parse each file
        #  - Generate output
        return
