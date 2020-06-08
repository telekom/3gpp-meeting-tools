import application
import tkinter
import gui.main
import os
import os.path
import server
import traceback
import parsing.html as html_parser
import parsing.excel as excel_parser
import parsing.outlook
import parsing.word as word_parser
import threading
import pythoncom
import tdoc
import datetime
import pandas

class ToolsDialog:

    export_text      = 'Export TDocs by Agenda of meeting to Excel + Add comments found in Agenda folder (saved in Agenda folder)'
    export_year_text = 'Export Tdocs for given year (saved in current meeting Agenda folder)'
    outlook_text_1   = "Order and summarize Outlook emails for current meeting's email approval/e-meeting"
    outlook_text_2   = "Order emails for current meeting's email approval/e-meeting"
    outlook_attachment_text = "Download SA2 email attachments (excluded are email approval emails)"
    word_report_text = 'Word report for AIs (empty=all)'
    bulk_tdoc_open_text = "Cache TDocs"
    bulk_ai_open_text = "Cache AIs (empty=all)"

    def __init__(self, parent, favicon):
        top = self.top = tkinter.Toplevel(parent)
        top.title("Extended tools")
        top.iconbitmap(favicon)

        columnspan = 4

        # Set the window to the front and wait until it is closed
        # https://stackoverflow.com/questions/1892339/how-to-make-a-tkinter-window-jump-to-the-front
        # top.attributes("-topmost", True)
        tkinter.Button(top, text="Open meeting folder for selected meeting", command=self.open_local_meeting_folder).grid(row=0, column=0, columnspan=int(columnspan/2), sticky="EW")
        tkinter.Button(top, text="Open server meeting folder", command=self.open_server_meeting_folder).grid(row=0, column=2, columnspan=int(columnspan/2), sticky="EW")

        # Row 1: Export TDocs by agenda to Excel
        self.export_button = tkinter.Button(top, text=ToolsDialog.export_text, command=self.export_tdocs_by_agenda_to_excel)
        self.export_button.grid(row=1, column=0, columnspan=columnspan, sticky="EW")

        # Row 4: Analyze TDoc
        self.analyze_tdoc_button = tkinter.Button(top, text='Analyze last opened TDoc from last opened TDocs by agenda', command=self.analyze_tdoc)
        self.analyze_tdoc_button.grid(row=4, column=0, columnspan=columnspan, sticky="EW")

        self.tkvar_tdoc           = tkinter.StringVar(top)
        self.tkvar_original_tdocs = tkinter.StringVar(top)
        self.tkvar_final_tdocs    = tkinter.StringVar(top)

        self.tkvar_original_tdocs.set('Press button to analyze')
        self.tkvar_final_tdocs.set('Press button to analyze')

        # Row 2: Outlook tools (email approval emails)
        self.outlook_button_text = tkinter.StringVar(top)
        self.outlook_button_text.set(ToolsDialog.outlook_text_2)
        self.outlook_generate_summary = tkinter.BooleanVar(top)
        self.outlook_generate_summary.set(False)

        def change_outlook_button_label():
            if self.outlook_generate_summary.get():
                self.outlook_button_text.set(ToolsDialog.outlook_text_1)
            else:
                self.outlook_button_text.set(ToolsDialog.outlook_text_2)
        self.email_attachments_generate_summary_checkbox = tkinter.Checkbutton(top, text="Cache emails & Excel summary", variable=self.outlook_generate_summary, command=change_outlook_button_label)
        self.email_attachments_generate_summary_checkbox.grid(row=2, column=0, sticky="EW")
        self.email_approval_button = tkinter.Button(top, textvariable=self.outlook_button_text, command=self.outlook_email_approval)
        self.email_approval_button.grid(row=2, column=1, columnspan=3, sticky="EW")

        # Row 3: Outlook tools (download email attachments)
        self.email_attachments_button = tkinter.Button(top, text=ToolsDialog.outlook_attachment_text, command=self.outlook_email_attachments)
        self.email_attachments_button.grid(row=3, column=0, columnspan=columnspan, sticky="EW")

        # Row 5: Export tdocs from a given year
        self.tkvar_year = tkinter.StringVar(top)
        self.year_entry = tkinter.Entry(top, textvariable=self.tkvar_year, width=25, font='TkDefaultFont')
        self.year_entry.insert(0, str(datetime.datetime.now().year))
        self.year_entry.grid(row=5, column=0, padx=10, pady=10)
        self.year_entry.config(state='normal')
        self.tdoc_report_button = tkinter.Button(top, text=ToolsDialog.export_year_text, command=self.export_year_tdocs_by_agenda_to_excel)
        self.tdoc_report_button.grid(row=5, column=1, columnspan=3, sticky="EW")

        # Row 6: Generate Word report
        self.tkvar_ai_list_word_report = tkinter.StringVar(top)
        self.ai_list_entry_word_report = tkinter.Entry(top, textvariable=self.tkvar_ai_list_word_report, width=90, font='TkDefaultFont')
        self.ai_list_entry_word_report.insert(0, '')
        self.ai_list_entry_word_report.grid(row=6, column=1, columnspan=3, padx=10, pady=10)
        self.tdoc_word_report_button = tkinter.Button(top, text=ToolsDialog.word_report_text, command=self.generate_word_report)
        self.tdoc_word_report_button.grid(row=6, column=0, columnspan=1, sticky="EW")

        # Row 7: Bulk AI cache
        self.tkvar_ai_list = tkinter.StringVar(top)
        self.ai_list_entry = tkinter.Entry(top, textvariable=self.tkvar_ai_list, width=90, font='TkDefaultFont')
        self.ai_list_entry.insert(0, '')
        self.ai_list_entry.grid(row=7, column=1, columnspan=3, padx=10, pady=10)

        self.ai_bulk_open_button = tkinter.Button(top, text=ToolsDialog.bulk_ai_open_text, command=self.bulk_cache_ais)
        self.ai_bulk_open_button.grid(row=7, column=0, columnspan=1, sticky="EW")

        # Info after analyzing TDoc
        current_row = 8
        tkinter.ttk.Separator(top,orient=tkinter.HORIZONTAL).grid(row=current_row, columnspan=3, sticky=(tkinter.W,tkinter.E))
        
        current_row += 1
        tkinter.Label(top, text='Last opened TDoc:').grid(row=current_row, column=0)
        tkinter.Label(top, textvariable=self.tkvar_tdoc).grid(row=current_row, column=1)

        current_row += 1
        tkinter.Label(top, text='Original TDocs:').grid(row=current_row, column=0)
        self.original_tdocs_textbox = gui.main.get_text_with_scrollbar(current_row, 1, height=4, current_main_frame=top, width=100)
        
        current_row += 1
        tkinter.Label(top, text='Final TDocs:').grid(row=current_row, column=0)
        self.final_tdocs_textbox = gui.main.get_text_with_scrollbar(current_row, 1, height=4, current_main_frame=top, width=100)

        def set_original_tdocs(*args):
            self.original_tdocs_textbox.delete('1.0', tkinter.END)
            self.original_tdocs_textbox.insert(tkinter.END, self.tkvar_original_tdocs.get())
        self.tkvar_original_tdocs.trace('w', set_original_tdocs)
        
        def set_final_tdocs(*args):
            self.final_tdocs_textbox.delete('1.0', tkinter.END)
            self.final_tdocs_textbox.insert(tkinter.END, self.tkvar_final_tdocs.get())
        self.tkvar_final_tdocs.trace('w', set_final_tdocs)

        # Configure column row widths
        top.grid_columnconfigure(0, weight=1)
        top.grid_columnconfigure(1, weight=1)
        top.grid_columnconfigure(2, weight=1)
        top.grid_columnconfigure(3, weight=1)

    def get_tdocs_of_selected_meeting(self):
        selected_meeting = gui.main.tkvar_meeting.get()
        meeting_folder = application.sa2_meeting_data.get_server_folder_for_meeting_choice(selected_meeting)
        local_agenda_file = gui.main.get_tdocs_by_agenda_file_or_url(server.get_local_tdocs_by_agenda_filename(meeting_folder))
        tdocs_df = html_parser.tdocs_by_agenda(local_agenda_file).tdocs
        return tdocs_df

    def analyze_tdoc(self):
        if not gui.main.current_tdocs_by_agenda_exists():
            # Remove all prior elements and return
            return
        
        # Show all data
        tdocs_df = self.get_tdocs_of_selected_meeting()

        last_tdoc = gui.main.tkvar_last_doc_tdoc.get()
        if not tdoc.is_tdoc(last_tdoc):
            print(str(last_tdoc) + ' is not a TDoc')
            return
        try:
            tdoc_row = tdocs_df.loc[last_tdoc,:]
            self.tkvar_tdoc.set(gui.main.tkvar_last_doc_tdoc.get())
            self.tkvar_original_tdocs.set('')
            self.tkvar_final_tdocs.set('')

            original_tdocs = html_parser.get_tdoc_infos(tdoc_row['Original TDocs'], tdocs_df)
            final_tdocs    = html_parser.get_tdoc_infos(tdoc_row['Final TDocs'], tdocs_df)

            original_tdocs_text = '\n'.join(['{0} ({1}); {2}'.format(tdoc.tdoc, tdoc.source, tdoc.title) for tdoc in original_tdocs])
            final_tdocs_text    = '\n'.join(['{0} ({1}); {2}'.format(tdoc.tdoc, tdoc.source, tdoc.title) for tdoc in final_tdocs])

            self.tkvar_original_tdocs.set(original_tdocs_text)
            self.tkvar_final_tdocs.set(final_tdocs_text)
        except:
            print('Could not retrieve data for ' + last_tdoc)
            traceback.print_exc()

    def open_local_meeting_folder(self):
        selected_meeting = gui.main.tkvar_meeting.get()
        meeting_folder   = application.sa2_meeting_data.get_server_folder_for_meeting_choice(selected_meeting)
        if meeting_folder is not None:
            local_folder = server.get_meeting_folder(meeting_folder)
            os.startfile(local_folder)

    def open_server_meeting_folder(self):
        selected_meeting = gui.main.tkvar_meeting.get()
        meeting_folder   = application.sa2_meeting_data.get_server_folder_for_meeting_choice(selected_meeting)
        if meeting_folder is not None:
            remote_folder = server.get_remote_meeting_folder(meeting_folder)
            os.startfile(remote_folder)

    def generate_word_report(self):
        current_tdocs_by_agenda = application.current_tdocs_by_agenda
        if current_tdocs_by_agenda is None:
            print('Could not generate report. No current TDocsByAgenda found')
            return
        selected_meeting = gui.main.tkvar_meeting.get()
        server_folder = application.sa2_meeting_data.get_server_folder_for_meeting_choice(selected_meeting)
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

        meeting_folder = server.get_local_agenda_folder(server_folder)

        df_tdocs = current_tdocs_by_agenda.tdocs
        email_approval_tdocs = df_tdocs[(df_tdocs['Result'] == 'For e-mail approval')]
        n_email_approval = len(email_approval_tdocs)
        print('TDocsByAgenda: {0} TDocs marked as "For e-mail approval"'.format(n_email_approval))

        # Force Agenda cache for the section titles
        server.get_last_agenda(server_folder)

        doc = word_parser.open_word_document()
        word_parser.insert_doc_data_to_doc_by_wi(
            df_tdocs, 
            doc, 
            server_folder,
            source=application.word_own_reporter_name,
            ais_to_output=ais_to_output,
            save_to_folder=meeting_folder)
        print('Finished exporting Word report')

    # ToDo: change output_meeting so that it can be a string!!!
    def export_tdocs_by_agenda_to_excel(self, selected_meeting=None, output_meeting=None, filename='TdocsByAgenda', close_file=False, use_thread=True, current_dt_str=None):
        try:
            if selected_meeting is None:
                selected_meeting = gui.main.tkvar_meeting.get()
            if output_meeting is None:
                output_meeting   = gui.main.tkvar_meeting.get()

            input_meeting_folder  = application.sa2_meeting_data.get_server_folder_for_meeting_choice(selected_meeting)
            output_meeting_folder = application.sa2_meeting_data.get_server_folder_for_meeting_choice(output_meeting)
            inbox_active          = gui.main.inbox_is_for_this_meeting()

            input_local_agenda_file = server.download_agenda_file(selected_meeting, inbox_active)
            if input_local_agenda_file is None:
                print('Could not find agenda file for {0}'.format(selected_meeting))
                return

            (head,tail) = os.path.split(server.get_local_tdocs_by_agenda_filename(output_meeting_folder))
            if current_dt_str is None:
                current_dt_str = application.get_now_time_str()
            excel_export = os.path.join(head, '{0} {1}.xlsx'.format(current_dt_str, filename))
            print('Exporting ' + input_local_agenda_file + ' to ' + excel_export)
            
            # Disable buttons
            self.export_button.config(text='Exporting... DO NOT interrupt Excel until COMPLETELY finished!', state='disabled')
            self.year_entry.config(state='readonly')
            self.tdoc_report_button.config(text='Exporting... DO NOT interrupt Excel until COMPLETELY finished!', state='disabled')

            if use_thread:
                t = threading.Thread(target=lambda:self.export_and_open_excel(input_local_agenda_file, excel_export, input_meeting_folder, selected_meeting,close_file))
                t.start()
            else:
                self.export_and_open_excel(input_local_agenda_file, excel_export, input_meeting_folder, selected_meeting,close_file)
        except:
            print('Could not export TDocs by agenda data')
            traceback.print_exc()

    def export_and_open_excel(self, local_agenda_file, excel_export, meeting_folder, sheet_name, close_file=False, writer=None):
        try:
            tdocs_by_agenda = html_parser.tdocs_by_agenda(gui.main.get_tdocs_by_agenda_file_or_url(local_agenda_file))

            # Do not export to Excel the last columns (just a lot of True/False columns for each vendor)
            tdocs_df = tdocs_by_agenda.tdocs.iloc[:, 0:19]
            
            tdocs = tdocs_df.index.tolist()
            server_urls = [(tdoc,server.get_remote_filename(meeting_folder, tdoc, use_inbox=False)) for tdoc in tdocs]
            tdocs_df.loc[:,parsing.excel.session_comments_column] = ''

            # Get TDoc comments from the comments files
            agenda_folder   = os.path.dirname(os.path.abspath(local_agenda_file))
            parsed_comments = parsing.excel.get_comments_from_dir_format(agenda_folder, merge_comments=True)
            fg_color   = {}
            text_color = {}
            if parsed_comments is not None:
                comments = []

                # Generate meta-comments for revision of and merge of for easier review
                for tdoc_idx in tdocs:
                    row = tdocs_df.loc[tdoc_idx,:]
                    merge_of    = row['Merge of']
                    revision_of = row['Revision of']
                    session_comments = row[parsing.excel.session_comments_column]

                    tdoc_parent_list = []
                    if revision_of != '':
                        tdoc_parent_list.append(revision_of)
                    if merge_of != '':
                        merge_of_parent_list = [e.strip() for e in merge_of.split(',') if e.strip()!='']
                        tdoc_parent_list.extend(merge_of_parent_list)

                    comments_for_this_tdoc = [ (parent_tdoc, parsed_comments[parent_tdoc]) for parent_tdoc in tdoc_parent_list if parent_tdoc in parsed_comments and len(parsed_comments[parent_tdoc]) > 0 ]
                    # Only if there are no comments and there are revisions to do
                    if len(comments_for_this_tdoc)>0:
                        if len(comments_for_this_tdoc)>0:
                            list_of_sublists = [tup[1] for tup in comments_for_this_tdoc]
                            fg_colors   = [e[2] for sublist in list_of_sublists for e in sublist]
                            text_colors = [e[3] for sublist in list_of_sublists for e in sublist]
                            fg_color   = parsing.excel.get_reddest_color(fg_colors)
                            text_color = parsing.excel.get_reddest_color(text_colors)
                        else:
                            fg_color   = None
                            text_color = None
                        merged_texts = []
                        for comment in comments_for_this_tdoc:
                            comment_tdoc = comment[0]
                            comment_list = comment[1]
                            comment_list = [parsing.excel.get_comment_full_text(e[0], '{0}'.format(e[1])) for e in comment_list]
                            merged_comment_list = '\n'.join(comment_list)
                            merged_comment = parsing.excel.get_comment_full_text(comment_tdoc, '{{\n{0}\n}}'.format(merged_comment_list))
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
                    comment_list_txt = [ parsing.excel.get_comment_full_text(comment[0], comment[1],) for comment in comment_list ]
                    full_comment = '\n'.join(comment_list_txt)
                    try:
                        tdocs_df.at[idx,parsing.excel.session_comments_column] = full_comment
                    except:
                        print('Did not find TDoc entry for comment {0}. Skipping'.format(idx))
                fg_color,text_color = parsing.excel.get_colors_from_comments(parsed_comments)

            # Need Pandas 0.24 for this .See https://stackoverflow.com/questions/42589835/adding-a-pandas-dataframe-to-existing-excel-file
            # or just use https://github.com/pandas-dev/pandas/issues/3441#issuecomment-24898286
            # Needs Pandas >= 0.24.0
            # Note that xlsxwriter does NOT support append mode
            if os.path.isfile(excel_export):
                write_mode = 'a'
            else:
                write_mode = 'w'
            with pandas.ExcelWriter(excel_export, engine = 'openpyxl', mode=write_mode) as writer:
                tdocs_df.to_excel(writer, sheet_name=sheet_name)  

            parsing.excel.apply_comments_coloring_and_hyperlinks(excel_export, fg_color, text_color, server_urls)
            
            # Need to reinitialize COM on each thread
            # https://stackoverflow.com/questions/26745617/win32com-client-dispatch-cherrypy-coinitialize-has-not-been-called
            pythoncom.CoInitialize()
            wb = excel_parser.open_excel_document(excel_export, excel=excel_parser.get_excel(), sheet_name=sheet_name)
            
            excel_parser.set_first_row_as_filter(wb)
            excel_parser.adjust_tdocs_by_agenda_column_width(wb)
            excel_parser.set_tdoc_colors(wb, server_urls)
            excel_parser.vertically_center_all_text(wb)
            excel_parser.save_wb(wb)
            if close_file:
                excel_parser.close_wb(wb)

            print('Non-parsed vendors: {0}'.format('\n'.join(tdocs_by_agenda.others_cosigners)))
        except:
            print('Could not export TDoc list + comments to Excel')
            traceback.print_exc()
        finally:
            self.export_button.config(text=ToolsDialog.export_text,state='normal')
            self.tdoc_report_button.config(text=ToolsDialog.export_year_text,state='normal')
            self.year_entry.config(state='normal')

    def outlook_email_approval(self):
        current_text = self.outlook_button_text.get()
        self.outlook_button_text.set('Exporting... DO NOT interrupt Outlook until COMPLETELY finished!')
        self.email_approval_button.config(state='disabled')
        self.email_attachments_generate_summary_checkbox.config(state='disabled')
        t = threading.Thread(target=lambda:self.on_outlook_email_approval(current_text))
        t.start()

    def on_outlook_email_approval(self, current_text):
        try:
            # Need to reinitialize COM on each thread
            # https://stackoverflow.com/questions/26745617/win32com-client-dispatch-cherrypy-coinitialize-has-not-been-called
            pythoncom.CoInitialize()
            selected_meeting = gui.main.tkvar_meeting.get()
            generate_summary = self.outlook_generate_summary.get()
            parsing.outlook.process_email_approval(selected_meeting, generate_summary)
        finally:
            self.outlook_button_text.set(current_text)
            self.email_approval_button.config(state='normal')
            self.email_attachments_generate_summary_checkbox.config(state='normal')

    def outlook_email_attachments(self):
        self.email_attachments_button.config(text='Processing... DO NOT interrupt Outlook until COMPLETELY finished!', state='disabled')
        t = threading.Thread(target=lambda:self.on_outlook_email_attachments())
        t.start()

    def on_outlook_email_attachments(self):
        try:
            # Need to reinitialize COM on each thread
            # https://stackoverflow.com/questions/26745617/win32com-client-dispatch-cherrypy-coinitialize-has-not-been-called
            pythoncom.CoInitialize()
            parsing.outlook.process_email_attachments()
        finally:
            self.email_attachments_button.config(text=ToolsDialog.outlook_attachment_text,state='normal')

    def export_year_tdocs_by_agenda_to_excel(self):
        try:
            year_to_search = int(self.tkvar_year.get())
            meetings_to_check = application.sa2_meeting_data.get_meetings_for_given_year(year_to_search)
            output_meeting = gui.main.tkvar_meeting.get()
            
            t = threading.Thread(target=lambda:self.on_export_year_tdocs_by_agenda_to_excel(meetings_to_check, output_meeting, 'Report {0}'.format(year_to_search)))
            t.start()
        except:
            print('Could not generate yearly report')
            traceback.print_exc()

    def on_export_year_tdocs_by_agenda_to_excel(self, meetings_to_check, output_meeting, filename):
        last_folder = meetings_to_check.meeting_folders[-1]
        current_dt_str = application.get_now_time_str()

        meeting_folder = application.sa2_meeting_data.get_server_folder_for_meeting_choice(output_meeting)
        local_agenda_file = server.get_local_tdocs_by_agenda_filename(meeting_folder)
        (head,tail) = os.path.split(local_agenda_file)
        excel_export = os.path.join(head, '{0} {1}.xlsx'.format(current_dt_str, filename))
        for meeting_to_check in meetings_to_check.meeting_folders:
            print(meeting_to_check)
            close_file = meeting_to_check != last_folder
            self.export_tdocs_by_agenda_to_excel(
                selected_meeting=meeting_to_check, 
                output_meeting=output_meeting, 
                filename=filename,
                close_file=close_file,
                use_thread=False,
                current_dt_str=current_dt_str)

    def cache_tdocs(self, tdoc_list):
        if tdoc_list is None:
            return

        gui.main.open_downloaded_tdocs = False
        for tdoc_to_download in tdoc_list:
            try:
                gui.main.tkvar_tdoc_id.set(tdoc_to_download)
                gui.main.download_and_open_tdoc()
            except:
                print('Could not download TDoc {0}'.format(tdoc_to_download))
                traceback.print_exc()
        gui.main.open_downloaded_tdocs = True

    def bulk_cache_ais(self):
        try:
            tdocs = self.get_tdocs_of_selected_meeting()
            ais = self.tkvar_ai_list.get()
            if ais is None:
                return
            if ais == '':
                ais = None
            else:
                ais = [ai.strip() for ai in ais.replace(',', '').split(' ') if (ai is not None) and (ai != '')]
            tdocs_do_cache = []
            if ais is None:
                ai_tdocs = tdocs.index.tolist()
                if len(ai_tdocs) > 0:
                    tdocs_do_cache.extend(ai_tdocs)
                print('{0} items in meeting: {1}'.format(len(ai_tdocs), len(tdocs.index)))
            else:
                for ai in ais:
                    ai_tdocs = tdocs.index[tdocs['AI']==ai].tolist()
                    if len(ai_tdocs) > 0:
                        tdocs_do_cache.extend(ai_tdocs)
                    print('{0} items in AI {1}, total items: {2}'.format(len(ai_tdocs), ai, len(tdocs.index)))
            self.cache_tdocs(tdocs_do_cache)
        except:
            print('General error performing bulk AI caching')
            traceback.print_exc()