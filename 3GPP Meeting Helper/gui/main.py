import os.path
import threading
import time
import tkinter
import tkinter.font
import tkinter.ttk
import traceback

from pyperclip import copy as clipboard_copy

import application.word
import application.meeting_helper
import gui.config
import gui.tools
import parsing.html as html_parser
import parsing.word
import server.common
import server.tdoc
import tdoc.utils
from gui.common import favicon

# tkinter initialization
root = tkinter.Tk()
root.title("3GPP SA2 Meeting helper")

root.iconbitmap(favicon)

# Add a grid
main_frame = tkinter.Frame(root)
main_frame.grid(column=0, row=0, sticky=(tkinter.N, tkinter.W, tkinter.E, tkinter.S))


def set_waiting_for_proxy_message():
    return gui.common.set_waiting_for_proxy_message(main_frame)


# global variables
inbox_tdoc_list_html = None
performing_search = False
open_downloaded_tdocs = True

# Tkinter variables
tkvar_meeting = tkinter.StringVar(root)
tkvar_inbox_meeting = tkinter.StringVar(root)
tkvar_inbox_meeting_label = tkinter.StringVar(root)
tkinter_label_3gpp = tkinter.IntVar(root)
tkinter_label_sync = tkinter.IntVar(root)
tkinter_label_inbox = tkinter.IntVar(root)

tkvar_last_agenda_version = tkinter.StringVar(root)
tkvar_last_agenda_vtext = tkinter.StringVar(root)
tkvar_tdoc_download_result = tkinter.StringVar()
tkvar_tdoc_id = tkinter.StringVar(root)
tkvar_tdoc_id_full = tkinter.StringVar(root)
tkvar_follow_current_tdoc = tkinter.IntVar(root)
tkvar_search_tdoc = tkinter.IntVar(root)

tkvar_tdocs_by_agenda_exist = tkinter.BooleanVar(root)
tkvar_last_doc_tdoc = tkinter.StringVar(root)
tkvar_last_doc_title = tkinter.StringVar(root)
tkvar_last_doc_source = tkinter.StringVar(root)
tkvar_last_tdoc_url = tkinter.StringVar(root)
tkvar_last_tdoc_status = tkinter.StringVar(root)

tkvar_override_tdocs_by_agenda = tkinter.BooleanVar(root)
tkvar_tdocs_by_agenda_path = tkinter.StringVar(root)
tkvar_tdocs_by_agenda_path.set('')

# Initial (static) values
tkinter_label_3gpp.set(1)
tkinter_label_sync.set(0)
tkinter_label_inbox.set(0)
tkvar_last_agenda_version.set('')
tkvar_tdoc_download_result.set('')
tkvar_tdoc_id.set('S2-XXXXXXX')
tkvar_follow_current_tdoc.set(0)
tkvar_search_tdoc.set(0)
tkvar_tdocs_by_agenda_exist.set(False)

tkvar_last_doc_tdoc.set('')
tkvar_last_doc_title.set('')
tkvar_last_doc_source.set('')
tkvar_last_tdoc_url.set('')

tkvar_inbox_from_selected_meeting = tkinter.BooleanVar(root)

# Tkinter elements that require variables
open_tdoc_button = tkinter.Button(main_frame, textvariable=tkvar_tdoc_id_full)
tdoc_entry = tkinter.Entry(main_frame, textvariable=tkvar_tdoc_id, width=25, font='TkDefaultFont')
open_last_agenda_button = tkinter.Button(main_frame, text='Open last agenda')
meeting_ftp_button = tkinter.Checkbutton(main_frame, state='disabled', variable=tkinter_label_inbox)
tdocs_by_agenda_entry = tkinter.Entry(main_frame, textvariable=tkvar_tdocs_by_agenda_path, width=25,
                                      font='TkDefaultFont')

# Other variables
last_override_tdocs_by_agenda = ''


# Utility methods
def inbox_is_for_this_meeting():
    meeting_number_from_dropdown = tkvar_meeting.get().split(',')[0]
    return tkvar_inbox_meeting.get() == meeting_number_from_dropdown


def set_selected_meeting_to_inbox_meeting():
    # Sets the selected meeting to the current inbox meeting
    if application.meeting_helper.sa2_meeting_data is None:
        return
    tkvar_meeting.set(application.meeting_helper.sa2_meeting_data.get_meeting_text_for_given_meeting_number(tkvar_inbox_meeting.get()))


def reset_status_labels():
    tkvar_last_agenda_version.set('')
    tkvar_tdoc_download_result.set('')
    # Set default TDoc name
    current_meeting = tkvar_meeting.get()
    if application.meeting_helper.sa2_meeting_data is None:
        return
    year = application.meeting_helper.sa2_meeting_data.get_year_from_meeting_text(current_meeting)
    if year is not None:
        try:
            if not performing_search:
                current_value = tkvar_tdoc_id.get()
                if not tdoc.utils.is_tdoc(current_value):
                    tkvar_tdoc_id.set('S2-' + str(year)[2:4] + 'XXXXX')
        except:
            pass


def update_ftp_button():
    meeting_ftp_button.config(text=server.common.private_server + ' (FTP)')


def get_tdocs_by_agenda_file_or_url(target):
    override_target = tkvar_tdocs_by_agenda_path.get()
    if override_target != '':
        print('Target TDocsByAgenda overridden with {0}'.format(override_target))
        return override_target
    else:
        print('Target TDocsByAgenda: not overridden')
    return target


def load_application_data():
    global inbox_tdoc_list_html

    file_to_parse = server.tdoc.get_sa2_inbox_tdoc_list()

    # In case we override the file
    inbox_tdoc_list_html = get_tdocs_by_agenda_file_or_url(file_to_parse)
    application.meeting_helper.current_tdocs_by_agenda = html_parser.get_tdocs_by_agenda_with_cache(inbox_tdoc_list_html)

    # Load SA2 meeting data
    sa2_folder_html = server.common.get_sa2_folder()
    application.meeting_helper.sa2_meeting_data = html_parser.parse_3gpp_meeting_list_object(
        sa2_folder_html,
        ordered=True, remove_old_meetings=True)

    # Double-check
    df_tdocs = application.meeting_helper.current_tdocs_by_agenda.tdocs
    email_approval_tdocs = df_tdocs[(df_tdocs['Result'] == 'For e-mail approval')]
    n_email_approval = len(email_approval_tdocs)
    print('Current TDocsByAgenda: {0} TDocs marked as "For e-mail approval" after de-duplication'.format(
        n_email_approval))


# Variable-change callbacks
def set_inbox_label(*args):
    tkvar_inbox_meeting_label.set('Inbox meeting: {0}'.format(tkvar_inbox_meeting.get()))


tkvar_inbox_meeting.trace('w', set_inbox_label)


def set_agenda_version_text(*args):
    current_version = tkvar_last_agenda_version.get()
    if (current_version is None) or (current_version == ''):
        tkvar_last_agenda_vtext.set('No known last agenda')
    else:
        tkvar_last_agenda_vtext.set('Last Agenda: ' + tkvar_last_agenda_version.get())


tkvar_last_agenda_version.trace('w', set_agenda_version_text)


def set_inbox_from_selected_meeting_state():
    # Checks whether the inbox is from the selected meeting and sets
    # some labels accordingly
    tkvar_inbox_from_selected_meeting.set(inbox_is_for_this_meeting())
    if inbox_is_for_this_meeting():
        tkinter_label_sync.set(1)
        if server.common.we_are_in_meeting_network(searching_for_a_file=True):
            tkinter_label_inbox.set(1)
        else:
            tkinter_label_inbox.set(0)
    else:
        tkinter_label_sync.set(0)


def change_meeting_dropdown(*args):
    set_inbox_from_selected_meeting_state()
    reset_status_labels()
    open_tdocs_by_agenda(open_this_file=False)


tkvar_meeting.trace('w', change_meeting_dropdown)


def on_follow_current_doc_change(*args):
    follow_current_tdoc = tkvar_follow_current_tdoc.get()
    if follow_current_tdoc:
        set_selected_meeting_to_inbox_meeting()
        tdoc_entry.config(state='readonly')
        # Force update
        retrieve_current_doc_yes()
    else:
        tdoc_entry.config(state='normal')


tkvar_follow_current_tdoc.trace('w', on_follow_current_doc_change)


# Text boxes
def get_text_with_scrollbar(row, column, height=2, current_main_frame=main_frame, width=50):
    scrollbar = tkinter.Scrollbar(current_main_frame)
    text = tkinter.Text(current_main_frame, height=height, width=width)
    scrollbar.config(command=text.yview)
    text.config(yscrollcommand=scrollbar.set)

    text.grid(row=row, column=column, columnspan=2)
    scrollbar.grid(row=row, column=column + 2, sticky=tkinter.N + tkinter.S + tkinter.W)
    return text


# Current doc checker thread
def retrieve_current_doc_yes():
    current_tdoc_html = server.tdoc.get_sa2_inbox_current_tdoc(searching_for_a_file=True)
    current_tdoc = html_parser.parse_current_document(current_tdoc_html)
    if current_tdoc is not None:
        tkvar_tdoc_id.set(current_tdoc)


def retrieve_current_doc():
    while True:
        # Case when we change the WLAN during the meeting
        set_inbox_from_selected_meeting_state()
        if tkvar_follow_current_tdoc.get():
            retrieve_current_doc_yes()
        else:
            pass
        time.sleep(10)


def search_netovate():
    """
    Search the Netovate website for a specific TDoc
    """
    tdoc_id = tkvar_tdoc_id.get()
    netovate_url = 'http://netovate.com/doc-search/?fname={0}'.format(tdoc_id)
    print('Opening {0}'.format(netovate_url))
    os.startfile(netovate_url)

def start_check_current_doc_thread():
    t = threading.Thread(target=retrieve_current_doc)
    t.start()


# Downloads the TDocs by Agenda file
def open_tdocs_by_agenda(open_this_file=True):
    try:
        (meeting_server_folder, local_file) = get_local_tdocs_by_agenda_filename_for_current_meeting()
        if meeting_server_folder is None:
            return
    except:
        return

    # Save opened Tdocs by Agenda file to global application
    html = get_tdocs_by_agenda_for_selected_meeting(meeting_server_folder)
    print('Retrieved local TDocsByAgenda data for {0}. Parsing TDocs'.format(meeting_server_folder))
    application.current_tdocs_by_agenda = html_parser.get_tdocs_by_agenda_with_cache(
        html,
        meeting_server_folder=meeting_server_folder)

    print('-------------------')
    print('Current TDocsByAgenda are {0}'.format(application.current_tdocs_by_agenda.meeting_number))
    print(application.current_tdocs_by_agenda.tdocs.index)
    print('-------------------')

    server.common.write_data_and_open_file(html, local_file)
    return application.current_tdocs_by_agenda


def get_tdocs_by_agenda_for_selected_meeting(
        meeting_folder,
        return_revisions_file=False,
        return_drafts_file=False):
    inbox_active = inbox_is_for_this_meeting()
    return_data = server.tdoc.get_tdocs_by_agenda_for_selected_meeting(meeting_folder, inbox_active)

    # Optional download of revisions
    revisions_file = None
    if return_revisions_file:
        try:
            revisions_file = server.tdoc.download_revisions_file(meeting_folder)
        except:
            # Not all meetings have revisions
            # traceback.print_exc()
            pass

    # Optional download of drafts
    drafts_file = None
    if return_drafts_file:
        try:
            drafts_file = server.tdoc.download_drafts_file(meeting_folder)
        except:
            # Not all meetings have drafts
            # traceback.print_exc()
            pass

    if return_revisions_file and return_drafts_file:
        return return_data, revisions_file, drafts_file

    if return_revisions_file and not return_drafts_file:
        return return_data, revisions_file

    return return_data


def get_local_tdocs_by_agenda_filename_for_current_meeting():
    try:
        current_selection = tkvar_meeting.get()
        if (current_selection is None) or (current_selection == ''):
            print('Empty current selection: current meeting not yet selected')
            return None, None
        else:
            print('Get TdocsByAgenda for {0} from local file'.format(current_selection))
        meeting_server_folder = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(current_selection)
        local_file = server.tdoc.get_local_tdocs_by_agenda_filename(meeting_server_folder)

        return meeting_server_folder, local_file
    except:
        print('Could not retrieve local TdocsByAgenda filename for current meeting')
        traceback.print_exc()
        return None


def current_tdocs_by_agenda_exists():
    try:
        (meeting_server_folder, local_file) = get_local_tdocs_by_agenda_filename_for_current_meeting()
        return os.path.isfile(local_file)
    except:
        return False

# Trigger to search TDoc with Enter key
def on_press_enter_key(event=None):
    print('<Return>> key pressed')
    try:
        button_status = open_tdoc_button['state']
    except:
        button_status = tkinter.DISABLED

    print('button_status={0}'.format(button_status))
    if button_status == tkinter.NORMAL:
        download_and_open_tdoc()


# Button to open TDoc
def download_and_open_tdoc(
        tdoc_id_to_override=None,
        cached_tdocs_list=None,
        copy_to_clipboard=False,
        skip_opening=False):
    global performing_search
    tkvar_tdoc_id.set(tkvar_tdoc_id.get().replace(' ', '').replace('\r', '').replace('\n', '').strip())
    if tdoc_id_to_override is None:
        # Normal flow
        tdoc_id = tkvar_tdoc_id.get()
    else:
        # Used to compare two tdocs
        tdoc_id = tdoc_id_to_override
    download_from_inbox = inbox_is_for_this_meeting()
    retrieved_files, tdoc_url = server.tdoc.get_tdoc(
        application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(tkvar_meeting.get()),
        tdoc_id,
        use_inbox=download_from_inbox,
        return_url=True,
        searching_for_a_file=True)
    if skip_opening:
        return retrieved_files
    if copy_to_clipboard:
        if tdoc_url is None:
            clipboard_text = tdoc_id
        else:
            clipboard_text = '{0}, {1}'.format(tdoc_id, tdoc_url)
        clipboard_copy(clipboard_text)
        print('Copied TDoc ID and URL (if available) to clipboard: {0}'.format(clipboard_text))
    # Set file information if available
    tkvar_last_tdoc_url.set(tdoc_url)
    # Set current status if found
    try:
        tdoc_status = application.meeting_helper.current_tdocs_by_agenda.tdocs.at[tdoc_id, 'Result']
        if tdoc_status is None:
            tdoc_status = ''
    except:
        tdoc_status = '<unknown>'
    tkvar_last_tdoc_status.set(tdoc_status)
    try:
        # ToDo: download current TDocs by agenda
        pass
    except:
        print('Could not load TDoc agenda info for {0}'.format(tkvar_tdoc_id.get()))
    if (retrieved_files is None) or (len(retrieved_files) == 0):
        tdoc_year, tdoc_number = tdoc.utils.get_tdoc_year(tdoc_id)
        # Search on meetings from the given year if the TDoc is not found
        if tkvar_search_tdoc.get() and (tdoc_year is not None) and (not performing_search):
            # Retrieve search for all meetings of the year
            performing_search = True
            try:
                # Search while we still did not find a match
                meetings_to_check = application.meeting_helper.sa2_meeting_data.get_meetings_for_given_year(tdoc_year)
                print('Will search meetings: {0}'.format('; '.join(meetings_to_check.meeting_folders)))
                for meeting_to_search in meetings_to_check.meeting_folders:
                    tkvar_meeting.set(meeting_to_search)
                    download_and_open_tdoc()
                    if not performing_search:
                        not_found_string = None
                        break
                    not_found_string = 'Not found (' + tdoc_id + ')'
            finally:
                performing_search = False
        else:
            not_found_string = 'Not found (' + tdoc_id + ')'

        if not_found_string is not None:
            tkvar_tdoc_download_result.set(not_found_string)
    else:
        if not open_downloaded_tdocs:
            found_string = 'Cached file(s)'
            opened_files = 0
            metadata_list = []
            if cached_tdocs_list is not None and isinstance(cached_tdocs_list, list):
                cached_tdocs_list.extend(retrieved_files)
        else:
            opened_files, metadata_list = parsing.word.open_files(retrieved_files, return_metadata=True)
            found_string = 'Opened {0} file(s)'.format(opened_files)
        tkvar_last_doc_tdoc.set(tkvar_tdoc_id.get())
        if performing_search:
            found_string += ' (' + tkvar_meeting.get() + ')'
        performing_search = False
        tkvar_tdoc_download_result.set(found_string)
        if len(metadata_list) > 0:
            last_metadata = metadata_list[-1]
            if last_metadata is not None:
                if last_metadata.title is not None:
                    tkvar_last_doc_title.set(last_metadata.title)
                if last_metadata.source is not None:
                    tkvar_last_doc_source.set(last_metadata.source)
        return retrieved_files


def start_main_gui():
    load_application_data()

    tkvar_inbox_meeting.set(application.meeting_helper.current_tdocs_by_agenda.meeting_number)
    tkvar_meeting.set(application.meeting_helper.sa2_meeting_data.get_meeting_text_for_given_meeting_number(
        application.meeting_helper.current_tdocs_by_agenda.meeting_number))

    popupMenu = tkinter.OptionMenu(main_frame, tkvar_meeting, *application.meeting_helper.sa2_meeting_data.meeting_folders)

    # Variable-change callbacks
    def set_tdoc_id_full(*args):
        # Sets the label for the download button
        tdoc_id = tkvar_tdoc_id.get()
        if tkvar_search_tdoc.get():
            command_string = 'Search'
        else:
            command_string = 'Open'
        button_label = command_string
        if tdoc.utils.is_tdoc(tdoc_id):
            button_label += ' ' + tdoc_id
        tkvar_tdoc_id_full.set(button_label)
        if tdoc.utils.is_tdoc(tdoc_id):
            # Enable button
            open_tdoc_button.configure(state=tkinter.NORMAL)
        else:
            # Disable button
            open_tdoc_button.configure(state=tkinter.DISABLED)

    set_tdoc_id_full()
    tkvar_tdoc_id.trace('w', set_tdoc_id_full)
    tkvar_search_tdoc.trace('w', set_tdoc_id_full)

    # Set initial selection to the inbox meeting (should be the current one)
    set_selected_meeting_to_inbox_meeting()

    # Row: Inbox info
    current_row = 0
    open_last_agenda_button.grid(row=0, column=0, sticky="EW")
    popupMenu.grid(row=current_row, column=1, sticky="ew", padx=10, pady=10)
    tkinter.Button(main_frame, text='TDocs by Agenda', command=open_tdocs_by_agenda).grid(row=current_row, column=2,
                                                                                          padx=10, pady=10, sticky="EW")

    # Row: Dropdown menu and meeting info
    current_row += 1
    tkinter.Button(main_frame, text='Network config', command=lambda: gui.config.NetworkConfigDialog(root, favicon,
                                                                                                     on_update_ftp=gui.main.update_ftp_button)).grid(
        row=current_row, column=0, sticky="EW")
    tkinter.Checkbutton(main_frame, text='3GPP sync (HTTP)', state='disabled', variable=tkinter_label_sync).grid(
        row=current_row, column=1)
    update_ftp_button()
    meeting_ftp_button.grid(row=current_row, column=2)

    # Row: Open TDoc
    current_row += 1
    tkinter.Label(main_frame, text="Open TDoc").grid(row=current_row, column=0)
    tdoc_entry.grid(row=current_row, column=1, padx=10, pady=10)
    tkinter.Checkbutton(main_frame, text='Track current TDoc ', variable=tkvar_follow_current_tdoc).grid(
        row=current_row, column=2)
    current_row += 1
    tkinter.Checkbutton(main_frame, text='Search if not found', variable=tkvar_search_tdoc).grid(row=current_row,
                                                                                                 column=2)

    # Rows: Download TDoc button and last agenda
    tkinter.Button(main_frame, text='Tools',
                   command=lambda: gui.tools.ToolsDialog(gui.main.root, gui.main.favicon)).grid(row=current_row,
                                                                                                column=0, sticky="EW")
    open_tdoc_button.configure(command=download_and_open_tdoc)
    root.bind('<Return>', on_press_enter_key) # Bind the enter key in this frame to searching for the TDoc
    print('Bound <Return> key to TDoc search')
    open_tdoc_button.grid(row=current_row, column=1, padx=10, pady=10, sticky="EW")

    # Override TDocs by Agenda if it is malformed
    current_row += 1
    tkinter.Checkbutton(main_frame, text='Override Tdocs by agenda', variable=tkvar_override_tdocs_by_agenda).grid(
        row=current_row, column=0)
    tdocs_by_agenda_entry.config(state='readonly')
    tdocs_by_agenda_entry.grid(row=current_row, column=1, padx=10, pady=10)

    # Add button to check Netovate (useful if you are searching for documents from other WGs
    tkinter.Button(main_frame, text='Search Netovate',
                   command=search_netovate).grid(row=current_row, column=2, sticky="EW")

    def set_override_tdocs_by_agenda_var(*args):
        global last_override_tdocs_by_agenda
        current_value = tkvar_override_tdocs_by_agenda.get()
        if not current_value:
            tdocs_by_agenda_entry.config(state='readonly')
            last_override_tdocs_by_agenda = tkvar_tdocs_by_agenda_path.get()
            tkvar_tdocs_by_agenda_path.set('')
        else:
            tdocs_by_agenda_entry.config(state='normal')
            tkvar_tdocs_by_agenda_path.set(last_override_tdocs_by_agenda)

    def set_override_tdocs_by_agenda_path(*args):
        current_path = tkvar_tdocs_by_agenda_path.get()
        try:
            if os.path.exists(current_path):
                print('Forcing loading TDocs by Agenda from {0}'.format(current_path))
                load_application_data()
        except:
            # Do nothing, path is not valid
            return

    tkvar_override_tdocs_by_agenda.trace('w', set_override_tdocs_by_agenda_var)
    tkvar_tdocs_by_agenda_path.trace('w', set_override_tdocs_by_agenda_path)

    def on_open_last_agenda(*args):
        open_last_agenda_button.config(state='disabled')
        t = threading.Thread(target=open_last_agenda(*args))
        t.start()

    open_last_agenda_button.configure(command=on_open_last_agenda)

    def open_last_agenda(*args):
        try:
            meeting_folder = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(tkvar_meeting.get())
            server.tdoc.get_agenda_files(meeting_folder, use_inbox=False)
            if inbox_is_for_this_meeting():
                server.tdoc.get_agenda_files(meeting_folder, use_inbox=True)
            last_agenda, agenda_version = server.tdoc.get_last_agenda(meeting_folder)
            if last_agenda is not None:
                parsing.word.open_file(last_agenda, go_to_page=-2)
                tkvar_last_agenda_version.set('v' + str(agenda_version))
            else:
                tkvar_last_agenda_version.set('Not found')
        finally:
            open_last_agenda_button.config(state='normal')

    # Row: Infos
    current_row += 1
    tkinter.Label(main_frame, textvariable=tkvar_tdoc_download_result).grid(row=current_row, column=1)
    tkinter.Label(main_frame, textvariable=tkvar_last_agenda_vtext).grid(row=current_row, column=2)

    # Row: info from last document
    current_row += 1
    tkinter.ttk.Separator(main_frame, orient=tkinter.HORIZONTAL).grid(row=current_row, columnspan=3,
                                                                      sticky=(tkinter.W, tkinter.E))

    current_row += 1
    tkinter.Label(main_frame, text='Last document:').grid(row=current_row, column=0)

    # Last opened document    
    def set_last_doc_title(*args):
        last_tdoc_title.delete('1.0', tkinter.END)
        last_tdoc_title.insert(tkinter.END, tkvar_last_doc_title.get())

    tkvar_last_doc_title.trace('w', set_last_doc_title)

    def set_last_doc_source(*args):
        last_tdoc_source.delete('1.0', tkinter.END)
        last_tdoc_source.insert(tkinter.END, tkvar_last_doc_source.get())

    tkvar_last_doc_source.trace('w', set_last_doc_source)

    def set_last_doc_url(*args):
        last_tdoc_url.delete('1.0', tkinter.END)
        last_tdoc_url.insert(tkinter.END, tkvar_last_tdoc_url.get())

    tkvar_last_tdoc_url.trace('w', set_last_doc_url)

    def set_last_doc_status(*args):
        last_tdoc_status.delete('1.0', tkinter.END)
        last_tdoc_status.insert(tkinter.END, tkvar_last_tdoc_status.get())

    tkvar_last_tdoc_status.trace('w', set_last_doc_status)

    current_row += 1
    tkinter.Label(main_frame, text='Title:').grid(row=current_row, column=0)
    last_tdoc_title = get_text_with_scrollbar(current_row, 1)

    current_row += 1
    tkinter.Label(main_frame, text='Source:').grid(row=current_row, column=0)
    last_tdoc_source = get_text_with_scrollbar(current_row, 1)

    current_row += 1
    tkinter.Label(main_frame, text='URL:').grid(row=current_row, column=0)
    last_tdoc_url = get_text_with_scrollbar(current_row, 1, height=1)

    current_row += 1
    tkinter.Label(main_frame, text='Status:').grid(row=current_row, column=0)
    last_tdoc_status = get_text_with_scrollbar(current_row, 1, height=1)

    # Configure column row widths
    main_frame.grid_columnconfigure(0, weight=1)
    main_frame.grid_columnconfigure(1, weight=1)
    main_frame.grid_columnconfigure(2, weight=1)
