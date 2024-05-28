import os.path
import tkinter
import tkinter.font
import tkinter.scrolledtext
import traceback
from tkinter import ttk
from typing import Tuple, List

from pyperclip import copy as clipboard_copy

import application.meeting_helper
import application.word
import gui.meetings_table
import gui.network_config
import gui.specs_table
import gui.tdocs_table
import gui.tools_overview
import gui.work_items_table
import parsing.html.common
import parsing.html.common as html_parser
import parsing.word.pywin32
import server.common
import server.tdoc
import server.tdoc_search
import tdoc.utils
import utils.local_cache
import utils.threading
from application.tkinter_config import root, font_big, ttk_style_tbutton_medium
from gui.common.utils import favicon
from server.specs import get_specs_folder

# tkinter initialization
root.title("3GPP SA2 Meeting helper")

root.iconbitmap(gui.common.utils.favicon)

# Add a grid
main_frame = tkinter.Frame(root)
main_frame.grid(column=0, row=0, sticky=''.join([tkinter.N, tkinter.W, tkinter.E, tkinter.S]))


def set_waiting_for_proxy_message():
    return gui.common.utils.set_waiting_for_proxy_message(main_frame)


# global variables
inbox_tdoc_list_html = None
open_downloaded_tdocs = True

# Tkinter variables
tkvar_meeting = tkinter.StringVar(root)
tk_combobox_meetings = ttk.Combobox(
    main_frame,
    textvariable=tkvar_meeting,
)
tkvar_3gpp_wifi_available = tkinter.BooleanVar(root)


def set_3gpp_network_status_in_application_info(*args):
    application.meeting_helper.last_known_3gpp_network_status = tkvar_3gpp_wifi_available.get()


tkvar_3gpp_wifi_available.trace('w', set_3gpp_network_status_in_application_info)
tkvar_3gpp_wifi_available.set(False)

tkvar_last_agenda_version = tkinter.StringVar(root)
tkvar_last_agenda_vtext = tkinter.StringVar(root)
tkvar_tdoc_download_result = tkinter.StringVar()
tkvar_tdoc_id = tkinter.StringVar(root)
tkvar_tdoc_id_full = tkinter.StringVar(root)
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
tkvar_last_agenda_version.set('')
tkvar_tdoc_download_result.set('')
tkvar_tdoc_id.set('S2-XXXXXXX')
tkvar_search_tdoc.set(0)
tkvar_tdocs_by_agenda_exist.set(False)

tkvar_last_doc_tdoc.set('')
tkvar_last_doc_title.set('')
tkvar_last_doc_source.set('')
tkvar_last_tdoc_url.set('')

# Tkinter elements that require variables
open_tdoc_button = ttk.Button(
    main_frame,
    textvariable=tkvar_tdoc_id_full,
    style=ttk_style_tbutton_medium)
tdoc_entry = tkinter.Entry(
    main_frame,
    textvariable=tkvar_tdoc_id,
    justify='center',
    font=font_big)
open_last_agenda_button = ttk.Button(
    main_frame,
    text='Open last agenda')
tkinter_checkbutton_3gpp_wifi_available = ttk.Checkbutton(
    main_frame,
    state='disabled',
    variable=tkvar_3gpp_wifi_available)
tdocs_by_agenda_entry = tkinter.Entry(
    main_frame,
    textvariable=tkvar_tdocs_by_agenda_path)

# Other variables
last_override_tdocs_by_agenda = ''


# Utility methods
def open_local_meeting_folder(*args):
    selected_meeting = gui.main_gui.tkvar_meeting.get()
    meeting_folder = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(
        selected_meeting)
    if meeting_folder is not None:
        local_folder = utils.local_cache.get_meeting_folder(meeting_folder)
        os.startfile(local_folder)


def open_server_meeting_folder(*args):
    selected_meeting = gui.main_gui.tkvar_meeting.get()
    meeting_folder = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(
        selected_meeting)
    if meeting_folder is not None:
        remote_folder = server.common.get_remote_meeting_folder(meeting_folder)
        os.startfile(remote_folder)


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
            current_value = tkvar_tdoc_id.get()
            if not tdoc.utils.is_sa2_tdoc(current_value):
                tkvar_tdoc_id.set('S2-' + str(year)[2:4] + 'XXXXX')
        except:
            pass


def update_ftp_button():
    tkinter_checkbutton_3gpp_wifi_available.config(text=server.common.private_server + ' (3GPP Wifi)')


def get_tdocs_by_agenda_file_or_url(target):
    override_target = tkvar_tdocs_by_agenda_path.get()
    if override_target != '':
        print('Target TDocsByAgenda overridden with {0}'.format(override_target))
        return override_target
    else:
        print('Target TDocsByAgenda: not overridden')
    return target


def load_application_data(reload_inbox_tdocs_by_agenda=False):
    """
    Load application data necessary for the GUI to work Args: reload_inbox_tdocs_by_agenda: Whether to force a (
    re-)download of https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/SA2/TdocsByAgenda.htm
    """
    global inbox_tdoc_list_html

    tdocs_by_agenda_from_sa2_inbox_bytes = server.tdoc.get_sa2_inbox_tdoc_list(
        open_tdocs_by_agenda_in_browser=False,
        use_cached_file_if_available=not reload_inbox_tdocs_by_agenda)

    # Substitutes the bytes content with a target file path if we want to override
    inbox_tdoc_list_html = get_tdocs_by_agenda_file_or_url(tdocs_by_agenda_from_sa2_inbox_bytes)

    # Parse TdocsByAgenda contents
    application.meeting_helper.current_tdocs_by_agenda = html_parser.get_tdocs_by_agenda_with_cache(
        inbox_tdoc_list_html)

    # Load SA2 meeting data
    sa2_meeting_list_from_server_html = server.common.get_sa2_folder(force_redownload=reload_inbox_tdocs_by_agenda)
    application.meeting_helper.sa2_meeting_data = html_parser.parse_3gpp_meeting_list_object(
        sa2_meeting_list_from_server_html,
        ordered=True,
        remove_old_meetings=True)

    # Double-check
    if application.meeting_helper.current_tdocs_by_agenda is not None:
        df_tdocs = application.meeting_helper.current_tdocs_by_agenda.tdocs
        email_approval_tdocs = df_tdocs[(df_tdocs['Result'] == 'For e-mail approval')]
        n_email_approval = len(email_approval_tdocs)
        print('Current TDocsByAgenda: {0} TDocs marked as "For e-mail approval" after de-duplication'.format(
            n_email_approval))


def set_agenda_version_text(*args):
    current_version = tkvar_last_agenda_version.get()
    if (current_version is None) or (current_version == ''):
        tkvar_last_agenda_vtext.set('No known last agenda')
    else:
        tkvar_last_agenda_vtext.set('Last Agenda: ' + tkvar_last_agenda_version.get())


tkvar_last_agenda_version.trace('w', set_agenda_version_text)


def detect_3gpp_network_state():
    # Checks whether the inbox is from the selected meeting and sets
    # some labels accordingly

    previous_state = tkvar_3gpp_wifi_available.get()
    new_state = server.common.we_are_in_meeting_network()
    if new_state:
        tkvar_3gpp_wifi_available.set(True)
    else:
        tkvar_3gpp_wifi_available.set(False)

    if new_state != previous_state:
        print(f'Changed 3GPP network state from {previous_state} to {new_state}')
        cache_tdocsbyagenda_path = server.tdoc.get_private_server_tdocs_by_agenda_local_cache()
        tdocsbyagenda_url = server.common.tdocs_by_agenda_for_checking_meeting_number_in_meeting

        if new_state:
            # Jumping from normal network to 3GPP Wifi
            # Download from 10.10.10.10 TdocsByAgenda and check meeting number
            # Then freeze the meeting choice to the one found

            meeting_text = None
            if server.common.download_file_to_location(tdocsbyagenda_url, cache_tdocsbyagenda_path):
                if utils.local_cache.file_exists(cache_tdocsbyagenda_path):
                    with open(cache_tdocsbyagenda_path, "r") as f:
                        cache_tdocsbyagenda_html_str = f.read()
                    meeting_number = parsing.html.common.TdocsByAgendaData.get_meeting_number(
                        cache_tdocsbyagenda_html_str)
                    meeting_data = application.meeting_helper.sa2_meeting_data
                    meeting_text = meeting_data.get_meeting_text_for_given_meeting_number(meeting_number)
                    print(f'Current meeting (10.10.10.10) is {meeting_number}: {meeting_text}')

            print(f'Selecting meeting {meeting_text} and disabling meeting drop-down list')
            tkvar_meeting.set(meeting_text)
            tk_combobox_meetings['state'] = tkinter.DISABLED
        else:
            # Jumping from 3GPP Wifi to normal network
            print(f'(Re-)enabling meeting drop-down list')
            tk_combobox_meetings['state'] = tkinter.NORMAL

    root.after(ms=10000, func=detect_3gpp_network_state)


def change_meeting_dropdown(*args):
    reset_status_labels()
    open_tdocs_by_agenda(open_this_file=False)


tkvar_meeting.trace('w', change_meeting_dropdown)


# Text boxes
def get_text_with_scrollbar(
        row,
        column,
        height=2,
        current_main_frame=main_frame,
        width=50
):
    text = tkinter.scrolledtext.ScrolledText(
        current_main_frame,
        height=height,
        width=width,
        wrap=tkinter.WORD
    )

    text.grid(row=row, column=column, columnspan=2)
    return text


def search_netovate():
    """
    Search the Netovate website for a specific TDoc
    """
    tdoc_id = tkvar_tdoc_id.get()
    netovate_url = 'http://netovate.com/doc-search/?fname={0}'.format(tdoc_id)
    print('Opening {0}'.format(netovate_url))
    os.startfile(netovate_url)


# Downloads the TDocs by Agenda file
def open_tdocs_by_agenda(open_this_file=True) -> parsing.html.common.TdocsByAgendaData | None:
    try:
        (meeting_server_folder, local_file) = get_local_tdocs_by_agenda_filename_for_current_meeting()
        if meeting_server_folder is None:
            return None
    except Exception as e:
        print(f'Could not get local TdocsByAgenda {e}')
        return None

    # Save opened Tdocs by Agenda file to global application
    html = get_tdocs_by_agenda_for_selected_meeting(meeting_server_folder,
                                                    open_tdocs_by_agenda_in_browser=open_this_file)
    print('Retrieved local TDocsByAgenda data for {0}. Parsing TDocs'.format(meeting_server_folder))
    application.meeting_helper.current_tdocs_by_agenda = html_parser.get_tdocs_by_agenda_with_cache(
        html,
        meeting_server_folder=meeting_server_folder)

    utils.local_cache.write_data_and_open_file(html, local_file)
    return application.meeting_helper.current_tdocs_by_agenda


def get_tdocs_by_agenda_for_selected_meeting(
        meeting_folder,
        return_revisions_file=False,
        return_drafts_file=False,
        open_tdocs_by_agenda_in_browser=False):
    return_data = server.tdoc.get_tdocs_by_agenda_for_selected_meeting(
        meeting_folder=meeting_folder,
        use_private_server=tkvar_3gpp_wifi_available.get(),
        open_tdocs_by_agenda_in_browser=open_tdocs_by_agenda_in_browser)

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


def get_local_tdocs_by_agenda_filename_for_current_meeting() -> Tuple[str, str] | Tuple[None, None] | None:
    """
    Gets the current file for the TDocsByAgenda file, as well as the meeting server's folder name
    Returns: A tuple containing (meeting_server_folder, local_file)
    """
    try:
        current_selection = tkvar_meeting.get()
        if (current_selection is None) or (current_selection == ''):
            print('Empty current selection: current meeting not yet selected')
            return None, None
        else:
            print('Get TdocsByAgenda for {0} from local file'.format(current_selection))
        meeting_server_folder = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(
            current_selection)
        local_file = server.tdoc.get_local_tdocs_by_agenda_filename(meeting_server_folder)

        return meeting_server_folder, local_file
    except Exception as e:
        print(f'Could not retrieve local TdocsByAgenda filename for current meeting: {e}')
        traceback.print_exc()
        return None


def current_tdocs_by_agenda_exists():
    try:
        (meeting_server_folder, local_file) = get_local_tdocs_by_agenda_filename_for_current_meeting()
        return os.path.isfile(local_file)
    except:
        return False


def cleanup_tdoc_id_in_entry_box():
    tkvar_tdoc_id.set(tkvar_tdoc_id.get().replace(' ', '').replace('\r', '').replace('\n', '').strip())


# Button to open TDoc
def download_and_open_tdoc(
        tdoc_id_to_override=None,
        cached_tdocs_list=None,
        copy_to_clipboard=False,
        skip_opening=False) -> str | List[str] | None:
    cleanup_tdoc_id_in_entry_box()

    if tdoc_id_to_override is None:
        # Normal flow
        tdoc_id = tkvar_tdoc_id.get()
    else:
        # Used to compare two tdocs
        tdoc_id = tdoc_id_to_override

    # If we are performing a global TDoc search
    if tkvar_search_tdoc.get():
        print(f'Will search for TDoc {tdoc_id}')
        retrieved_files, metadata_list = server.tdoc_search.search_download_and_open_tdoc(tdoc_id)
        if retrieved_files is None:
            not_found_string = 'Not found (' + tdoc_id + ')'
            tkvar_tdoc_download_result.set(not_found_string)
        return retrieved_files

    # Search in meeting
    meeting_folder_name = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(
        tkvar_meeting.get())

    using_private_server = tkvar_3gpp_wifi_available.get()
    if using_private_server:
        additional_folders = ['ftp/SA/SA2/Inbox/']
    else:
        additional_folders = None
    retrieved_files, tdoc_url = server.tdoc.get_tdoc(
        meeting_folder_name=meeting_folder_name,
        tdoc_id=tdoc_id,
        use_private_server=using_private_server,
        return_url=True,
        additional_folders=additional_folders
    )

    if cached_tdocs_list is not None and isinstance(cached_tdocs_list, list):
        if retrieved_files is not None:
            print("Added files to cached TDocs list: {0}".format(retrieved_files))
            cached_tdocs_list.extend(retrieved_files)

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
    except Exception as e:
        print(f'Could not get current TdocsByAgenda: {e}')
        tdoc_status = '<unknown>'
    tkvar_last_tdoc_status.set(tdoc_status)
    if (retrieved_files is None) or (len(retrieved_files) == 0):
        pass
    else:
        if not open_downloaded_tdocs:
            found_string = 'Cached file(s)'
            metadata_list = []

        else:
            opened_files, metadata_list = parsing.word.pywin32.open_files(retrieved_files, return_metadata=True)
            found_string = 'Opened {0} file(s)'.format(opened_files)
        tkvar_last_doc_tdoc.set(tkvar_tdoc_id.get())
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

    tkvar_meeting.set(
        application.meeting_helper.sa2_meeting_data.get_meeting_text_for_given_meeting_number(
            application.meeting_helper.current_tdocs_by_agenda.meeting_number))

    tk_combobox_meetings['values'] = application.meeting_helper.sa2_meeting_data.meeting_names
    tk_combobox_meetings['font'] = font_big

    # Variable-change callbacks
    def set_tdoc_id_full(*args):
        # Sets the label for the download button
        tdoc_id = tkvar_tdoc_id.get()
        if tkvar_search_tdoc.get():
            command_string = 'Search'
        else:
            command_string = 'Open'
        button_label = command_string
        if tdoc.utils.is_sa2_tdoc(tdoc_id):
            button_label += ' ' + tdoc_id
        tkvar_tdoc_id_full.set(button_label)
        if tdoc.utils.is_generic_tdoc(tdoc_id) is not None:
            # Enable button
            open_tdoc_button.configure(state=tkinter.NORMAL)
        else:
            # Disable button
            open_tdoc_button.configure(state=tkinter.DISABLED)

    set_tdoc_id_full()
    tkvar_tdoc_id.trace('w', set_tdoc_id_full)
    tkvar_search_tdoc.trace('w', set_tdoc_id_full)

    # Row: Network configuration and application data update
    current_row = 0
    (ttk.Button(
        main_frame,
        text='Network config',
        command=lambda: gui.network_config.NetworkConfigDialog(
            root,
            favicon,
            on_update_ftp=gui.main_gui.update_ftp_button))
    .grid(
        row=current_row,
        column=0,
        sticky="EW"))
    (ttk.Button(
        main_frame,
        text='Reload meeting info',
        command=lambda: load_application_data(reload_inbox_tdocs_by_agenda=True))
    .grid(
        row=current_row,
        column=1,
        sticky="EW"))

    # Row: Meeting Selector
    current_row += 1
    tk_combobox_meetings.grid(
        row=current_row,
        column=0,
        columnspan=2,
        sticky=tkinter.E + tkinter.W,
        padx=10,
        pady=10)

    update_ftp_button()
    tkinter_checkbutton_3gpp_wifi_available.grid(
        row=current_row,
        column=2)

    # Row: Dropdown menu and meeting info
    current_row += 1

    open_last_agenda_button.grid(
        row=current_row,
        column=0,
        sticky="EW")
    ttk.Button(
        main_frame,
        text='Open TDocs by Agenda',
        command=open_tdocs_by_agenda).grid(
        row=current_row,
        column=1,
        sticky="EW")
    ttk.Button(
        main_frame,
        text="Open server meeting folder",
        command=open_server_meeting_folder).grid(
        row=current_row,
        column=2,
        columnspan=1,
        sticky="EW")

    # Row: Open TDoc
    current_row += 1
    tdoc_entry.grid(
        row=current_row,
        column=0,
        padx=10,
        pady=10)
    open_tdoc_button.grid(
        row=current_row,
        column=1,
        sticky="EW")
    open_tdoc_button.configure(command=download_and_open_tdoc)
    ttk.Checkbutton(
        main_frame,
        text='Search all WGs/meetings',
        variable=tkvar_search_tdoc).grid(
        row=current_row,
        column=2)

    # Row: Tools, TDoc table, Open Netovate
    current_row += 1
    (ttk.Button(
        main_frame,
        text='Tools',
        command=lambda: gui.tools_overview.ToolsDialog(
            gui.main_gui.root,
            gui.main_gui.favicon,
            selected_meeting_fn=gui.main_gui.tkvar_meeting.get))
    .grid(
        row=current_row,
        column=0,
        sticky="EW"))

    tdoc_table_button = ttk.Button(
        main_frame,
        text='Open Tdoc table',
        command=lambda: gui.tdocs_table.TdocsTable(
            favicon=favicon,
            parent_widget=root,
            meeting_name=gui.main_gui.tkvar_meeting.get(),
            retrieve_current_tdocs_by_agenda_fn=lambda: gui.main_gui.open_tdocs_by_agenda(open_this_file=False),
            get_tdocs_by_agenda_for_selected_meeting_fn=gui.main_gui.get_tdocs_by_agenda_for_selected_meeting,
            download_and_open_tdoc_fn=gui.main_gui.download_and_open_tdoc,
            get_current_meeting_name_fn=gui.main_gui.tkvar_meeting.get,
            download_and_open_generic_tdoc_fn=server.tdoc_search.search_download_and_open_tdoc
        ))
    tdoc_table_button.grid(
        row=current_row,
        column=1,
        columnspan=1,
        sticky="EW")

    # Add button to check Netovate (useful if you are searching for documents from other WGs
    (ttk.Button(
        main_frame,
        text='Search Netovate',
        command=search_netovate)
    .grid(
        row=current_row,
        column=2,
        sticky="EW"))

    # Row: Open local folder, open server folder
    current_row += 1
    ttk.Button(
        main_frame,
        text="Open local meeting folder",
        command=open_local_meeting_folder).grid(
        row=current_row,
        column=0,
        columnspan=1,
        sticky="EW")
    ttk.Button(
        main_frame,
        text="Open local specs folder",
        command=lambda: os.startfile(get_specs_folder())).grid(
        row=current_row,
        column=1,
        columnspan=1,
        sticky="EW")
    ttk.Button(
        main_frame,
        text="Close Word",
        command=application.word.close_word).grid(
        row=current_row,
        column=2,
        columnspan=1,
        sticky="EW")

    # Configure <RETURN> key shortcut to open a Tdoc
    gui.common.utils.bind_key_to_button(
        frame=root,
        key_press='<Return>',
        tk_button=open_tdoc_button)

    # Row: Table containing all 3GPP specs
    current_row += 1
    launch_spec_table = ttk.Button(
        main_frame,
        text='Open Specifications table',
        command=lambda: gui.specs_table.SpecsTable(
            root_widget=root,
            parent_widget=root,
            favicon=favicon))
    launch_spec_table.grid(row=current_row, column=0, columnspan=1, sticky="EW")

    # Row: Table containing all 3GPP meetings
    launch_spec_table = ttk.Button(
        main_frame,
        text='Open Meetings table',
        command=lambda: gui.meetings_table.MeetingsTable(
            root_widget=root,
            parent_widget=root,
            favicon=favicon))
    launch_spec_table.grid(row=current_row, column=1, columnspan=1, sticky="EW")

    # Row: Table containing all 3GPP WIs
    launch_spec_table = ttk.Button(
        main_frame,
        text='Open 3GPP WI table',
        command=lambda: gui.work_items_table.WorkItemsTable(
            root_widget=root,
            parent_widget=root,
            favicon=favicon))
    launch_spec_table.grid(row=current_row, column=2, columnspan=1, sticky="EW")

    # Override TDocs by Agenda if it is malformed
    current_row += 1
    (ttk.Checkbutton(
        main_frame,
        text='Override Tdocs by agenda',
        variable=tkvar_override_tdocs_by_agenda)
    .grid(
        row=current_row,
        column=0))
    tdocs_by_agenda_entry.config(state='readonly')
    tdocs_by_agenda_entry.grid(
        row=current_row,
        column=1)

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
        utils.threading.do_something_on_thread(
            task=open_last_agenda(*args),
            before_starting=open_last_agenda_button.config(state='disabled'),
            after_task=open_last_agenda_button.config(state='normal')
        )

    open_last_agenda_button.configure(command=on_open_last_agenda)

    def open_last_agenda(*args):
        try:
            meeting_folder = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(
                tkvar_meeting.get())
            server.tdoc.get_agenda_files(meeting_folder, use_inbox=False)
            last_agenda_info = server.tdoc.get_last_agenda(meeting_folder)
            if last_agenda_info is not None:
                # Starting with SA2#161, there is also a Session Plan (separated from the agenda)
                if last_agenda_info.agenda_path is not None:
                    parsing.word.pywin32.open_file(last_agenda_info.agenda_path)
                    tkvar_last_agenda_version.set('v' + str(last_agenda_info.agenda_version_int))
                if last_agenda_info.session_plan_path is not None:
                    parsing.word.pywin32.open_file(last_agenda_info.session_plan_path)
                    tkvar_last_agenda_version.set('v' + str(last_agenda_info.session_plan_version_int))
            else:
                tkvar_last_agenda_version.set('Not found')
        except Exception as e:
            print(f'Could not open last agenda: {e}')

    # Row: Infos
    current_row += 1
    ttk.Label(
        main_frame,
        textvariable=tkvar_tdoc_download_result).grid(
        row=current_row,
        column=1)
    ttk.Label(
        main_frame,
        textvariable=tkvar_last_agenda_vtext).grid(
        row=current_row,
        column=2)

    # Row: info from last document
    current_row += 1
    tkinter.ttk.Separator(
        main_frame,
        orient=tkinter.HORIZONTAL).grid(
        row=current_row,
        columnspan=3,
        sticky="WE")

    current_row += 1
    ttk.Label(
        main_frame,
        text='Last document:').grid(
        row=current_row,
        column=0)

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
    ttk.Label(
        main_frame,
        text='Title:').grid(
        row=current_row,
        column=0)
    last_tdoc_title = get_text_with_scrollbar(
        current_row,
        1)

    current_row += 1
    ttk.Label(
        main_frame,
        text='Source:').grid(
        row=current_row,
        column=0)
    last_tdoc_source = get_text_with_scrollbar(
        current_row,
        1)

    current_row += 1
    ttk.Label(
        main_frame,
        text='URL:').grid(
        row=current_row,
        column=0)
    last_tdoc_url = get_text_with_scrollbar(
        current_row, 1,
        height=1
    )

    current_row += 1
    ttk.Label(
        main_frame,
        text='Status:').grid(
        row=current_row,
        column=0)
    last_tdoc_status = get_text_with_scrollbar(
        current_row,
        1,
        height=1)

    # Configure column row widths
    main_frame.grid_columnconfigure(0, weight=1)
    main_frame.grid_columnconfigure(1, weight=1)
    main_frame.grid_columnconfigure(2, weight=1)

    # Finish by setting periodic checking of the network status
    root.after(ms=1000, func=detect_3gpp_network_state)


# Avoid circular references by setting the TDoc open function at runtime
parsing.word.pywin32.open_tdoc_for_compare_fn = lambda tdoc_id, cached_tdocs_list: gui.main_gui.download_and_open_tdoc(
    tdoc_id_to_override=tdoc_id,
    cached_tdocs_list=cached_tdocs_list,
    skip_opening=True)
