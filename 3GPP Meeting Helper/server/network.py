import tkinter

import application.meeting_helper
import parsing.html
import server.common
import utils.local_cache
from application.tkinter_config import tkvar_3gpp_wifi_available, tkvar_meeting, tk_combobox_meetings


def detect_3gpp_network_state(
        root: tkinter.Tk,
        loop=True,
        interval_ms=10000,
):
    """
    Network detection loop
    Args:
        root: The Tkinter root
        loop: Whether to loop
        interval_ms: How often to loop
    """
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
        cache_tdocsbyagenda_path = utils.local_cache.get_private_server_tdocs_by_agenda_local_cache()
        tdocsbyagenda_url = server.common.tdocs_by_agenda_for_checking_meeting_number_in_meeting

        if new_state:
            # Jumping from normal network to 3GPP Wi-fi
            # Download from 10.10.10.10 TdocsByAgenda and check meeting number
            # Then freeze the meeting choice to the one found

            meeting_text = None
            if server.common.download_file_to_location(tdocsbyagenda_url, cache_tdocsbyagenda_path):
                if utils.local_cache.file_exists(cache_tdocsbyagenda_path):
                    with open(cache_tdocsbyagenda_path, "r") as f:
                        cache_tdocsbyagenda_html_str = f.read()
                    meeting_number = parsing.html.tdocs_by_agenda.TdocsByAgendaData.get_meeting_number(
                        cache_tdocsbyagenda_html_str)
                    meeting_data = application.meeting_helper.sa2_meeting_data
                    meeting_text = meeting_data.get_meeting_text_for_given_meeting_number(meeting_number)
                    print(f'Current meeting (10.10.10.10) is {meeting_number}: {meeting_text}')

            print(f'Selecting meeting {meeting_text} and disabling meeting drop-down list')
            if tkvar_meeting.get() != meeting_text:
                # Trigger change only if necessary
                tkvar_meeting.set(meeting_text)
            tk_combobox_meetings['state'] = tkinter.DISABLED
        else:
            # Jumping from 3GPP Wi-fi to normal network
            print(f'(Re-)enabling meeting drop-down list')
            tk_combobox_meetings['state'] = tkinter.NORMAL

    if loop:
        root.after(ms=interval_ms, func=lambda: detect_3gpp_network_state(root))


