import tkinter

import application.meeting_helper
import application.tkinter_config
import parsing.html
import server.common
import utils.local_cache
import gui.common.common_elements


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

    previous_state = gui.common.common_elements.tkvar_3gpp_wifi_available.get()
    new_state = server.common.we_are_in_meeting_network()
    if new_state:
        gui.common.common_elements.tkvar_3gpp_wifi_available.set(True)
    else:
        gui.common.common_elements.tkvar_3gpp_wifi_available.set(False)

    if new_state != previous_state:
        print(f'Changed 3GPP network state from {previous_state} to {new_state}')
        application.meeting_helper.last_known_3gpp_network_status = new_state

        # No need to do this if we do not have SA2-specific logic
        if (application.tkinter_config.main_frame is not None and
                (gui.common.common_elements.tk_combobox_meetings is not None)):
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

                print(f'Selecting meeting {meeting_text}')
                if gui.common.common_elements.tkvar_meeting.get() != meeting_text:
                    # Trigger change only if necessary
                    gui.common.common_elements.tkvar_meeting.set(meeting_text)
                # gui.common.common_elements.tk_combobox_meetings['state'] = tkinter.DISABLED
            else:
                # Jumping from 3GPP Wi-fi to normal network
                print(f'(Re-)enabling meeting drop-down list')
                gui.common.common_elements.tk_combobox_meetings['state'] = tkinter.NORMAL

    if loop:
        root.after(ms=interval_ms, func=lambda: detect_3gpp_network_state(root))
