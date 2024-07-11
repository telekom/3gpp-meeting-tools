import tkinter
from tkinter import ttk

from application.tkinter_config import root, main_frame, font_big, ttk_style_tbutton_medium
import config.networking
import gui.common.utils
import gui.meetings_table
import gui.network_config
import gui.specs_table
import gui.work_items_table
from application.word import close_word

from config.networking import NetworkingConfig
from gui.common.common_elements import tkvar_3gpp_wifi_available
from server.network import detect_3gpp_network_state
import server.tdoc_search

# GUI init
tk_root = root
tk_root.title("3GPP Delegate helper")
tk_root.iconbitmap(gui.common.utils.favicon)

gui.common.utils.fix_blurry_fonts_in_windows_10()
gui.common.utils.set_font_size()

main_frame = main_frame
main_frame.grid(column=0, row=0, sticky=''.join([tkinter.N, tkinter.W, tkinter.E, tkinter.S]))

# Launch proxy config window and wait until it is closed
proxy_dialog = gui.network_config.NetworkConfigDialog(tk_root, gui.common.utils.favicon)
waiting_for_proxy_label = gui.common.utils.set_waiting_for_proxy_message(main_frame)
tk_root.wait_window(proxy_dialog.top)
waiting_for_proxy_label.grid_forget()

button_width = 25

# Row 1: Table containing all 3GPP specs
launch_spec_table = ttk.Button(
    main_frame,
    text='Specifications',
    width=button_width,
    command=lambda: gui.specs_table.SpecsTable(tk_root, gui.common.utils.favicon, parent_widget=None))
launch_spec_table.grid(
    row=0,
    column=0)

# Row 1: Table containing all 3GPP meetings
launch_spec_table = ttk.Button(
    main_frame,
    text='Meetings',
    width=button_width,
    command=lambda: gui.meetings_table.MeetingsTable(
        root_widget=tk_root,
        favicon=gui.common.utils.favicon,
        parent_widget=None))
launch_spec_table.grid(
    row=0,
    column=1)

# Row 1: Table containing all 3GPP WIs
launch_spec_table = ttk.Button(
    main_frame,
    width=button_width,
    text='3GPP WIs',
    command=lambda: gui.work_items_table.WorkItemsTable(tk_root, gui.common.utils.favicon, root_widget=None))
launch_spec_table.grid(
    row=0,
    column=2)

# Row 2:
(ttk.Button(
    main_frame,
    width=button_width,
    text="Close Word",
    command=close_word)
 .grid(
    row=1,
    column=0
))

# 3GPP Wi-fi status
tkinter_checkbutton_3gpp_wifi_available = ttk.Checkbutton(
    main_frame,
    state='disabled',
    variable=tkvar_3gpp_wifi_available)
tkinter_checkbutton_3gpp_wifi_available.config(text=config.networking.private_server + ' (3GPP Wifi)')
tkinter_checkbutton_3gpp_wifi_available.grid(
    row=1,
    column=2,
    padx=10
)

# Open TDoc
tkvar_tdoc_id = tkinter.StringVar(root)
tkvar_tdoc_id_full = tkinter.StringVar(root)
open_tdoc_button = ttk.Button(
    main_frame,
    width=20,
    text='Search TDoc',
    style=ttk_style_tbutton_medium)
tdoc_entry = tkinter.Entry(
    main_frame,
    width=13,
    textvariable=tkvar_tdoc_id,
    justify='center',
    font=font_big)

tdoc_entry.grid(
    row=2,
    column=0,
    padx=10,
    pady=10)
open_tdoc_button.grid(
    row=2,
    column=1
)


def search_and_open_tdoc():
    tdoc_str = tkvar_tdoc_id.get()
    print(f'Will search for TDoc {tdoc_str}')
    retrieved_files, metadata_list = server.tdoc_search.search_download_and_open_tdoc(tdoc_str)


open_tdoc_button.configure(command=search_and_open_tdoc)
# Configure <RETURN> key shortcut to open a Tdoc
gui.common.utils.bind_key_to_button(
    frame=root,
    key_press='<Return>',
    tk_button=open_tdoc_button,
    check_state=False
)

tk_root.after(
    ms=NetworkingConfig.network_check_interval_ms,
    func=lambda: detect_3gpp_network_state(
        tk_root,
        loop=True,
        interval_ms=NetworkingConfig.network_check_interval_ms))
tk_root.mainloop()
