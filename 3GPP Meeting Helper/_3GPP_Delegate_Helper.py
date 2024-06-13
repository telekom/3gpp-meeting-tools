import tkinter
from tkinter import ttk

import application.tkinter_config
import config.networking
import gui.common.utils
import gui.meetings_table
import gui.network_config
import gui.specs_table
import gui.work_items_table

from config.networking import NetworkingConfig
from gui.common.common_elements import tkvar_3gpp_wifi_available
from server.network import detect_3gpp_network_state

# GUI init
tk_root = application.tkinter_config.root
tk_root.title("3GPP Delegate helper")
tk_root.iconbitmap(gui.common.utils.favicon)

gui.common.utils.fix_blurry_fonts_in_windows_10()
gui.common.utils.set_font_size()

main_frame = application.tkinter_config.main_frame
main_frame.grid(column=0, row=0, sticky=''.join([tkinter.N, tkinter.W, tkinter.E, tkinter.S]))

# Launch proxy config window and wait until it is closed
proxy_dialog = gui.network_config.NetworkConfigDialog(tk_root, gui.common.utils.favicon)
waiting_for_proxy_label = gui.common.utils.set_waiting_for_proxy_message(main_frame)
tk_root.wait_window(proxy_dialog.top)
waiting_for_proxy_label.grid_forget()

# Row 1: Table containing all 3GPP specs
launch_spec_table = ttk.Button(
    main_frame,
    text='Open Specifications table',
    command=lambda: gui.specs_table.SpecsTable(tk_root, gui.common.utils.favicon, None))
launch_spec_table.grid(
    row=0,
    column=0,
    sticky="EW")

# Row 1: Table containing all 3GPP meetings
launch_spec_table = ttk.Button(
    main_frame,
    text='Open Meetings table',
    command=lambda: gui.meetings_table.MeetingsTable(
        root_widget=tk_root,
        favicon=gui.common.utils.favicon,
        parent_widget=None))
launch_spec_table.grid(
    row=0,
    column=1,
    sticky="EW")

# Row 1: Table containing all 3GPP WIs
launch_spec_table = ttk.Button(
    main_frame,
    text='Open 3GPP WI table',
    command=lambda: gui.work_items_table.WorkItemsTable(tk_root, gui.common.utils.favicon, None))
launch_spec_table.grid(
    row=0,
    column=2,
    sticky="EW")

# Row 2: 3GPP Wi-fi status
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

tk_root.after(
    ms=NetworkingConfig.network_check_interval_ms,
    func=lambda: detect_3gpp_network_state(
        tk_root,
        loop=True,
        interval_ms=NetworkingConfig.network_check_interval_ms))
tk_root.mainloop()
