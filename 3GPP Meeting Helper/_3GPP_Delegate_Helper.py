import gui.network_config
import gui.common.utils
import gui.meetings_table
import gui.work_items_table
import gui.specs_table
import tkinter
from tkinter import ttk

# GUI init
tk_root = tkinter.Tk()
tk_root.title("3GPP Delegate helper")
tk_root.iconbitmap(gui.common.utils.favicon)

gui.common.utils.fix_blurry_fonts_in_windows_10()
gui.common.utils.set_font_size()

main_frame = tkinter.Frame(tk_root)
main_frame.grid(column=0, row=0, sticky=''.join([tkinter.N, tkinter.W, tkinter.E, tkinter.S]))

# Launch proxy config window and wait until it is closed
proxy_dialog = gui.network_config.NetworkConfigDialog(tk_root, gui.common.utils.favicon)
waiting_for_proxy_label = gui.common.utils.set_waiting_for_proxy_message(main_frame)
tk_root.wait_window(proxy_dialog.top)
waiting_for_proxy_label.grid_forget()

# Row 4: Table containing all 3GPP specs
launch_spec_table = ttk.Button(
    main_frame,
    text='Open Specifications table',
    command=lambda: gui.specs_table.SpecsTable(tk_root, gui.common.utils.favicon, None))
launch_spec_table.grid(row=0, column=0, columnspan=1, sticky="EW")

# Row 4: Table containing all 3GPP meetings
launch_spec_table = ttk.Button(
    main_frame,
    text='Open Meetings table',
    command=lambda: gui.meetings_table.MeetingsTable(
        root_widget=tk_root,
        favicon=gui.common.utils.favicon,
        parent_widget=None))
launch_spec_table.grid(row=0, column=1, columnspan=1, sticky="EW")

# Row 4: Table containing all 3GPP WIs
launch_spec_table = ttk.Button(
    main_frame,
    text='Open 3GPP WI table',
    command=lambda: gui.work_items_table.WorkItemsTable(tk_root, gui.common.utils.favicon, None))
launch_spec_table.grid(row=0, column=2, columnspan=1, sticky="EW")

tk_root.mainloop()
