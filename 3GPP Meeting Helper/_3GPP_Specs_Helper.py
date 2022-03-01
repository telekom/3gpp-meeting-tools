import gui.config
import gui.common
import tkinter

# GUI init
gui.common.fix_blurry_fonts_in_windows_10()
gui.common.set_font_size()

tk_root = tkinter.Tk()
tk_root.title("3GPP SA2 Meeting helper")
tk_root.iconbitmap(gui.common.favicon)
main_frame = tkinter.Frame(tk_root)
main_frame.grid(column=0, row=0, sticky=(tkinter.N, tkinter.W, tkinter.E, tkinter.S))

# Lauch proxy config window and wait until it is closed
proxy_dialog = gui.config.NetworkConfigDialog(tk_root, gui.common.favicon)
waiting_for_proxy_label = gui.common.set_waiting_for_proxy_message(main_frame)
tk_root.wait_window(proxy_dialog.top)
waiting_for_proxy_label.grid_forget()

gui.specs_table.SpecsTable(parent=tk_root, favicon=gui.common.favicon, parent_gui_tools=None)

tk_root.mainloop()