import gui.common.utils
import gui.network_config
import gui.main_gui

# GUI init
gui.common.utils.fix_blurry_fonts_in_windows_10()
gui.common.utils.set_font_size()

# Launch proxy config window and wait until it is closed
proxy_dialog = gui.network_config.NetworkConfigDialog(gui.main_gui.root, gui.main_gui.favicon)
waiting_for_proxy_label = gui.main_gui.set_waiting_for_proxy_message()
gui.main_gui.root.wait_window(proxy_dialog.top)
waiting_for_proxy_label.grid_forget()

gui.main_gui.start_main_gui()
gui.main_gui.root.mainloop()
