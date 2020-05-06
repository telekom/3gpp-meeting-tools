import application
import gui.main
import gui.config

# GUI init
gui.main.fix_blurry_fonts_in_windows_10()
gui.main.set_font_size()

# Lauch proxy config window and wait until it is closed
proxy_dialog = gui.config.NetworkConfigDialog(gui.main.root, gui.main.favicon)
waiting_for_proxy_label = gui.main.set_waiting_for_proxy_message()
gui.main.root.wait_window(proxy_dialog.top)
waiting_for_proxy_label.grid_forget()

gui.main.start_main_gui()

gui.main.start_check_current_doc_thread()
gui.main.root.mainloop()