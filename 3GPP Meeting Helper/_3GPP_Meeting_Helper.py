import sv_ttk

import gui.common.utils
import gui.network_config
import gui.main_gui

import platform
import ctypes

# Force Windows to see this script as a distinct application on the taskbar
if platform.system() == 'Windows':
    try:
        myappid = 'telekom.3gpp.meetingtools.helper.1.0'  # Arbitrary unique string identifier
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception as e:
        print(f"Could not set AppUserModelID for taskbar icon: {e}")

# GUI init
gui.common.utils.fix_blurry_fonts_in_windows_10()
gui.common.utils.set_font_size()

# Launch proxy config window and wait until it is closed
proxy_dialog = gui.network_config.NetworkConfigDialog(gui.main_gui.root, gui.main_gui.favicon)
waiting_for_proxy_label = gui.main_gui.set_waiting_for_proxy_message()
gui.main_gui.root.wait_window(proxy_dialog.top)
waiting_for_proxy_label.grid_forget()

gui.main_gui.start_main_gui()

# Apply the modern theme
sv_ttk.set_theme("light")

gui.main_gui.root.mainloop()
