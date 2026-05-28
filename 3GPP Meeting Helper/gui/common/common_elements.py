import tkinter
from tkinter import ttk, font

from application.tkinter_config import root

# Shows whether we are in the 3GPP Wi-Fi or not
tkvar_3gpp_wifi_available = tkinter.BooleanVar(root)
tkvar_3gpp_wifi_available.set(False)

# Shows whether we have detected a VPN
tkvar_3gpp_vpn_detected = tkinter.BooleanVar(root)
tkvar_3gpp_vpn_detected.set(False)

# Needs to be here because it is part of the common networking code
tkvar_meeting = tkinter.StringVar(root)
tk_combobox_meetings: ttk.Combobox | None = None



