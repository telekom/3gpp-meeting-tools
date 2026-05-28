import tkinter
from tkinter import ttk

import config.networking
import server.common.server_utils
from urllib.parse import urlparse, quote_plus
import traceback
import gui.common.utils

import server.common.connection


# Wait for proxy settings
# https://stackoverflow.com/questions/38678415/how-to-wait-for-response-from-modal-window-before-continuing-using-tkinter
class NetworkConfigDialog:

    def __init__(self, parent, favicon, on_update_ftp=None):
        top = self.top = tkinter.Toplevel(parent)
        top.title("HTTP Network Configuration")
        top.iconbitmap(favicon)

        # Setup trigger to update FTP server label on main GUI
        self.on_update_ftp = on_update_ftp

        # Set the window to the front and wait until it is closed
        # https://stackoverflow.com/questions/1892339/how-to-make-a-tkinter-window-jump-to-the-front
        top.attributes("-topmost", True)

        ttk.Label(top, text="HTTP proxy host:port").grid(row=0, column=0)
        self.proxy_server = tkinter.Entry(top)
        self.proxy_server.insert(0, config.networking.default_http_proxy)
        self.proxy_server.grid(row=0, column=1, columnspan=2, sticky="EW")

        ttk.Label(top, text="HTTP proxy user").grid(row=1, column=0)
        self.proxy_user = tkinter.Entry(top)
        self.proxy_user.grid(row=1, column=1, columnspan=2, sticky="EW")

        ttk.Label(top, text="HTTP proxy password").grid(row=2, column=0)
        self.proxy_password = tkinter.Entry(top, show='*')
        self.proxy_password.grid(row=2, column=1, columnspan=2, sticky="EW")

        ttk.Label(top, text="Apply on VPN only").grid(row=3, column=0)
        self.proxy_only_on_vpn = tkinter.BooleanVar(value=False)
        self.proxy_only_on_vpn_checkbox = ttk.Checkbutton(
            top,
            state='enabled',
            variable=self.proxy_only_on_vpn
        )
        self.proxy_only_on_vpn_checkbox.grid(row=3, column=1, columnspan=2, sticky="EW")

        ok_button = ttk.Button(top, text="Press enter to use HTTP proxy and close window", command=self.ok)
        ok_button.grid(row=4, column=1, sticky="EW")
        ttk.Button(top, text="No proxy and close window", command=self.ko).grid(row=4, column=2, sticky="EW")

        gui.common.utils.bind_key_to_button(
            frame=top,
            key_press='<Return>',
            tk_button=ok_button,
            check_state=False,
            task_str='Set HTTP(s) proxy'
        )

        ttk.Label(top, text="Meeting HTTP server").grid(row=5, column=0)
        self.meeting_server = ttk.Label(
            top,
            text=config.networking.private_server)
        self.meeting_server.grid(row=5, column=1, columnspan=2, sticky="EW")

        # Configure column row widths
        top.grid_columnconfigure(0, weight=1)
        top.grid_columnconfigure(1, weight=1)
        top.grid_columnconfigure(2, weight=1)

    def ok(self):
        # Setup a proxy
        print('Setting up proxy with basic authentication')
        try:
            server.common.connection.store_proxy(
                server=self.proxy_server.get(),
                user=self.proxy_user.get(),
                password=self.proxy_password.get())
            server.common.connection.set_http_proxy()
            server.common.connection.use_proxy_only_of_vpn = self.proxy_only_on_vpn.get()
        except Exception as e:
            print(f'Could not setup HTTP proxy with Basic Authentication: {e}')
            traceback.print_exc()

        self.top.destroy()

    def ko(self):
        # No need to set up a proxy
        server.common.connection.clear_http_proxies()
        self.top.destroy()

