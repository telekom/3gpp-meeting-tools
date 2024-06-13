import tkinter
from tkinter import ttk

import config.networking
import server.common
from urllib.parse import urlparse, quote_plus
import traceback

import server.connection


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

        ttk.Button(top, text="Use HTTP proxy and close window", command=self.ok).grid(row=3, column=1, sticky="EW")
        ttk.Button(top, text="No proxy and close window", command=self.ko).grid(row=3, column=2, sticky="EW")

        ttk.Label(top, text="Meeting HTTP server").grid(row=4, column=0)
        self.meeting_server = ttk.Label(
            top,
            text=config.networking.private_server)
        self.meeting_server.grid(row=4, column=1, columnspan=2, sticky="EW")

        # Configure column row widths
        top.grid_columnconfigure(0, weight=1)
        top.grid_columnconfigure(1, weight=1)
        top.grid_columnconfigure(2, weight=1)

    def ok(self):
        # Setup a proxy
        print('Setting up proxy with basic authentication')
        try:
            user_password = ''
            user = self.proxy_user.get().strip()
            if len(user) > 0:
                the_password = self.proxy_password.get()
                the_password = quote_plus(the_password)
                user_password = '{0}:{1}@'.format(user, the_password)
            o = urlparse(self.proxy_server.get())
            print('Using proxy {0}://{1}'.format(o.scheme, o.netloc))
            proxies = {
                'http': '{0}://{2}{1}'.format(o.scheme, o.netloc, user_password),
                'https': '{0}://{2}{1}'.format(o.scheme, o.netloc, user_password)
            }
            server.connection.non_cached_http_session.proxies = proxies
        except:
            print('Could not setup HTTP proxy with Basic Authentication')
            traceback.print_exc()

        self.top.destroy()

    def ko(self):
        # No need to set up a proxy
        server.connection.non_cached_http_session.proxies = None
        self.top.destroy()

