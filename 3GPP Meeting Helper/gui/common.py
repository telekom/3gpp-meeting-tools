import ctypes
import sys
import tkinter
import tkinter.font
import os

default_font_size = 12

# Set application icon
# https://stackoverflow.com/questions/18537918/set-window-icon
favicon = os.path.join(os.path.dirname(os.path.realpath(__file__)), '..', 'favicon.ico')


def fix_blurry_fonts_in_windows_10():
    # Fix to avoid blurry fonts
    # https://stackoverflow.com/questions/36514158/tkinter-output-blurry-for-icon-and-text-python-2-7
    if 'win' in sys.platform:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except:
            print('Could not set DPI awareness')


def set_font_size(size=default_font_size):
    default_font = tkinter.font.nametofont("TkDefaultFont")
    default_font.configure(size=size)


def set_waiting_for_proxy_message(main_frame):
    label = tkinter.Label(main_frame, text="Please configure proxy")
    label.grid(row=0, column=0)
    return label