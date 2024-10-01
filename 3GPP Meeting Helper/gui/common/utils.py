import ctypes
import os
import sys
import tkinter
from tkinter import ttk
import tkinter.font

default_font_size = 12

# Set application icon
# https://stackoverflow.com/questions/18537918/set-window-icon
favicon = os.path.join(os.path.dirname(os.path.realpath(__file__)), '..', '..', 'favicon.ico')


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
    label = ttk.Label(main_frame, text="Please configure proxy")
    label.grid(row=0, column=0)
    return label


def bind_key_to_button(
        frame: tkinter.Tk | tkinter.Toplevel,
        key_press: str,
        tk_button: tkinter.Button,
        check_state=True,
        task_str='TDoc search'
):
    def on_key_button_press(*args):
        print(f'{key_press} pressed')
        try:
            button_status = tk_button['state']
        except Exception as e:
            print(f'Could not get button status: {e}')
            button_status = tkinter.DISABLED

        if check_state:
            print(f'Button state="{button_status}"')
            match button_status:
                case tkinter.NORMAL:
                    print(f'Invoking button {tk_button}')
                    tk_button.invoke()
                case _:
                    print(f'Not invoking button. Not in state "{tkinter.NORMAL}"')
        else:
            print(f'Invoking button without state check: {tk_button}')
            tk_button.invoke()

    # Bind the enter key in this frame to a button press (if the button is active)
    frame.bind(key_press, on_key_button_press)
    print(f'Bound {key_press} key to {task_str}')


def fixed_map(style, style_name, option):
    # See https://bugs.python.org/issue36468
    # Fix for setting text colour for Tkinter 8.6.9
    # From: https://core.tcl.tk/tk/info/509cafafae
    #
    # Returns the style map for 'option' with any styles starting with
    # ('!disabled', '!selected', ...) filtered out.

    # style.map() returns an empty list for missing options, so this
    # should be future-safe.
    return [elm for elm in style.map(style_name, query_opt=option) if
            elm[:2] != ('!disabled', '!selected')]


def get_new_style(style_name: str = None):
    style = ttk.Style()
    if style_name is not None:
        fg = fixed_map(style, style_name, 'foreground')
        bg = fixed_map(style, style_name, 'background')
        print(f'Applying styles for style {style_name}: foreground={fg}, background={bg}')
        style.map(
            style_name,
            foreground=fg,
            background=bg)
    return style
