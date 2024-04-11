import os

import pyperclip


def open_url_and_copy_to_clipboard(url_to_open: str):
    """
    Opens a given URL and copies it to the clipboard
    Args:
        url_to_open: A URL
    """
    pyperclip.copy(url_to_open)
    os.startfile(url_to_open)
    print('Opened {0} and copied to clipboard'.format(url_to_open))
