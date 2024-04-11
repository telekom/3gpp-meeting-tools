import os

import pyperclip


def open_url_and_copy_to_clipboard(url_to_open: str|None):
    """
    Opens a given URL and copies it to the clipboard
    Args:
        url_to_open: A URL
    """
    pyperclip.copy(url_to_open)
    open_url(url_to_open)
    print('Opened {0} and copied to clipboard'.format(url_to_open))


def open_url(url_to_open: str|None):
    """
    Opens a given URL
    Args:
        url_to_open: A URL
    """
    if url_to_open is None or url_to_open == '':
        return

    os.startfile(url_to_open)
    print('Opened {0} '.format(url_to_open))
