import os.path

from parsing.html_specs import extract_releases_from_latest_folder
from server.common import get_html
from server.common import root_folder, create_folder_if_needed, decode_string

specs_url = 'https://www.3gpp.org/ftp/Specs/latest'


def get_specs_folder():
    html = get_html(specs_url)
    return html


def get_specs():
    html_latest_specs = get_specs_folder()
    print('Specs folder {0}'.format(html_latest_specs))
    html_latest_specs_bytes = decode_string(html_latest_specs, 'Latest specs')

    print('Retrieved specs HTML')
    print(html_latest_specs_bytes)
    specs = extract_releases_from_latest_folder(html_latest_specs)
    print('Extracted specs')
    print(specs)


def get_specs_folder(create_dir=True):
    folder_name = os.path.expanduser(os.path.join('~', root_folder, 'tmp'))
    create_folder_if_needed(folder_name, create_dir)
    return folder_name
