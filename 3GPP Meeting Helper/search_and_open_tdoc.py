import os.path

import server.tdoc_search
import sys

script_path: str = __file__
[folder, _] = os.path.split(script_path)


# e.g. python search_and_open_tdoc.py SP-250417

def register_application():
    script_path_for_reg = script_path.replace("\\", "\\\\")
    reg_str = f"""Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\SOFTWARE\Classes\\tdoc]
"URL Protocol"=""
@="3GPP Delegate Helper"

[HKEY_CURRENT_USER\SOFTWARE\Classes\\tdoc\shell]

[HKEY_CURRENT_USER\SOFTWARE\Classes\\tdoc\shell\open]

[HKEY_CURRENT_USER\SOFTWARE\Classes\\tdoc\shell\open\command]
@="python \\\"{script_path_for_reg}\\\" %1"
"""

    print(f'Storing tdoc.reg in {folder}')
    reg_file = os.path.join(folder, 'tdoc.reg')
    with open(reg_file, 'w') as f:
        f.write(reg_str)

    os.startfile(reg_file)


if __name__ == "__main__":
    old_stdout = sys.stdout
    log_file_path = os.path.join(folder, "search_and_open_tdoc.log")
    log_file = open(log_file_path, "w", buffering=1)
    try:
        sys.stdout = log_file

        print(f"""Running on {__file__}""")
        first_arg = sys.argv[1]
        print(f'Arg: {first_arg}')

        match first_arg:
            case '-register':
                register_application()
            case _:
                server.tdoc_search.search_download_and_open_tdoc(first_arg)
                pass
    finally:
        sys.stdout = old_stdout
        log_file.close()
