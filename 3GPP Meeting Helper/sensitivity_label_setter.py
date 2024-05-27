# This script allows you to change the sesitivity label of Word documents in a folder
# It basically opens each document  in Word, sets the sensitivity label to "OFFEN" and then closes it
# Quick instructions:
#   - In order to run this script, you will need to run "pip install pywin32" before the first run
#   - pip needs to run WHILE DIRECTLY CONNECTED TO THE INTERNET (unless you instruct it to use a proxy)
#   - Run the program by running "python sensitivity_label_setter" (you could just create a shortcut)

import os.path
import tkinter
import traceback
import uuid
import win32com.client


def get_word():
    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except:
        try:
            word = win32com.client.Dispatch("Word.Application")
        except:
            word = None
    if word is not None:
        try:
            word.Visible = True
        except:
            pass
    return word


def set_sensitivity_label(
        document,
        label_id='55339bf0-f345-473a-9ec8-6ca7c8197055',
        label_name='OFFEN',
        set_date='2022-08-18T09:04:30Z'
):
    # Added code in case sensitivity labels are required. Meeting documents are public
    did_something = False
    try:
        sensitivity_label = document.SensitivityLabel.GetLabel()
        # See https://docs.microsoft.com/en-us/office/vba/api/overview/library-reference/labelinfo-members-office
        if sensitivity_label.AssignmentMethod == -1:
            # Need to set default sensitivity label. This part may vary depending on your organization
            did_something = True
            new_sl = document.SensitivityLabel.CreateLabelInfo()
            new_sl.ActionId = str(uuid.uuid1())
            new_sl.AssignmentMethod = 1
            new_sl.ContentBits = 0
            new_sl.IsEnabled = True
            new_sl.Justification = ''
            new_sl.LabelId = label_id
            new_sl.LabelName = label_name
            new_sl.SetDate = set_date
            new_sl.SiteId = str(uuid.uuid1())
            document.SensitivityLabel.SetLabel(new_sl, new_sl)
            print('  Set SensitivityLabel to {0}'.format(new_sl.LabelName))
        else:
            print('  Not setting sensitivity label (already set)')
    except:
        print('  Could not get sensitivity label info. Probably this feature is not used by your installation')
        if did_something:
            traceback.print_exc()

    return document


def set_label_to_docs_list(*args):
    folder = tkvar_folder_name.get()
    if folder is None or not os.path.exists(folder) or not os.path.isdir(folder):
        print('"{0}" is not a valid folder'.format(folder))
        return

    doc_files = [f for f in os.listdir(folder) if (f.endswith('doc') or f.endswith('docx')) and not f.startswith('~')]
    files_to_open = [os.path.join(folder, f) for f in doc_files]

    word_instance = get_word()
    for doc_file in files_to_open:
        print('Opening {0}'.format(doc_file))
        doc = word_instance.Documents.Open(doc_file)
        doc = set_sensitivity_label(doc)
        doc.Save()
        print('  Saved {0}'.format(doc_file))
        doc.Close()


# Initialize GUI
root = tkinter.Tk()
root.title("Sensitivity Label Setter")
root.bind('<Return>', set_label_to_docs_list)  # Bind the enter key in this frame to same function as pressing button

main_frame = tkinter.Frame(root)
main_frame.grid(column=0, row=0, sticky=(tkinter.N, tkinter.W, tkinter.E, tkinter.S))

tkvar_folder_name = tkinter.StringVar(root)

open_tdoc_button = ttk.Button(main_frame, text='Set Sensitivity Label', command=set_label_to_docs_list)
open_tdoc_button.grid(row=0, column=1, padx=10, pady=10, sticky="EW", )
folder_entry = ttk.Entry(main_frame, textvariable=tkvar_folder_name, width=75, font='TkDefaultFont')
folder_entry.grid(row=0, column=2, padx=10, pady=10)

root.mainloop()
