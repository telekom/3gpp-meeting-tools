import traceback

import win32com.client

# Global Outlook instance
outlook = None


def get_outlook():
    global outlook
    if outlook is not None:
        return outlook

    try:
        return win32com.client.Dispatch("Outlook.Application")
    except:
        outlook = None
        traceback.print_exc()
        return None


def get_outlook_inbox():
    try:
        outlook_ = get_outlook()
        if outlook_ is None:
            print('Could not retrieve Outlook instance')
            return None
        mapi_namespace = outlook_.GetNamespace("MAPI")
        # https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
        olFolderInbox = 6
        inbox = mapi_namespace.getDefaultFolder(olFolderInbox)

        return inbox
    except:
        print('Could not retrieve Outlook inbox')
        traceback.print_exc()
        return None


def get_subfolder(root_folder, folder_name):
    try:
        folders = root_folder.Folders
        if (folders is None) or (folder_name is None) or (folder_name == ''):
            return None
        for folder in folders:
            if folder.Name == folder_name:
                return folder
    except:
        return None


def get_folder(root_folder, address, create_if_needed=True):
    if (root_folder is None) or (address is None) or (address == ''):
        return None

    try:
        names = address.split('/')
        name = names[0]
        requested_folder = get_subfolder(root_folder, name)
        if (requested_folder is None) and create_if_needed:
            root_folder.Folders.Add(name)
            print('Created folder {0} under {1}'.format(name, root_folder.Name))
            requested_folder = get_subfolder(root_folder, name)

        # Return recursively the last folder in the chain
        if (len(names) > 1) and (requested_folder is not None):
            subfolders = '/'.join(names[1:])
            requested_folder = get_folder(requested_folder, subfolders, create_if_needed)

        return requested_folder
    except:
        print('Could not create folder')
        traceback.print_exc()
        return None

# Some folders hang from the Outlook root, while others hang from the inbox root
sa2_list_from_inbox = True
sa2_email_approval_from_inbox = True