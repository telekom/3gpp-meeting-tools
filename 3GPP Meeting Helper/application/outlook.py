import traceback
import platform

if platform.system() == 'Windows':
    print('Windows System detected. Importing win32.client')
    import win32com.client

# Global Outlook instance does not work (removed)
# outlook = None


def get_outlook():
    if platform.system() != 'Windows':
        return None

    try:
        return win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        print(f'Could not retrieve Outlook instance: {e}')
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
    except Exception as e:
        print(f'Could not retrieve Outlook inbox: {e}')
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
    except Exception as e:
        print(f'Could get Outlook subfolder: {e}')
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
    except Exception as e:
        print(f'Could not create folder: {e}')
        traceback.print_exc()
        return None

# Some folders hang from the Outlook root, while others hang from the inbox root
sa2_list_from_inbox = True
sa2_email_approval_from_inbox = True