import traceback
import re

from application.outlook import get_outlook
from tdoc.utils import tdoc_regex

# Trimmed-down version of the outlook script without any dependency on the rest of the application.
# Can be called with fewer dependencies

email_approval_regex = re.compile(r'e[-]?mail approval')
emeeting_regex = re.compile(r'.*\[SA2[ ]*#([\d]+(A)?(H)?(-)?[Ee])[ ,]+AI[#]?([\d\.]+)[ ,]+(S2-(S2-)?[\d]+)\][ ]*(.*)')


def get_email_approval_emails(folder, target_folder, tdoc_data, use_tdoc_data=True, email_subject_regex=None,
                              folder_parse_regex=None, remove_non_tdoc_emails=True):
    if tdoc_data is None and use_tdoc_data:
        return []

    def regex_data_check(regex_list, subject_str):
        for regex in regex_list:
            if regex.search(subject_str) is not None:
                return True

        return False

    # Also catch e-meeting emails
    if email_subject_regex is None:
        email_subject_regex = [email_approval_regex, emeeting_regex]
        print('Using Regex {0}'.format(email_subject_regex))

    email_approval_emails = [(mail_item, mail_item.Subject, tdoc_regex.search(mail_item.Subject))
                             for mail_item in folder.Items
                             if regex_data_check(email_subject_regex, mail_item.Subject)]
    if remove_non_tdoc_emails:
        email_approval_emails_for_tdoc = [item for item in email_approval_emails if item[2] is not None]
    else:
        email_approval_emails_for_tdoc = email_approval_emails

    emails_to_move = []
    for mail_item, subject, tdoc_match in email_approval_emails_for_tdoc:
        try:
            folder_name = ''
            if tdoc_match is not None and use_tdoc_data:
                tdoc_number = tdoc_match.group(0)
                tdoc_is_from_this_meeting = (tdoc_number in tdoc_data.tdocs.index)

                if tdoc_is_from_this_meeting and use_tdoc_data:
                    ai = tdoc_data.tdocs.at[tdoc_number, 'AI']
                    work_item = tdoc_data.tdocs.at[tdoc_number, 'Work Item']
                    if (work_item == '') or (work_item is None):
                        folder_name = ai
                    else:
                        folder_name = '{0}, {1}'.format(ai, work_item)
                        # There is always an AI, but not always a work item description
                else:
                    print('Not found in TDocsByAgenda: {0}'.format(tdoc_number))
            elif not use_tdoc_data:
                tdoc_is_from_this_meeting = True
                if folder_parse_regex is not None:
                    print('Matching {0}'.format(mail_item.Subject))
                    ai_match = folder_parse_regex.search(mail_item.Subject)
                if ai_match is not None:
                    folder_name = ai_match.groupdict()
            else:
                tdoc_is_from_this_meeting = False

            if tdoc_is_from_this_meeting:
                emails_to_move.append((mail_item, folder_name))
        except:
            print('Could not move email item')
            traceback.print_exc()
    # To Do add handling and creation of individual foldrs per agenda item
    return emails_to_move


def search_subject_in_all_outlook_items(tdoc_id, new_window=True):
    outlook_instance = get_outlook()
    if outlook_instance is None:
        print('Could not retrieve Outlook instance')
        return None

    try:
        # https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
        olFolderInbox = 6
        default_folder = outlook_instance.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
        if new_window:
            # https://docs.microsoft.com/en-us/office/vba/api/outlook.olfolderdisplaymode
            olFolderDisplayNormal = 0
            new_explorer = outlook_instance.Explorers.Add(default_folder, olFolderDisplayNormal)
        else:
            new_explorer = outlook_instance.ActiveExplorer()

        # Search scope: https://docs.microsoft.com/en-us/office/vba/api/outlook.olsearchscope
        olSearchScopeAllOutlookItems = 2

        # Achtung: it will only work in German localization!
        new_explorer.Search(
            'betreff:"{0}"'.format(tdoc_id),
            olSearchScopeAllOutlookItems)
    except:
        print('Could not search for emails')
        traceback.print_exc()
