import application.outlook
import parsing.outlook_utils as outlook_utils
import config.outlook_regex_matching as email_regex
import re
from time import sleep
import traceback

#########################################################
# Configuration variables: where to find the emails
#########################################################

# How to choose sub-folder (named group in the regex)
sub_folder_order = 'ai'

# Filter messages for this meeting (will be output folder name AND used for parsing email subjects)
# meeting_name_filter = r'SA3#102bis-e'
meeting_name_filter = r'SA2#144E'
meeting_folder_name = '144E, Electronic'

# Where to find the emails
# email_folder = 'Standardisierung/3GPP/SA3'
email_folder = 'Standardisierung/3GPP/SA2'

# Where to move the emails
# target_folder = 'Standardisierung/3GPP/SA3/meetings'
target_folder = 'Standardisierung/3GPP/SA2/email approval'

# The regex used to parse the email subjects
# meeting_regex = email_regex.sa3
meeting_regex = email_regex.sa2

# Change this if your origin folder hangs from the root folder and not the inbox
inbox_folder = application.outlook.get_outlook_inbox()
root_folder = inbox_folder  # inbox_folder.Parent

#########################################################
# Code
#########################################################

email_parsing_regex = [re.compile(e) for e in [meeting_regex, meeting_name_filter]]

# First step: get the object references to the emails to be moved
print('**************************')
print('Creating folders')
print('**************************')
source_folder = application.outlook.get_folder(root_folder, email_folder, create_if_needed=True)
target_folder = application.outlook.get_folder(root_folder, target_folder, create_if_needed=True)
target_folder = application.outlook.get_folder(target_folder, meeting_folder_name, create_if_needed=True)

print('**************************')
print('Matching emails')
print('**************************')
emails_to_move = outlook_utils.get_email_approval_emails(
    source_folder,
    target_folder,
    tdoc_data=None,
    use_tdoc_data=False,
    email_subject_regex=email_parsing_regex,
    folder_parse_regex=re.compile(meeting_regex),
    remove_non_tdoc_emails=False)


def get_subfolder_from_match_dict(e):
    """
    Create folders where to place the emails. The named group 'ai' from the regex match is used as sub-folder name
    Args:
        e: An Email object

    Returns:

    """
    email_subject = ''
    try:
        email_subject = e[0].Subject
        match_dict = e[1]
        if sub_folder_order not in match_dict:
            return ''
        output = match_dict[sub_folder_order]
        if output is None or not isinstance(output, str):
            return ''
        return output
    except:
        print('Exception parsing {0}'.format(email_subject))
        traceback.print_exc()
        return ''


folders = set([get_subfolder_from_match_dict(e) for e in emails_to_move])
folder_to_com_object = {}
for folder in folders:
    folder_to_com_object[folder] = application.outlook.get_folder(target_folder, folder)

# Move emails
print('**************************')
print('Moving emails')
print('**************************')
for mail_item_tuple in emails_to_move:
    try:
        mail_item = mail_item_tuple[0]
        print('Email "{0}"'.format(mail_item.Subject))
        if (sub_folder_order not in mail_item_tuple[1]) or (mail_item_tuple[1] is None):
            print('  was not matched to an AI folder'.format(mail_item.Subject))
            continue
        mail_folder = mail_item_tuple[1][sub_folder_order]
        print('  to subfolder')
        mail_item.Move(folder_to_com_object[mail_folder])
        sleep(0.1)
    except:
        print('Could not move email item. Maybe a security issue?')
        traceback.print_exc()
