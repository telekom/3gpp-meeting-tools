import parsing.outlook as outlook_parser
import config.outlook_regex_matching as email_regex
import re
from time import sleep
import traceback

#########################################################
# Configuration variables: where to find the emails
#########################################################

# Filter messages for this meeting (will be output folder name AND used for parsing email subjects)
meeting_name_filter = r'SA3#102bis-e'

# Where to find the emails
email_folder = 'Standardisierung/3GPP/SA3'

# Where to move the emails
target_folder = 'Standardisierung/3GPP/SA3/meetings'

# The regex used to parse the email subjects
meeting_regex = email_regex.sa3

# Change this if your origin folder hangs from the root folder and not the inbox
inbox_folder = outlook_parser.get_outlook_inbox()
root_folder = inbox_folder # inbox_folder.Parent

#########################################################
# Code
#########################################################

email_parsing_regex = [re.compile(e) for e in [meeting_regex, meeting_name_filter]]

# First step: get the object references to the emails to be moved
source_folder = outlook_parser.get_folder(root_folder, email_folder, create_if_needed=True)
target_folder = outlook_parser.get_folder(root_folder, target_folder, create_if_needed=True)
target_folder = outlook_parser.get_folder(target_folder, meeting_name_filter, create_if_needed=True)
emails_to_move = outlook_parser.get_email_approval_emails(
    source_folder,
    target_folder,
    tdoc_data=None,
    use_tdoc_data=False,
    email_subject_regex=email_parsing_regex,
    folder_parse_regex=re.compile(meeting_regex),
    remove_non_tdoc_emails=False)

# Create folders where to place the emails. The named group 'ai' from the regex match is used as sub-folder name
folders = set([e[1]['ai'] for e in emails_to_move])
folder_to_com_object = {}
for folder in folders:
    folder_to_com_object[folder] = outlook_parser.get_folder(target_folder, folder)

# Move emails
for mail_item_tuple in emails_to_move:
    mail_item   = mail_item_tuple[0]
    mail_folder = mail_item_tuple[1]['ai']
    try:
        print(mail_item.Subject)
        mail_item.Move(folder_to_com_object[mail_folder])
        sleep(0.1)
    except:
        print('Could not move email item. Maybe a security issue?')
        traceback.print_exc()
