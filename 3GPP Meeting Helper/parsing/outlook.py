import collections
import csv
import os.path
import re
import traceback
from time import sleep

import pandas as pd

import gui.main
import parsing.excel as excel_parser
import server.common
import server.tdoc
import tdoc.utils
from parsing.outlook_utils import get_email_approval_emails, email_approval_regex
from application.outlook import get_outlook_inbox, get_folder, sa2_list_from_inbox, sa2_email_approval_from_inbox
from application.meeting_helper import sa2_list_folder_name, sa2_email_approval_folder_name
import application.meeting_helper

attachment_regex = re.compile(r'Attachment.*\.txt')
doc_regex = re.compile(r'.*\.doc[x]?')
attachment_content_regex = re.compile(r'Attachment:[\r\n]\"(.*)\"[\r\n](.*)[\r\n](.*)', re.MULTILINE)

AttachmentFile = collections.namedtuple('AttachmentFile', 'filename url')
RevisionDoc = collections.namedtuple('RevisionDoc',
                                     'time tdoc filename absolute_url sender_name sender_address email_url ai_folder chairman_notes')


def get_attachment_data(text):
    if (text is None) or (text == ''):
        return None
    try:
        match = attachment_content_regex.search(text)
        if match is None:
            return None
        return AttachmentFile(match.group(1), match.group(3))
    except:
        return None


def organize_email_approval_attachments(meeting_name, ai_folders):
    tmp_folder = server.common.get_tmp_folder()
    local_meeting_folder = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(meeting_name)
    download_from_inbox = gui.main.inbox_is_for_this_meeting()
    found_emails_with_chairmans_notes = []
    email_list = []
    checked_tdocs = set()
    last_dot = False
    ai_folders_list = list(ai_folders)
    ai_folders_list.sort(key=lambda folder: folder.Name)
    for folder_idx, ai_folder in enumerate(ai_folders_list):
        ai_folder_name = ai_folder.Name
        if last_dot:
            print()
        last_dot = False
        print('{0} ({1}/{2})'.format(ai_folder_name, folder_idx + 1, len(ai_folders_list)))
        mail_items_with_attachments_tuple = [(f, f.Subject) for f in ai_folder.Items]
        print('{0} emails in folder'.format(len(mail_items_with_attachments_tuple)))
        for mail_item, email_subject in mail_items_with_attachments_tuple:
            try:
                email_date = mail_item.ReceivedTime
                email_year = email_date.year
                email_month = email_date.month
                email_day = email_date.day
                email_hour = email_date.hour
                email_minute = email_date.minute
                email_second = email_date.second
            except:
                email_year = 0
                email_month = 0
                email_day = 0
                email_hour = 0
                email_minute = 0
                email_second = 0
                # Error retrieving email date (known Outlook bug)
                # See https://stackoverflow.com/questions/62169709/python-valueerror-microsecond-must-be-in-0-999999-while-using-win32com
                print('E')
                # traceback.print_exc()

            date_str = '{0:04d}.{1:02d}.{2:02d} {3:02d}{4:02d}{5:02d}'.format(email_year, email_month, email_day,
                                                                              email_hour, email_minute, email_second)
            date_str_excel = '{0:04d}.{1:02d}.{2:02d} {3:02d}:{4:02d}:{5:02d}'.format(email_year, email_month,
                                                                                      email_day, email_hour,
                                                                                      email_minute, email_second)

            # Download original file (not in email approval folder)
            # Only criteria is if there is a Tdoc ID in the subject
            approval_tdoc_match = tdoc.utils.tdoc_regex.search(email_subject)
            if approval_tdoc_match is None:
                print('Could not parse TDoc ID from subject: {0}'.format(email_subject))
                continue

            tdoc_id = approval_tdoc_match.group(0)
            if tdoc_id not in checked_tdocs:
                server.tdoc.get_tdoc(
                    local_meeting_folder,
                    tdoc_id,
                    use_inbox=download_from_inbox,
                    return_url=True,
                    searching_for_a_file=True)
                checked_tdocs.add(tdoc_id)
            else:
                # Avoid repeated calls to get_tdoc()
                pass
            local_folder_for_tdoc = server.tdoc.get_local_folder(local_meeting_folder, tdoc_id, create_dir=True,
                                                            email_approval=True)

            try:
                sender = mail_item.Sender
                sender_name = sender.Name
                sender_address = sender.Address
            except:
                print('Could not read sender address')
                traceback.print_exc()
                sender_name = 'Unknown'
                sender_address = 'unknown@unknown'

            # Tried to convert Asian characters, but since sender.Name IS ALREADY A STRING, it is not possible
            # codepage       = mail_item.InternetCodepage
            # encoding       = internet_codepage_to_character_set(codepage)
            # if encoding is not None:
            #     try:
            #         encoded_sender_name = sender_name.encode(encoding)
            #         decoded_sender_name = encoded_sender_name.decode('utf-8')
            #         sender_name         = decoded_sender_name
            #     except:
            #         print('Could not re-encode sender name {0} with encoding {1}'.format(sender_name, encoding))
            #         traceback.print_exc()

            # Fix sender email address for encrypted emails
            if 'dmarc-request@LIST.ETSI.ORG' in sender_address:
                try:
                    reply_recipients = [recipient.Address for recipient in mail_item.ReplyRecipients]
                    if len(reply_recipients) > 0:
                        sender_address = reply_recipients[0]
                except:
                    print('Error reading sender from encrypted sender')
                    traceback.print_exc()
                    pass

            # Add email to list (regarding of whether it has an attachment or not)
            ai = ai_folder_name.split(',')[0]
            email_data = (date_str_excel, ai, ai_folder_name, email_subject, sender_name, sender_address)
            email_list.append(email_data)

            # Save local copy of email and remove problematic characters
            email_local_copy = '{0} {1}.msg'.format(date_str, sender_address).replace('/', '')
            email_local_copy_path = os.path.join(local_folder_for_tdoc, email_local_copy)
            if not os.path.isfile(email_local_copy_path):
                if last_dot:
                    print()
                print('Saving email to {0}'.format(email_local_copy_path))
                last_dot = False
                mail_item.SaveAs(email_local_copy_path)
            else:
                print(".", end='')
                last_dot = True

            # Check if email contains message body with a revision (SA2-138E eMeeting)
            email_body = mail_item.Body
            start_str = '<<START>>'
            end_str = '<<END>>'
            comment_start = email_body.find(start_str)
            comment_end = email_body.find(end_str)
            start_of_prior_emails = email_body.find('3GPP_TSG_SA_WG2@LIST.ETSI.ORG')

            # We are now using another email list for the eMeetings
            if start_of_prior_emails == -1:
                start_of_prior_emails = email_body.find('3GPP_TSG_SA_WG2_EMEET@LIST.ETSI.ORG')

            # Remove body text from previous emails
            if start_of_prior_emails > -1:
                email_body = email_body[0:start_of_prior_emails]

            # Record only mails with chairman's notes. The rest are not needed
            if comment_start > -1 and comment_end > -1:
                chairman_notes_comment = email_body[(comment_start + len(start_str)):(comment_end - 1)]
                chairman_notes_comment = chairman_notes_comment.replace('\n', '').replace('\r', '').strip()
                found_revisions = re.findall(r'r[\d]{2}', chairman_notes_comment)
                if len(found_revisions) > 0:
                    revisions = ','.join(found_revisions)
                    found_emails_with_chairmans_notes.append(
                        RevisionDoc(date_str_excel, tdoc_id, revisions, '', sender_name, sender_address,
                                    email_local_copy_path, ai_folder.Name, chairman_notes_comment))
                else:
                    found_emails_with_chairmans_notes.append(
                        RevisionDoc(date_str_excel, tdoc_id, '', '', sender_name, sender_address, email_local_copy_path,
                                    ai_folder.Name, chairman_notes_comment))

            # Moved attachments check here so that all emails get indexed by the prior lines
            if mail_item.Attachments.Count < 1:
                continue

            # Download/copy attachments to local folder
            for attachment in mail_item.Attachments:
                try:
                    name = attachment.FileName
                except:
                    print('Could not retrieve attachment name from {0}, skipping'.format(email_subject))
                    continue
                is_txt_attachment = (attachment_regex.match(name) is not None)
                is_doc_attachment = (doc_regex.match(name) is not None)
                if is_txt_attachment:
                    local_tmp_filename = os.path.join(tmp_folder, attachment.FileName)
                    print('  {0}, {1}'.format(mail_item.Subject, name))
                    attachment.SaveAsFile(local_tmp_filename)
                    attachment_data = None
                    with open(local_tmp_filename, "r") as file:
                        attachment_content = file.read()
                    os.remove(local_tmp_filename)
                    attachment_data = get_attachment_data(attachment_content)
                    if attachment_data is not None:
                        tdoc_data = tdoc.utils.tdoc_regex.match(attachment_data.filename)
                        if tdoc_data is not None:
                            tdoc_id = tdoc_data.group(0)
                            filename_for_file = '{0} {1}'.format(date_str, attachment_data.filename)
                            attachment_local_filename = os.path.join(local_folder_for_tdoc, filename_for_file)
                            print('  TDOC {0}, {1}'.format(tdoc_id, filename_for_file))
                            if not os.path.isfile(attachment_local_filename):
                                server.common.download_file_to_location(attachment_data.url, attachment_local_filename)
                            found_emails_with_chairmans_notes.append(
                                RevisionDoc(date_str_excel, tdoc_id, attachment_data.filename,
                                            attachment_local_filename, sender_name, sender_address,
                                            email_local_copy_path, ai_folder.Name, ''))
                elif is_doc_attachment:
                    tdoc_data = tdoc.utils.tdoc_regex.match(name)
                    if tdoc_data is not None:
                        attachment_local_filename = os.path.join(local_folder_for_tdoc, name)
                        if not os.path.isfile(attachment_local_filename):
                            attachment.SaveAsFile(attachment_local_filename)
                        found_emails_with_chairmans_notes.append(
                            RevisionDoc(date_str_excel, tdoc_id, name, attachment_local_filename, sender_name,
                                        sender_address, email_local_copy_path, ai_folder.Name, ''))

    if last_dot:
        print()
    return found_emails_with_chairmans_notes, email_list


def internet_codepage_to_character_set(codepage):
    if codepage is None:
        return None

    # As listed in https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem.internetcodepage
    switcher = {
        28596: 'iso-8859-6',
        1256: 'windows-1256',
        28594: 'iso-8859-4',
        1257: 'windows-1257',
        28592: 'iso-8859-2',
        1250: 'windows-1250',
        936: 'gb2312',
        52936: 'hz-gb-2312',
        950: 'big5',
        28595: 'iso-8859-5',
        20866: 'koi8-r',
        21866: 'koi8-u',
        1251: 'windows-1251',
        28597: 'iso-8859-7',
        1253: 'windows-1253',
        38598: 'iso-8859-8-i',
        1255: 'windows-1255',
        51932: 'euc-jp',
        50220: 'iso-2022-jp',
        50221: 'csISO2022JP',
        932: 'iso-2022-jp',
        949: 'ks_c_5601-1987',
        51949: 'euc-kr',
        28593: 'iso-8859-3',
        28605: 'iso-8859-15',
        874: 'windows-874',
        28599: 'iso-8859-9',
        1254: 'windows-1254',
        65000: 'utf-7',
        65001: 'utf-8',
        20127: 'us-ascii',
        1258: 'windows-1258',
        28591: 'iso-8859-1',
        1252: 'Windows-1252',
    }

    try:
        return switcher[codepage]
    except:
        print('Could not map InternetCodePage {0} to an encoding'.format(codepage))
        return None


# See https://stackoverflow.com/questions/24321752/outlook-vba-how-to-loop-through-inbox-and-list-from-email-email-address-if-subje
def process_email_approval(meeting_name, generate_summary=True):
    inbox_folder = get_outlook_inbox()
    if inbox_folder is None:
        return

    root_folder = inbox_folder.Parent

    if sa2_list_from_inbox:
        sa2_folder = get_folder(inbox_folder, sa2_list_folder_name, create_if_needed=False)
    else:
        sa2_folder = get_folder(root_folder, sa2_list_folder_name, create_if_needed=False)

    if sa2_folder is None:
        print('Could not find SA2 folder {0}'.format(sa2_list_folder_name))
        return

    if sa2_email_approval_from_inbox:
        sa2_email_approval_folder = get_folder(inbox_folder, sa2_email_approval_folder_name, create_if_needed=True)
    else:
        sa2_email_approval_folder = get_folder(root_folder, sa2_email_approval_folder_name, create_if_needed=True)

    sa2_email_approval_meeting_folder = get_folder(sa2_email_approval_folder, meeting_name, create_if_needed=True)
    if (sa2_folder is None) or (sa2_email_approval_folder is None) or (sa2_email_approval_meeting_folder is None):
        return

    # We will need this to organize the emails
    tdoc_data = gui.main.open_tdocs_by_agenda(open_this_file=False)

    print('Parsing SA2 emails and searching for email approval emails ({0})'.format(meeting_name))
    email_approval_emails = get_email_approval_emails(sa2_folder, sa2_email_approval_meeting_folder, tdoc_data)

    folders = set([e[1] for e in email_approval_emails])
    folder_to_com_object = {}
    for folder in folders:
        folder_to_com_object[folder] = get_folder(sa2_email_approval_meeting_folder, folder)

    for mail_item_tuple in email_approval_emails:
        mail_item = mail_item_tuple[0]
        mail_folder = mail_item_tuple[1]
        try:
            print(mail_item.Subject)
            mail_item.Move(folder_to_com_object[mail_folder])
            sleep(0.1)
        except:
            print('Could not move email item. Maybe a security issue?')
            traceback.print_exc()

    remaining_email_approval_emails = get_email_approval_emails(sa2_folder, sa2_email_approval_meeting_folder,
                                                                tdoc_data)
    print('Finished moving approval emails. Remaining email approval emails: {0} ({1})'.format(
        len(remaining_email_approval_emails), meeting_name))

    if not generate_summary:
        print('Skipping caching and summary of emails (takes less time)')
        return

    print('Organizing email attachments attachments')
    ai_folders = sa2_email_approval_meeting_folder.Folders
    found_attachments, email_list = organize_email_approval_attachments(meeting_name, ai_folders)
    print('Total {0} emails processed'.format(len(email_list)))

    # Check if there are TDocs without emails (potentially automatically approved in eMeeting)
    tdocs_with_emails = set([e.tdoc for e in found_attachments])
    all_tdocs = set(list(tdoc_data.tdocs.index.values))
    tdocs_without_emails = all_tdocs - tdocs_with_emails

    print('{0} TDocs: {1} with emails, {2} without emails'.format(len(all_tdocs), len(tdocs_with_emails),
                                                                  len(tdocs_without_emails)))

    time_now = application.meeting_helper.get_now_time_str()
    file_summary_file = os.path.join('{0} email approval.xlsx'.format(time_now))
    local_meeting_folder = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(meeting_name)
    if len(found_attachments) > 0:
        excel_parser.export_email_approval_list(os.path.join(
            server.common.get_local_agenda_folder(local_meeting_folder), file_summary_file), found_attachments,
            tdocs_without_emails, tdoc_data)
    else:
        print('No file attachments found to export to list. Skipping Excel summary of attachmented files')

    # Save email list for statistics use if so wanted
    df_email_list = pd.DataFrame(email_list,
                                 columns=['date', 'ai', 'folder', 'subject', 'sender name', 'sender address'])
    email_summary_file = os.path.join('{0} email summary.csv'.format(time_now))
    email_summary_file_path = os.path.join(server.common.get_local_agenda_folder(local_meeting_folder),
                                           email_summary_file)
    df_email_list.to_csv(
        email_summary_file_path,
        quoting=csv.QUOTE_NONNUMERIC,
        encoding="utf-8",
        escapechar='\\',
        doublequote=False)

    print('Finished organizing email attachments attachments')


def process_email_attachments():
    inbox_folder = get_outlook_inbox()
    if inbox_folder is None:
        return

    root_folder = inbox_folder.Parent

    if sa2_list_from_inbox:
        sa2_folder = get_folder(inbox_folder, sa2_list_folder_name, create_if_needed=False)
    else:
        sa2_folder = get_folder(root_folder, sa2_list_folder_name, create_if_needed=False)

    if sa2_folder is None:
        print('Could not find SA2 folder {0}'.format(sa2_list_folder_name))
        return

    print('Parsing SA2 emails and searching for email attachments to download')
    tmp_folder = server.common.get_tmp_folder()

    mail_items_with_attachments = [mail_item for mail_item in sa2_folder.Items if
                                   (email_approval_regex.search(mail_item.Subject) is None) and (
                                               mail_item.Attachments.Count > 0)]
    for mail_item in mail_items_with_attachments:
        # Download/copy attachments to local folder
        download_attachments = [a for a in mail_item.Attachments if (attachment_regex.match(a.FileName) is not None)]
        if len(download_attachments) > 0:
            print('{0}'.format(mail_item.Subject))
        email_attachment_files_to_add = []
        attachments_to_delete = []

        # Next email if nothing to do
        if len(download_attachments) < 1:
            continue

        # Process attachments
        for attachment in download_attachments:
            name = attachment.FileName
            tmp_filename = attachment.FileName.replace('Attachment', 'Downloaded')
            local_tmp_filename = os.path.join(tmp_folder, tmp_filename)
            print('  {0}, {1}'.format(mail_item.Subject, name))
            attachment.SaveAsFile(local_tmp_filename)
            attachment_data = None
            with open(local_tmp_filename, "r") as file:
                attachment_content = file.read()
            attachment_data = get_attachment_data(attachment_content)
            email_attachment_files_to_add.append(local_tmp_filename)

            # Download file in attachment data
            if attachment_data is not None:
                filename_for_file = attachment_data.filename
                attachment_local_filename = os.path.join(tmp_folder, filename_for_file)
                print('  {0}'.format(filename_for_file))
                if not os.path.isfile(attachment_local_filename):
                    server.common.download_file_to_location(attachment_data.url, attachment_local_filename)
                email_attachment_files_to_add.append(attachment_local_filename)

            # Mark this attachment for deletion
            attachments_to_delete.append(attachment)

        # Add files to email and remove temporary data
        added_new_attachments = []
        error_when_adding_attachments = False
        for att_local_file in email_attachment_files_to_add:
            try:
                added_new_attachments.append(mail_item.Attachments.Add(att_local_file))
                os.remove(att_local_file)
            except:
                error_when_adding_attachments = True
                print('Error when adding new attachment')
                traceback.print_exc()

        # If there are errors, roll back. If not, delete the original .txt attachments
        if error_when_adding_attachments:
            for added_attachment in added_new_attachments:
                added_attachment.Delete()
        else:
            for attachment in download_attachments:
                attachment.Delete()
            mail_item.Save()

    print('Finished parsing emails and searching for email attachments to download')
