import configparser
import datetime
import traceback

import application.outlook
import config.cache as local_cache_config
import server
from config.word import WordConfig
from parsing.html.common import MeetingData
from parsing.html.tdocs_by_agenda import TdocsByAgendaData

# Read config
config = configparser.ConfigParser()
config.sections()
config.read('config.ini')

sa2_current_meeting_tdoc_data = None
sa2_inbox_tdoc_data = None
sa2_meeting_data: MeetingData | None = None

# Global store of the current TDocsByAgenda data
# No type hint to avoid circular references. It should be ": parsing.html.tdocs_by_agenda.tdocs_by_agenda"
current_tdocs_by_agenda: TdocsByAgendaData | None = None

word_own_reporter_name = None
home_directory = None

# Write config
try:
    server.default_http_proxy = config['HTTP']['DefaultHttpProxy']
    sa2_list_folder_name = config['OUTLOOK']['Sa2MailingListFolder']
    sa2_email_approval_folder_name = config['OUTLOOK']['Sa2EmailApprovalFolder']
except KeyError as e:
    print(f'Could not load configuration file: {e}')
    traceback.print_exc()

    server.default_http_proxy = ""
    sa2_list_folder_name = ""
    sa2_email_approval_folder_name = ""

if len(sa2_list_folder_name) > 0 and sa2_list_folder_name[0] == '/':
    sa2_list_folder_name = sa2_list_folder_name[1:]
    application.outlook.sa2_list_from_inbox = False

if len(sa2_email_approval_folder_name) > 0 and sa2_email_approval_folder_name[0] == '/':
    sa2_email_approval_folder_name = sa2_email_approval_folder_name[1:]
    application.outlook.sa2_email_approval_from_inbox = False

# Write other configuration
try:
    word_own_reporter_name = config['REPORTING']['ContributorName']
    print(f'Using Contributor Name for Word report {word_own_reporter_name}')
except Exception as e:
    print(f'Not using Contributor Name for Word report: {e}')

home_directory = '~'
try:
    home_directory = config['GENERAL']['HomeDirectory']
    print(f'Using Home Directory {home_directory}')
except Exception as e:
    print(f'HomeDirectory not set. Using "{home_directory}": {e}')
finally:
    local_cache_config.CacheConfig.user_folder = home_directory

application_folder = '3GPP_SA2_Meeting_Helper'
try:
    application_folder = config['GENERAL']['ApplicationFolder']
    print(f'Using Application Folder {application_folder}')
except Exception as e:
    print(f'ApplicationFolder not set. Using "{application_folder}": {e}')
finally:
    local_cache_config.CacheConfig.root_folder = application_folder

WordConfig.sensitivity_level_label_id = None
WordConfig.sensitivity_level_label_name = None
WordConfig.save_document_after_setting_sensitivity_label = False
try:
    WordConfig.sensitivity_level_label_id = config['WORD']['SensitivityLevelLabelId']
    print(f'Set Word Sensitivity Label ID to {WordConfig.sensitivity_level_label_id}')
except Exception as e:
    print(f'Set Word Sensitivity Label ID not set. Using "{WordConfig.sensitivity_level_label_id}": {e}')
try:
    WordConfig.sensitivity_level_label_name = config['WORD']['SensitivityLevelLabelName']
    print(f'Set Word Sensitivity Label name to {WordConfig.sensitivity_level_label_name}')
except Exception as e:
    print(f'Set Word Sensitivity Label name not set. Using "{WordConfig.sensitivity_level_label_name}": {e}')
try:
    WordConfig.save_document_after_setting_sensitivity_label = config['WORD']['SaveDocumentAfterSettingSensitivityLabel'].lower() in ("yes", "true")
    print(f'Word will save document after setting sensitivity level '
          f'{WordConfig.save_document_after_setting_sensitivity_label}')
except Exception as e:
    print(f'Saving after setting sensitivity level not set. Using "'
          f'{WordConfig.save_document_after_setting_sensitivity_label}": {e}')

print('Loaded configuration file')


def get_now_time_str():
    current_dt = datetime.datetime.now()
    current_dt_str = '{0:04d}.{1:02d}.{2:02d} {3:02d}{4:02d}{5:02d}'.format(current_dt.year, current_dt.month,
                                                                            current_dt.day, current_dt.hour,
                                                                            current_dt.minute, current_dt.second)
    return current_dt_str



