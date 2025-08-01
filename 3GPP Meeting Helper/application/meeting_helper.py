import configparser
import datetime
import os
import traceback
from typing import NamedTuple, List

import application.outlook
import config.cache as local_cache_config
from config.markdown import MarkdownConfig
from config.meetings import MeetingConfig
from config.word import WordConfig

import config.networking
from parsing.html.common import MeetingData
from parsing.html.tdocs_by_agenda import TdocsByAgendaData

# Read config
config_parser = configparser.ConfigParser()
config_parser.sections()

root_folder = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
config_file = os.path.join(root_folder, 'config.ini')
print(f'Reading config file from {config_file}')
config_parser.read(config_file)

sa2_current_meeting_tdoc_data = None
sa2_inbox_tdoc_data = None
sa2_meeting_data: MeetingData | None = None


class TdocTag(NamedTuple):
    tag: str
    agenda_item: str


# Contains a list of tags to mark TDocs
tdoc_tags: List[TdocTag] = []

# Global store of the current TDocsByAgenda data
# No type hint to avoid circular references. It should be ": parsing.html.tdocs_by_agenda.tdocs_by_agenda"
current_tdocs_by_agenda: TdocsByAgendaData | None = None

company_report_name = None
home_directory = None

# Default Proxy
try:
    config.networking.default_http_proxy = config_parser['HTTP']['DefaultHttpProxy']
    print(f'Set default HTTP(s) proxy to {config.networking.default_http_proxy}')
except KeyError as e:
    print(f'Could not read default HTTP proxy: {e}')
    traceback.print_exc()

# Write config
try:
    sa2_list_folder_name = config_parser['OUTLOOK']['Sa2MailingListFolder']
    sa2_email_approval_folder_name = config_parser['OUTLOOK']['Sa2EmailApprovalFolder']
except KeyError as e:
    print(f'Could not load configuration file: {e}')
    traceback.print_exc()

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
    company_report_name = config_parser['REPORTING']['ContributorName']
    MarkdownConfig.company_name_regex_for_report = company_report_name
    print(f'Using Contributor Name for company report {company_report_name}')
except Exception as e:
    print(f'Not using Contributor Name for company report: {e}')

home_directory = '~'
try:
    home_directory = config_parser['GENERAL']['HomeDirectory']
    print(f'Using Home Directory {home_directory}')
except Exception as e:
    print(f'HomeDirectory not set. Using "{home_directory}": {e}')
finally:
    local_cache_config.CacheConfig.user_folder = home_directory

application_folder = '3GPP_SA2_Meeting_Helper'
try:
    application_folder = config_parser['GENERAL']['ApplicationFolder']
    print(f'Using Application Folder {application_folder}')
except Exception as e:
    print(f'ApplicationFolder not set. Using "{application_folder}": {e}')
finally:
    local_cache_config.CacheConfig.root_folder = application_folder

try:
    open_sa2_drafts_url = config_parser['GUI']['SA2_Drafts_URL']
    print(f'Using SA2 Drafts URL {open_sa2_drafts_url}')
except Exception as e:
    open_sa2_drafts_url = 'https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/SA2/Inbox/Drafts'
    print(f'SA2 Drafts URL not set. Using "{open_sa2_drafts_url}": {e}')

WordConfig.sensitivity_level_label_id = None
WordConfig.sensitivity_level_label_name = None
WordConfig.save_document_after_setting_sensitivity_label = False
try:
    WordConfig.sensitivity_level_label_id = config_parser['WORD']['SensitivityLevelLabelId']
    print(f'Set Word Sensitivity Label ID to {WordConfig.sensitivity_level_label_id}')
except Exception as e:
    print(f'Set Word Sensitivity Label ID not set. Using "{WordConfig.sensitivity_level_label_id}": {e}')
try:
    WordConfig.sensitivity_level_label_name = config_parser['WORD']['SensitivityLevelLabelName']
    print(f'Set Word Sensitivity Label name to {WordConfig.sensitivity_level_label_name}')
except Exception as e:
    print(f'Set Word Sensitivity Label name not set. Using "{WordConfig.sensitivity_level_label_name}": {e}')
try:
    WordConfig.save_document_after_setting_sensitivity_label = config_parser['WORD'][
                                                                   'SaveDocumentAfterSettingSensitivityLabel'].lower() in (
                                                                   "yes", "true")
    print(f'Word will save document after setting sensitivity level '
          f'{WordConfig.save_document_after_setting_sensitivity_label}')
except Exception as e:
    print(f'Saving after setting sensitivity level not set. Using "'
          f'{WordConfig.save_document_after_setting_sensitivity_label}": {e}')

try:
    config.networking.http_user_agent = config_parser['HTTP']['UserAgent']
    print(f'Using HTTP User Agent "{config.networking.http_user_agent}"')
except Exception as e:
    print(f'HTTP User Agent not set. Using "{config.networking.http_user_agent}": {e}')
finally:
    local_cache_config.CacheConfig.root_folder = application_folder

# Load TDoc tags
try:
    tdoc_tags_in_config_file = config_parser['TDOC_TAGS']
except Exception as e:
    tdoc_tags_in_config_file = {}
    print(f'No TDoc tags to load in config section {e}')

for k, v in tdoc_tags_in_config_file.items():
    print(f'Storing TDoc tags: {tdoc_tags_in_config_file}')
    try:
        tag = k
        tag_ais = v.split(',')
        tag_ais = [s.strip() for s in tag_ais]
        for tag_ai in tag_ais:
            if tag is not None and tag != '':
                tdoc_tags.append(TdocTag(tag=tag, agenda_item=tag_ai))
    except Exception as e:
        print(f'Could not process tag {k}:{v}. {e}')
if len(tdoc_tags) > 0:
    print(f'TDoc tags: {tdoc_tags}')

# Load markdown configuration
try:
    tdoc_tags_in_config_file = config_parser['MARKDOWN']['Columns']
    MarkdownConfig.columns_for_3gu_tdoc_export = [e.strip() for e in tdoc_tags_in_config_file.split(',')]
    print(f"Imported configuration for Markdown export of current Excel's TDocs: {MarkdownConfig.columns_for_3gu_tdoc_export}")
except Exception as e:
    print(f'No Markdown configuration to load {e}')

try:
    tdoc_tags_in_config_file = config_parser['MARKDOWN']['ExportColumnsLs']
    MarkdownConfig.columns_for_3gu_tdoc_export_ls = [e.strip() for e in tdoc_tags_in_config_file.split(',')]
    print(f"Imported configuration for Markdown export of current meeting's LS: {MarkdownConfig.columns_for_3gu_tdoc_export_ls}")
except Exception as e:
    print(f'No LS Markdown configuration to load {e}')

try:
    tdoc_tags_in_config_file = config_parser['MARKDOWN']['ExportColumnsPcr']
    MarkdownConfig.columns_for_3gu_tdoc_export_pcr = [e.strip() for e in tdoc_tags_in_config_file.split(',')]
    print(f"Imported configuration for Markdown export of current meeting's pCRs: {MarkdownConfig.columns_for_3gu_tdoc_export_pcr}")
except Exception as e:
    print(f'No pCR Markdown configuration to load {e}')

try:
    tdoc_tags_in_config_file = config_parser['MARKDOWN']['ExportColumnsCr']
    MarkdownConfig.columns_for_3gu_tdoc_export_cr = [e.strip() for e in tdoc_tags_in_config_file.split(',')]
    print(f"Imported configuration for Markdown export of current meeting's CRs: {MarkdownConfig.columns_for_3gu_tdoc_export_cr}")
except Exception as e:
    print(f'No CR Markdown configuration to load {e}')

try:
    tdoc_tags_in_config_file = config_parser['MARKDOWN']['ExportColumnsLsOut']
    MarkdownConfig.columns_for_3gu_tdoc_export_ls_out = [e.strip() for e in tdoc_tags_in_config_file.split(',')]
    print(f"Imported configuration for Markdown export of LS OUTs: {MarkdownConfig.columns_for_3gu_tdoc_export_cr}")
except Exception as e:
    print(f'No CR Markdown configuration to load {e}')

try:
    tdoc_tags_in_config_file = config_parser['MARKDOWN']['ExportColumnsContributor']
    MarkdownConfig.columns_for_3gu_tdoc_export_contributor = [e.strip() for e in tdoc_tags_in_config_file.split(',')]
    print(f"Imported configuration for Markdown export of contributions for {MarkdownConfig.company_name_regex_for_report}: {MarkdownConfig.columns_for_3gu_tdoc_export_cr}")
except Exception as e:
    print(f'No CR Markdown configuration to load {e}')

try:
    meeting_list_group_filter = config_parser['MEETINGS']['MeetingGroupFilter']
    MeetingConfig.meeting_list_group_filter = meeting_list_group_filter
    print(f"Imported configuration for initial filter of meeting list: {MeetingConfig.meeting_list_group_filter}")
except Exception as e:
    print(f'No meeting config filter configuration found: {e}')

print('Loaded configuration file')

last_known_3gpp_network_status = False


def get_now_time_str():
    current_dt = datetime.datetime.now()
    current_dt_str = '{0:04d}.{1:02d}.{2:02d} {3:02d}{4:02d}{5:02d}'.format(current_dt.year, current_dt.month,
                                                                            current_dt.day, current_dt.hour,
                                                                            current_dt.minute, current_dt.second)
    return current_dt_str
