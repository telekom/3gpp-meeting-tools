import os
import re
import traceback
from enum import Enum
from typing import NamedTuple, List, Tuple
from urllib.parse import urljoin

import application.meeting_helper
import parsing.word.docx
import server.common
import tdoc.utils
import utils.local_cache
from application.zip_files import unzip_files_in_zip_file
from parsing.html import common as html_parser
from server.common import ServerType, get_document_or_folder_url, DocumentType
from server.connection import get_remote_file
from server.tdoc import agenda_docx_regex, agenda_draft_docx_regex, agenda_version_regex, ai_names_cache, \
    get_local_tdocs_by_agenda_filename, get_tdocs_by_agenda_for_selected_meeting, agenda_regex
from utils.local_cache import get_local_agenda_folder


class AgendaType(Enum):
    AGENDA = 1
    SESSION_PLAN = 2


def get_latest_agenda_file(
        agenda_files,
        agenda_type: AgendaType = AgendaType.AGENDA) -> str | None:
    if agenda_files is None:
        return None

    # Remove case when Word temporary documents are stored
    agenda_files = [file for file in agenda_files if agenda_docx_regex.match(file) and ('~$' not in file)]

    # Filter by Agenda or Session plan
    if agenda_type == AgendaType.AGENDA:
        agenda_files = [agenda_file for agenda_file in agenda_files if "Agenda" in agenda_file]
        search_type = "agenda"
    else:
        agenda_files = [agenda_file for agenda_file in agenda_files if "Session" in agenda_file]
        search_type = "session plan"

    if len(agenda_files) == 0:
        return None

    # Parse file data
    draft_agenda_files = [file for file in agenda_files if agenda_draft_docx_regex.match(file)]
    if (len(draft_agenda_files) > 0) and (len(draft_agenda_files) != len(agenda_files)):
        # Non-draft agendas have priority over draft agenda files
        agenda_files = [agenda_file for agenda_file in agenda_files if agenda_file not in draft_agenda_files]

    last_agenda_or_session_plan = max(agenda_files, key=get_agenda_file_version_number)
    print(f'Most recent {search_type} file seems to be {last_agenda_or_session_plan}')
    return last_agenda_or_session_plan


# Parses the order (including TDoc number) of an Agenda file.
# This new format was introduced in SA2#158
agenda_tdoc_regex = re.compile('(Draft_)?' + tdoc.utils.tdoc_regex_str)


def get_agenda_file_version_number(x):
    if len(x) == 0 or x[0] == '~':
        # Sanity check for empty strings and/or temporary files
        return -1
    tdoc_match = agenda_tdoc_regex.match(x)
    if tdoc_match is not None:
        tdoc_year = float(tdoc_match.groupdict()['year'])
        tdoc_id = float(tdoc_match.groupdict()['tdoc_number'])
        tdoc_number = tdoc_year * 100000 + tdoc_id
    else:
        print('Could not parse TDoc number from agenda file {0}'.format(x))
        tdoc_number = 0
    x_without_dashes = x.replace('-', '')
    agenda_match = agenda_version_regex.match(x_without_dashes)
    try:
        agenda_version = agenda_match.groupdict()['version']
    except Exception as e:
        print(f'Could not parse Agenda version number for file {x}: {e}')
        return -1
    print('{0} is agenda/session plan version {1} for tdoc ID {2:.0f}'.format(x, agenda_version, tdoc_number))

    # Support up to 100 agenda versions. It should be OK...
    agenda_version = tdoc_number + float(agenda_version) / 100
    return agenda_version


class AgendaInfo(NamedTuple):
    agenda_path: str | None
    agenda_version_int: int | None
    session_plan_path: str | None
    session_plan_version_int: int | None


def get_last_agenda(meeting_folder):
    agenda_folder = get_local_agenda_folder(meeting_folder)
    agenda_files = [file for file in os.listdir(agenda_folder)]

    last_agenda = get_latest_agenda_file(agenda_files, AgendaType.AGENDA)
    last_session_plan = get_latest_agenda_file(agenda_files, AgendaType.SESSION_PLAN)

    if last_agenda is None:
        return AgendaInfo(
            agenda_path=None,
            agenda_version_int=None,
            session_plan_path=None,
            session_plan_version_int=None)

    agenda_path = os.path.join(agenda_folder, last_agenda)
    try:
        agenda_item_descriptions = parsing.word.docx.import_agenda(agenda_path)
    except Exception as e:
        print(f'Could not parse AI descriptions from agenda: {e}')
        agenda_item_descriptions = {}
    ai_names_cache[meeting_folder] = agenda_item_descriptions

    # Convert agenda version
    agenda_version_str = ''
    agenda_version_int = -1
    try:
        last_agenda_match = agenda_docx_regex.match(last_agenda)
        agenda_version_str = last_agenda_match.groupdict()['version']
        agenda_version_int = int(agenda_version_str)
    except ValueError as e:
        print(f"Could not parse agenda version: {agenda_version_str}. Agenda: {last_agenda}: {e}")
        traceback.print_exc()

    session_plan_path = None
    session_plan_version_int = None
    if last_session_plan is not None:
        session_plan_path = os.path.join(agenda_folder, last_session_plan)
        session_plan_version_str = ''
        try:
            session_plan_match = agenda_docx_regex.match(last_agenda)
            session_plan_version_str = session_plan_match.groupdict()['version']
            session_plan_version_int = int(agenda_version_str)
        except ValueError as e:
            print(
                f"Could not parse session plan version: {session_plan_version_str}. "
                f"Session plan: {last_session_plan}: {e}")
            traceback.print_exc()

    return AgendaInfo(
        agenda_path=agenda_path,
        agenda_version_int=agenda_version_int,
        session_plan_path=session_plan_path,
        session_plan_version_int=session_plan_version_int)


def download_agenda_file(meeting, inbox_active=False, open_tdocs_by_agenda_in_browser=False):
    try:
        print('Downloading Agenda File')
        meeting_server_folder = application.meeting_helper.sa2_meeting_data.get_server_folder_for_meeting_choice(
            meeting)
        local_file = get_local_tdocs_by_agenda_filename(meeting_server_folder)
        html = get_tdocs_by_agenda_for_selected_meeting(meeting_server_folder, inbox_active,
                                                        open_tdocs_by_agenda_in_browser=open_tdocs_by_agenda_in_browser)
        if html is None:
            print('Agenda file for {0} not found'.format(meeting))
            return None
        utils.local_cache.write_data_and_open_file(html, local_file)
        return local_file
    except Exception as e:
        print(f'Could not download agenda file for {meeting}: {e}')
        traceback.print_exc()
        return None


def get_agenda_files(
        meeting_folder_name,
        server_type: ServerType) -> None:
    """
    Retrieves all the agenda (and session plan) files from the agenda folders (both in drafts and non-drafts)
    Args:
        meeting_folder_name: The server folder name for this meeting
        server_type: Whether we are querying the private (3GPP Wi-Fi) or public (www.3gpp.org) server
    """
    server_folders = get_document_or_folder_url(
        server_type=server_type,
        document_type=DocumentType.AGENDA,
        meeting_folder_in_server=meeting_folder_name)

    print(f'Will download all agenda files in {server_folders}')
    agenda_folders = [retrieved_data
                      for retrieved_data
                      in [(server_folder, get_remote_file(server_folder, cache=False))
                          for server_folder in server_folders]
                      if retrieved_data[1] is not None]

    agenda_files: List[Tuple[str, str]] = []
    agenda_local_folder = get_local_agenda_folder(meeting_folder_name)
    for agenda_folder in agenda_folders:
        folder_url = agenda_folder[0]
        html = agenda_folder[1]
        parsed_folder = html_parser.parse_3gpp_http_ftp(html)
        agenda_files_in_folder = [file for file in parsed_folder.files if agenda_regex.match(file)]
        agenda_files_url_local_file_in_folder = [
            (urljoin(folder_url, f), os.path.join(agenda_local_folder, f))
            for f in agenda_files_in_folder]
        agenda_files.extend(agenda_files_url_local_file_in_folder)
    print(f'Will download agenda files: ')
    for agenda_file in agenda_files:
        print(f'  {agenda_file[0]}')

    for agenda_file in agenda_files:
        local_file = agenda_file[1]
        server.common.download_file_to_location(
            url=agenda_file[0],
            local_location=local_file,
            cache=True)
        filename, file_extension = os.path.splitext(local_file)

        if file_extension == '.zip':
            unzip_files_in_zip_file(local_file)
