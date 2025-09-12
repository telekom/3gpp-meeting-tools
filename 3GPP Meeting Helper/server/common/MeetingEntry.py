import datetime
import os
import re
import traceback
from dataclasses import dataclass
from functools import cached_property
from typing import List
from urllib.parse import urlparse, parse_qs

import pandas as pd
from pandas import DataFrame

import tdoc.utils
from application.excel_openpyxl import parse_tdoc_3gu_list_for_wis
from server.common.server_utils import ServerType, DocumentType, TdocType, WorkingGroup, host_public_server
from server.common.server_utils import meeting_id_regex, get_document_or_folder_url, host_private_server
from tdoc.utils import GenericTdoc
from utils.caching.common import hash_file, retrieve_pickle_cache_for_file, store_pickle_cache_for_file
from utils.local_cache import file_exists, get_work_items_cache_folder
from utils.local_cache import get_cache_folder, create_folder_if_needed


@dataclass(frozen=True)
class MeetingEntry:
    meeting_group: str
    meeting_number: str
    meeting_url_3gu: str
    meeting_name: str
    meeting_location: str
    meeting_url_invitation: str
    start_date: datetime.datetime
    meeting_url_agenda: str
    end_date: datetime.datetime
    meeting_url_report: str
    tdoc_start: tdoc.utils.GenericTdoc | None
    tdoc_end: tdoc.utils.GenericTdoc | None
    meeting_url_docs: str
    meeting_folder_url: str

    @property
    def meeting_folder(self) -> str | None:
        """
        The remote meeting folder name in the 3GPP server's group directory based on the meeting_folder URL
        Returns: The remote folder of the meeting in the 3GPP server. If the folder URL is not set, it may return None

        """
        folder_url = self.meeting_folder_url
        if folder_url is None or folder_url == '':
            return folder_url
        split_folder_url = [f for f in folder_url.split('/') if f != '']
        return split_folder_url[-1]

    @cached_property
    def meeting_id(self) -> str | None:
        """
        Parses the meeting ID from the Meeting's URL. This ID is used in 3GU to identify the meeting, e.g.
        https://portal.3gpp.org/Home.aspx#/meeting?MtgId=60623 -> 60623
        Returns: The ID of the meeting. None if the ID could not be parsed

        """
        if self.meeting_url_3gu is None:
            return None

        id_match = meeting_id_regex.match(self.meeting_url_3gu)

        if id_match is None:
            return None

        return id_match.group('meeting_id')

    @cached_property
    def meeting_calendar_ics_url(self) -> str | None:
        """
        Generates a URL for the 3GPP server containing the calendar entry in ICS format
        Returns: The URL of the ICS file

        """
        the_meeting_id = self.meeting_id
        if the_meeting_id is None:
            return None
        return f"https://portal.3gpp.org/webservices/Rest/Meetings.svc/GetiCal/{the_meeting_id}.ics"

    @cached_property
    def meeting_tdoc_list_url(self) -> str | None:
        """
        Returns, based on the meeting ID, the TDoc list URL from the 3GPP portal
        Returns: The URL, None if the meeting ID is not available/parseable
        """
        meeting_id = self.meeting_id
        if meeting_id is None:
            return None

        # e.g. https://portal.3gpp.org/ngppapp/TdocList.aspx?meetingId=60394
        return 'https://portal.3gpp.org/ngppapp/TdocList.aspx?meetingId=' + meeting_id

    @cached_property
    def meeting_tdoc_list_excel_url(self) -> str | None:
        """
        Returns, based on the meeting ID, the TDoc list URL for the Excel file from the 3GPP portal
        Returns: The URL, None if the meeting ID is not available/parseable
        """
        meeting_id = self.meeting_id
        if meeting_id is None:
            return None

        # e.g. https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=60394
        return 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=' + meeting_id

    def get_tdoc_url(self, tdoc_to_get: tdoc.utils.GenericTdoc | str) -> str:
        """
        For a string containing a potential TDoc, returns a URL concatenating the Docs folder and the input TDoc and
        adds a .'zip' extension.
        Args:
            tdoc_to_get: A TDoc ID. Either an object (GenericTdoc) or string. Note that the input is NOT checked!

        Returns: A URL

        """
        if isinstance(tdoc_to_get, tdoc.utils.GenericTdoc):
            tdoc_file = tdoc_to_get.__str__() + '.zip'
        else:
            tdoc_file = tdoc_to_get + '.zip'
        return self.meeting_url_docs + tdoc_file

    def get_tdoc_inbox_url(self, tdoc_to_get: tdoc.utils.GenericTdoc | str):
        """
        For a string containing a potential TDoc, returns a URL concatenating the Inbox folder and the input TDoc and
        adds a .'zip' extension.
        Args:
            tdoc_to_get: A TDoc ID. Either an object (GenericTdoc) or string. Note that the input is NOT checked!

        Returns: A URL

        """
        docs_url = self.get_tdoc_url(tdoc_to_get)
        try:
            inbox_url = re.sub('Docs', 'Inbox', string=docs_url, flags=re.IGNORECASE)
            return inbox_url
        except Exception as e:
            print(f'Could not generate inbox URL, returning Docs URL: {e}')
            return docs_url

    @cached_property
    def local_folder_path(self) -> str | None:
        """
        For a given meeting, returns the cache folder and creates it if it does not exist
        Returns:

        """
        folder_name = self.meeting_folder
        if folder_name is None:
            return None
        full_path = os.path.join(get_cache_folder(), folder_name)
        return full_path

    @property
    def local_agenda_folder_path(self) -> str:
        """
        For a given meeting, returns the cache folder located at meeting_folder/Agenda and creates
        it if it does not exist
        Returns:

        """
        full_path = os.path.join(self.local_folder_path, 'Agenda')
        create_folder_if_needed(full_path, create_dir=True)
        return full_path

    @property
    def local_export_folder_path(self) -> str:
        """
        For a given meeting, returns the cache folder located at meeting_folder/Export and creates
        it if it does not exist
        Returns:

        """
        full_path = os.path.join(self.local_folder_path, 'Export')
        create_folder_if_needed(full_path, create_dir=True)
        return full_path

    @cached_property
    def local_tdoc_list_excel_path(self):
        return os.path.join(self.local_agenda_folder_path, 'TDoc_List.xlsx')

    @cached_property
    def is_li(self):
        return '-LI' in self.meeting_number

    @cached_property
    def meeting_folders_3gpp_wifi_url(self) -> List[str]:
        wg = WorkingGroup.from_string(self.meeting_group)
        candidate_folders = get_document_or_folder_url(
            server_type=ServerType.PRIVATE,
            document_type=DocumentType.TDOC,
            meeting_folder_in_server='',
            tdoc_type=TdocType.NORMAL,
            working_group=wg
        )
        return candidate_folders

    @cached_property
    def working_group_enum(self) -> WorkingGroup:
        return WorkingGroup.from_string(self.meeting_group)

    def get_tdoc_3gpp_wifi_url(self, tdoc_id_str: str) -> List[str]:
        candidate_folders = self.meeting_folders_3gpp_wifi_url
        candidate_urls = [f'{f}{tdoc_id_str}.zip' for f in candidate_folders]
        return candidate_urls

    @property
    def meeting_is_now(self) -> bool:
        if self.start_date is None or self.end_date is None:
            return False

        # Add some time delta
        days_delta = datetime.timedelta(days=3)
        if self.start_date - days_delta < datetime.datetime.now() < self.end_date + days_delta:
            return True
        return False

    @cached_property
    def local_server_url(self):
        return f'{host_private_server}/{self.working_group_enum.get_wg_folder_name(ServerType.PRIVATE)}'

    @cached_property
    def sync_server_url(self):
        return f'{host_public_server}/{self.working_group_enum.get_wg_folder_name(ServerType.SYNC)}'

    def get_tdoc_local_path(self, tdoc_str: str | GenericTdoc) -> str | None:
        """
        Generates the local path for a given TDoc
        Args:
            tdoc_str: The TDoc for which the local path is queried

        Returns: The TDoc local path. None if it could not be generated, e.g. if the local folder cannot be established.
        """
        local_folder = self.local_folder_path
        if local_folder is None:
            return None

        if isinstance(tdoc_str, tdoc.utils.GenericTdoc):
            tdoc_str = tdoc_str.__str__()

        local_file = os.path.join(
            local_folder,
            str(tdoc_str),
            f'{tdoc_str}.zip')
        local_file.replace(f'{os.path.pathsep}{os.path.pathsep}', f'{os.path.pathsep}')
        return local_file

    @cached_property
    def tdoc_excel_local_path(self) -> str | None:
        download_folder = self.local_agenda_folder_path
        if download_folder is None:
            return None
        return os.path.join(download_folder, f'{self.meeting_name}_TDoc_List.xlsx')

    @property
    def tdoc_excel_exists_in_local_folder(self) -> bool | None:
        local_path = self.tdoc_excel_local_path
        if local_path is None:
            return None
        return file_exists(local_path)

    def starts_in_given_year(self, year: int) -> bool:
        if self.start_date is None:
            return False
        return self.start_date.year == year

    @property
    def tdoc_data_from_excel(self):
        if self.tdoc_excel_local_path is None:
            return None
        tdoc_data = CachedMeetingTdocData.from_excel(
            self.tdoc_excel_local_path,
            meeting=self
        )
        return tdoc_data

    @property
    def tdoc_data_from_excel_with_cache_overwrite(self):
        if self.tdoc_excel_local_path is None:
            return None
        tdoc_data = CachedMeetingTdocData.from_excel(
            self.tdoc_excel_local_path,
            meeting=self,
            overwrite_cache=True
        )
        return tdoc_data


@dataclass(frozen=True)
class WorkItem:
    acronym: str
    url: str

    @cached_property
    def work_item_id(self):
        # e.g. "https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId=1060084"
        return parse_qs(urlparse(self.url).query).get('workitemId', [None])[0]

    @cached_property
    def crs_url(self):
        return f'https://portal.3gpp.org/ChangeRequests.aspx?q=1&workitem={self.work_item_id}'

    @cached_property
    def specs_url(self):
        return f'https://portal.3gpp.org/Specifications.aspx?q=1&WiUid={self.work_item_id}'

    @cached_property
    def local_path(self):
        folder = get_work_items_cache_folder()
        file_name = f'{self.work_item_id}.html'
        return os.path.join(folder, file_name)


@dataclass(frozen=True)
class CachedMeetingTdocData:
    tdocs_df: DataFrame
    wi_hyperlinks: dict[str, str]
    meeting: MeetingEntry
    hash: str

    @cached_property
    def work_items(self) -> List[WorkItem]:
        return list([WorkItem(k, v) for k, v in self.wi_hyperlinks.items()])

    @cached_property
    def version(self):
        return 2

    @staticmethod
    def from_excel(
            tdoc_excel_path: str,
            meeting: MeetingEntry,
            from_cache_if_available=True,
            create_cache_if_not_exists=True,
            overwrite_cache=False
    ):
        excel_hash = hash_file(tdoc_excel_path)
        if from_cache_if_available:
            found_cache = CachedMeetingTdocData.get_cache(tdoc_excel_path, excel_hash)
            if found_cache is not None:
                return found_cache

        tdocs_df: DataFrame = pd.read_excel(
            io=tdoc_excel_path,
            index_col=0)
        wi_hyperlinks = parse_tdoc_3gu_list_for_wis(tdoc_excel_path)

        tdoc_data = CachedMeetingTdocData(
            tdocs_df=tdocs_df,
            wi_hyperlinks=wi_hyperlinks,
            meeting=meeting,
            hash=excel_hash
        )

        if create_cache_if_not_exists:
            tdoc_data.store_cache(tdoc_excel_path, overwrite_cache=overwrite_cache)

        return tdoc_data

    @staticmethod
    def get_cache(tdoc_excel_path: str, excel_hash: str = None):
        try:
            cached_data: CachedMeetingTdocData | None = retrieve_pickle_cache_for_file(
                file_path=tdoc_excel_path,
                file_prefix=TDOCS_3GU_PREFIX,
                file_hash=excel_hash
            )
            print(f'Cache version: {cached_data.version}, {len(cached_data.work_items)} WIs')
            return cached_data
        except Exception as e:
            print(f'Could not load CachedMeetingTdocData {tdoc_excel_path}: {e}')
            traceback.print_exc()
            return None

    def store_cache(self, tdoc_excel_path: str, overwrite_cache=False):
        store_pickle_cache_for_file(
            file_path=tdoc_excel_path,
            file_prefix=TDOCS_3GU_PREFIX,
            file_hash=self.hash,
            data=self,
            overwrite_cache=overwrite_cache
        )


TDOCS_3GU_PREFIX = 'TDocs_3GU'
