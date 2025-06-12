from typing import NamedTuple

import pandas as pd
from pandas import DataFrame

from application.excel_openpyxl import parse_tdoc_3gu_list_for_wis
from server.common import MeetingEntry
from utils.caching.common import hash_file, store_pickle_cache_for_file, retrieve_pickle_cache_for_file

TDOCS_3GU_PREFIX = 'TDocs_3GU'

class CachedMeetingTdocData(NamedTuple):
    tdocs_df: DataFrame
    wi_hyperlinks: dict[str, str]
    meeting:MeetingEntry
    hash: str

    @staticmethod
    def from_excel(
            tdoc_excel_path:str,
            meeting:MeetingEntry,
            from_cache_if_available=True,
            create_cache_if_not_exists=True
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
            tdoc_data.store_cache(tdoc_excel_path)

        return tdoc_data


    @staticmethod
    def get_cache(tdoc_excel_path:str, excel_hash:str=None):
        cached_data: CachedMeetingTdocData|None = retrieve_pickle_cache_for_file(
            file_path=tdoc_excel_path,
            file_prefix=TDOCS_3GU_PREFIX,
            file_hash=excel_hash
        )
        return cached_data

    def store_cache(self, tdoc_excel_path:str):
        store_pickle_cache_for_file(
            file_path=tdoc_excel_path,
            file_prefix=TDOCS_3GU_PREFIX,
            file_hash=self.hash,
            data=self
        )
