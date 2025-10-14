import os
from dataclasses import dataclass
from functools import cached_property

from server.common.MeetingEntry import MeetingEntry
from tdoc.utils import GenericTdoc
from utils.caching.common import export_subfolder


@dataclass(frozen=True)
class Tdoc:
    tdoc: GenericTdoc
    meeting: MeetingEntry

    @cached_property
    def get_3gpp_server_url(self) -> str | None:
        return self.meeting.get_tdoc_url(self.tdoc)

    @cached_property
    def get_local_path(self) -> str | None:
        return self.meeting.get_tdoc_local_path(self.tdoc)

    @cached_property
    def get_local_folder(self) -> str | None:
        return os.path.dirname(self.get_local_path)

    @cached_property
    def get_local_export_path(self) -> str | None:
        return os.path.join(self.get_local_folder, export_subfolder)
