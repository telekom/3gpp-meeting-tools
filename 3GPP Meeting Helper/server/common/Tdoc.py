from dataclasses import dataclass
from functools import cached_property

from server.common.MeetingEntry import MeetingEntry
from tdoc.utils import GenericTdoc

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