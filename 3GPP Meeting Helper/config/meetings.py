from datetime import datetime
from server.common import MeetingEntry
from tdoc.utils import GenericTdoc


class MeetingConfig:
    # Meetings not pertaining to a specific WG
    additional_meetings = [
        MeetingEntry(
            meeting_group="SP",
            meeting_number="1",
            meeting_url_3gu="https://portal.3gpp.org/Home.aspx#/meeting?MtgId=60806",
            meeting_name="3GPP WS on 6G",
            meeting_location="Incheon, KR",
            meeting_url_invitation="https://www.3gpp.org/ftp/workshop/2025-03-10_3GPP_6G_WS/Invitation/",
            start_date=datetime.fromisoformat('2025-03-10'),
            meeting_url_agenda="https://ftp.3gpp.org/workshop/2025-03-10_3GPP_6G_WS/Agenda/",
            end_date=datetime.fromisoformat('2025-03-11'),
            meeting_url_report="https://www.3gpp.org/ftp/workshop/2025-03-10_3GPP_6G_WS/Report/",
            tdoc_start=GenericTdoc("6GWS-250001"),
            tdoc_end=GenericTdoc("6GWS-250243"),
            meeting_url_docs="https://www.3gpp.org/ftp/workshop/2025-03-10_3GPP_6G_WS/Docs/",
            meeting_folder_url="https://www.3gpp.org/ftp/workshop/2025-03-10_3GPP_6G_WS/")
    ]
