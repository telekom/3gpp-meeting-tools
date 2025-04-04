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
            meeting_folder_url="https://www.3gpp.org/ftp/workshop/2025-03-10_3GPP_6G_WS/"),
        MeetingEntry(
            meeting_group="RP",
            meeting_number="1",
            meeting_url_3gu="https://portal.3gpp.org/Home.aspx#/meeting?MtgId=60517",
            meeting_name="RAN R19 WS",
            meeting_location="Taipei, TW",
            meeting_url_invitation="https://ftp.3gpp.org/TSG_RAN/TSG_RAN/TSGR_AHs/2023_06_RAN_Rel19_WS/Invitation/",
            start_date=datetime.fromisoformat('2023-06-15'),
            meeting_url_agenda="https://ftp.3gpp.org/TSG_RAN/TSG_RAN/TSGR_AHs/2023_06_RAN_Rel19_WS/Agenda/",
            end_date=datetime.fromisoformat('2023-06-16'),
            meeting_url_report="https://ftp.3gpp.org/TSG_RAN/TSG_RAN/TSGR_AHs/2023_06_RAN_Rel19_WS/Report/",
            tdoc_start=GenericTdoc("RWS-230001"),
            tdoc_end=GenericTdoc("RWS-230491"),
            meeting_url_docs="https://ftp.3gpp.org/TSG_RAN/TSG_RAN/TSGR_AHs/2023_06_RAN_Rel19_WS/Docs/",
            meeting_folder_url="https://ftp.3gpp.org/TSG_RAN/TSG_RAN/TSGR_AHs/2023_06_RAN_Rel19_WS/"),
        MeetingEntry(
            meeting_group="SP",
            meeting_number="1",
            meeting_url_3gu="https://portal.3gpp.org/Home.aspx#/meeting?MtgId=60539",
            meeting_name="SA R19 WS",
            meeting_location="Taipei, TW",
            meeting_url_invitation="https://www.3gpp.org/ftp/tsg_sa/TSG_SA/Workshops/2023-06-13_Rel-19_WorkShop/Invitation/",
            start_date=datetime.fromisoformat('2023-06-13'),
            meeting_url_agenda="https://www.3gpp.org/ftp/tsg_sa/TSG_SA/Workshops/2023-06-13_Rel-19_WorkShop/Agenda/",
            end_date=datetime.fromisoformat('2023-06-14'),
            meeting_url_report="https://www.3gpp.org/ftp/tsg_sa/TSG_SA/Workshops/2023-06-13_Rel-19_WorkShop/Report/",
            tdoc_start=GenericTdoc("SWS-230001"),
            tdoc_end=GenericTdoc("SWS-230088"),
            meeting_url_docs="https://www.3gpp.org/ftp/tsg_sa/TSG_SA/Workshops/2023-06-13_Rel-19_WorkShop/Docs/",
            meeting_folder_url="https://www.3gpp.org/ftp/tsg_sa/TSG_SA/Workshops/2023-06-13_Rel-19_WorkShop/")
    ]
