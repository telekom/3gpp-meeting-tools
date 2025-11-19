from enum import Enum


class DocumentFileType(Enum):
    UNKNOWN = 0
    DOCX = 1
    DOC = 2
    PPTX = 3
    PDF = 4
    HTML = 5
    YAML = 6


class ServerType(Enum):
    PUBLIC = 1
    PRIVATE = 2
    SYNC = 3


class DocumentType(Enum):
    TDOCS_BY_AGENDA = 1
    AGENDA = 2
    TDOC = 3
    MEETING_ROOT = 4
    CHAIR_NOTES = 5
    DOCUMENTS_FOLDER = 6
    INBOX_FOLDER = 7


class TdocType(Enum):
    NORMAL = 1
    REVISION = 2
    DRAFT = 3


class WorkingGroup(Enum):
    SP = 1
    S1 = 2
    S2 = 3
    S3 = 4
    S3LI = 5
    S4 = 6
    S5 = 7
    S6 = 8
    CP = 9
    C1 = 10
    C3 = 11
    C4 = 12
    C6 = 13
    RP = 14
    R1 = 15
    R2 = 16
    R3 = 17
    R4 = 18
    R5 = 19

    @staticmethod
    def from_string(wg_str_from_tdoc: str):
        wg_str_from_tdoc = wg_str_from_tdoc.upper()
        match wg_str_from_tdoc:
            case 'SP':
                return WorkingGroup.SP
            case 'S1':
                return WorkingGroup.S1
            case 'S2':
                return WorkingGroup.S2
            case 'S3':
                return WorkingGroup.S3
            case 'S3LI':
                return WorkingGroup.S3LI
            case 'S4':
                return WorkingGroup.S4
            case 'S5':
                return WorkingGroup.S5
            case 'S6':
                return WorkingGroup.S6
            case 'RP':
                return WorkingGroup.RP
            case 'R1':
                return WorkingGroup.R1
            case 'R2':
                return WorkingGroup.R2
            case 'R3':
                return WorkingGroup.R3
            case 'R4':
                return WorkingGroup.R4
            case 'R5':
                return WorkingGroup.R5
            case 'CP':
                return WorkingGroup.CP
            case 'C1':
                return WorkingGroup.C1
            case 'C3':
                return WorkingGroup.C3
            case 'C4':
                return WorkingGroup.C4
            case 'C6':
                return WorkingGroup.C6
            case _:
                print(f'Could not parse WG {wg_str_from_tdoc}. Returning SA WG2')
                return WorkingGroup.S2

    def get_wg_folder_name(self, server_type: ServerType) -> str:
        # Groups have different names depending on where you access them!
        match server_type:
            case ServerType.PRIVATE:
                # When we are connected to 10.10.10.10
                # See https://www.3gpp.org/ftp/Meetings_3GPP_SYNC
                prefix = 'ftp'
                if server_type == ServerType.SYNC:
                    prefix = f'{prefix}/Meetings_3GPP_SYNC'
                match self:
                    case WorkingGroup.SP:
                        return f'{prefix}/SA'
                    case WorkingGroup.S1:
                        return f'{prefix}/SA/SA1'
                    case WorkingGroup.S2:
                        return f'{prefix}/SA/SA2'
                    case WorkingGroup.S3:
                        return f'{prefix}/SA/SA3'
                    case WorkingGroup.S3LI:
                        return f'{prefix}/SA/SA3LI'
                    case WorkingGroup.S4:
                        return f'{prefix}/SA/SA4'
                    case WorkingGroup.S5:
                        return f'{prefix}/SA/SA5'
                    case WorkingGroup.S6:
                        return f'{prefix}/SA/SA6'
                    case WorkingGroup.CP:
                        return f'{prefix}/CT'
                    case WorkingGroup.C1:
                        return f'{prefix}/CT/CT1'
                    case WorkingGroup.C3:
                        return f'{prefix}/CT/CT3'
                    case WorkingGroup.C4:
                        return f'{prefix}/CT/CT4'
                    case WorkingGroup.C6:
                        return f'{prefix}/CT/CT6'
                    case WorkingGroup.RP:
                        return f'{prefix}/RAN'
                    case WorkingGroup.R1:
                        return f'{prefix}/RAN/RAN1'
                    case WorkingGroup.R2:
                        return f'{prefix}/RAN/RAN2'
                    case WorkingGroup.R3:
                        return f'{prefix}/RAN/RAN3'
                    case WorkingGroup.R4:
                        return f'{prefix}/RAN/RAN4'
                    case WorkingGroup.R5:
                        return f'{prefix}/RAN/RAN5'
            case ServerType.SYNC:
                # When we are connected to the sync server
                # See https://www.3gpp.org/ftp/Meetings_3GPP_SYNC
                prefix = 'ftp'
                if server_type == ServerType.SYNC:
                    prefix = f'{prefix}/Meetings_3GPP_SYNC'
                match self:
                    case WorkingGroup.SP:
                        return f'{prefix}/SA'
                    case WorkingGroup.S1:
                        return f'{prefix}/SA1'
                    case WorkingGroup.S2:
                        return f'{prefix}/SA2'
                    case WorkingGroup.S3:
                        return f'{prefix}/SA3'
                    case WorkingGroup.S3LI:
                        return f'{prefix}/SA3LI'
                    case WorkingGroup.S4:
                        return f'{prefix}/SA4'
                    case WorkingGroup.S5:
                        return f'{prefix}/SA5'
                    case WorkingGroup.S6:
                        return f'{prefix}/SA6'
                    case WorkingGroup.CP:
                        return f'{prefix}/CT'
                    case WorkingGroup.C1:
                        return f'{prefix}/CT1'
                    case WorkingGroup.C3:
                        return f'{prefix}/CT3'
                    case WorkingGroup.C4:
                        return f'{prefix}/CT4'
                    case WorkingGroup.C6:
                        return f'{prefix}/CT6'
                    case WorkingGroup.RP:
                        return f'{prefix}/RAN'
                    case WorkingGroup.R1:
                        return f'{prefix}/RAN1'
                    case WorkingGroup.R2:
                        return f'{prefix}/RAN2'
                    case WorkingGroup.R3:
                        return f'{prefix}/RAN3'
                    case WorkingGroup.R4:
                        return f'{prefix}/RAN4'
                    case WorkingGroup.R5:
                        return f'{prefix}/RAN5'
            case _:
                prefix = 'ftp'
                match self:
                    case WorkingGroup.SP:
                        return f'{prefix}/tsg_sa/TSG_SA'
                    case WorkingGroup.S1:
                        return f'{prefix}/tsg_sa/WG1_Serv'
                    case WorkingGroup.S2:
                        return f'{prefix}/tsg_sa/WG2_Arch'
                    case WorkingGroup.S3:
                        return f'{prefix}/tsg_sa/WG3_Security'
                    case WorkingGroup.S3LI:
                        return f'{prefix}/tsg_sa/WG3_Security/TSGS3_LI'
                    case WorkingGroup.S4:
                        return f'{prefix}/tsg_sa/WG4_CODEC'
                    case WorkingGroup.S5:
                        return f'{prefix}/tsg_sa/WG5_TM'
                    case WorkingGroup.S6:
                        return f'{prefix}/tsg_sa/WG6_MissionCritical'
                    case WorkingGroup.CP:
                        return f'{prefix}/tsg_ct/TSG_CT'
                    case WorkingGroup.C1:
                        return f'{prefix}/tsg_ct/WG1_mm-cc-sm_ex-CN1'
                    case WorkingGroup.C3:
                        return f'{prefix}/tsg_ct/WG3_interworking_ex-CN3'
                    case WorkingGroup.C4:
                        return f'{prefix}/tsg_ct/WG4_protocollars_ex-CN4'
                    case WorkingGroup.C6:
                        return f'{prefix}/tsg_ct/WG6_Smartcard_Ex-T3'
                    case WorkingGroup.RP:
                        return f'{prefix}/tsg_ran/TSG_RAN'
                    case WorkingGroup.R1:
                        return f'{prefix}/tsg_ran/WG1_RL1'
                    case WorkingGroup.R2:
                        return f'{prefix}/tsg_ran/WG2_RL2'
                    case WorkingGroup.R3:
                        return f'{prefix}/tsg_ran/WG3_Iu'
                    case WorkingGroup.R4:
                        return f'{prefix}/tsg_ran/WG4_Radio'
                    case WorkingGroup.R5:
                        return f'{prefix}/tsg_ran/WG5_Test_ex-T1'

    def get_wg_inbox_folder(self, server_type: ServerType) -> str:
        """
        Returns the inbox folder for this meeting
        Args:
            server_type (object):
        """
        wg_folder = self.get_wg_folder_name(server_type)
        inbox_folder = f'{wg_folder}/Inbox'
        return inbox_folder
