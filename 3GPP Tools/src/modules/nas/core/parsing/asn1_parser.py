from modules.nas.core.parsing.asn1_base_parser import BaseAsn1DocxParser
from modules.nas.core.parsing.rrc_asn1_parser import RRCAsn1Parser
from modules.nas.core.parsing.ran3_asn1_parser import RAN3Asn1Parser

# Backward-compatibility aliases
RRCAsn1DocxParser = RRCAsn1Parser
ASN1DocxParser = RAN3Asn1Parser

__all__ = [
    "BaseAsn1DocxParser",
    "RRCAsn1Parser",
    "RAN3Asn1Parser",
    "RRCAsn1DocxParser",
    "ASN1DocxParser",
]