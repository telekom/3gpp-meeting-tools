import logging
from pathlib import Path
from typing import Optional, List, Dict, Any, Union, Callable, Tuple

from modules.nas.core.parsing.protocol_parser_constants import (
    RE_PART_INDEX, RE_SPEC_NUMBER, RE_VERSION_STEM
)
from modules.specifications.utils.utils import file_version_to_version
from modules.nas.core.parsing.rrc_asn1_parser import RRCAsn1Parser
from modules.nas.core.parsing.ran3_asn1_parser import RAN3Asn1Parser
from modules.nas.core.parsing.nas_parser import NASDocxParser
from modules.nas.core.parsing.pfcp_parser import PFCPDocxParser
from modules.nas.core.parsing.gtpu_parser import GTPUDocxParser
from modules.nas.core.parsing.pdu_session_up_parser import PDUSessionUPDocxParser  # <-- Add import


class ProtocolDocxDispatcher:
    """Unified entry point routing specifications to dedicated RRC, RAN3, or NAS parsers."""

    def __init__(self, docx_paths: Union[Path, str, List[Union[Path, str]]]):
        if isinstance(docx_paths, (str, Path)):
            self.docx_paths = [Path(docx_paths)]
        else:
            self.docx_paths = [Path(p) for p in docx_paths]

        self.docx_paths.sort(key=self._extract_part_index)
        self.logger = logging.getLogger(__name__)

    @staticmethod
    def _extract_part_index(path: Path) -> int:
        match = RE_PART_INDEX.search(path.name)
        return int(match.group(1)) if match else 0

    def extract_spec_number(self) -> str:
        if not self.docx_paths:
            return "24.501"
        match = RE_SPEC_NUMBER.search(self.docx_paths[0].stem)
        return f"{match.group(1)}.{match.group(2)}" if match else "24.501"

    def extract_version_from_filename(self) -> str:
        if not self.docx_paths:
            return ""
        stem = self.docx_paths[0].stem
        match = RE_VERSION_STEM.search(stem)
        if match:
            parsed_ver = file_version_to_version(match.group(1))
            if parsed_ver:
                return parsed_ver
        return stem

    def parse(
            self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        spec_num = self.extract_spec_number()

        # 1. Route RAN2 RRC specifications (38.331, 36.331)
        if any(s in spec_num for s in ["38.331", "36.331"]):
            parser = RRCAsn1Parser(self.docx_paths, spec_number=spec_num)
            return parser.parse(progress_callback=progress_callback)

        # 2. Route RAN3 Object Set specifications (38.413, 38.423, 38.473, 38.463, 36.413)
        if any(s in spec_num for s in ["38.413", "38.423", "38.473", "38.463", "36.413", "38.412"]):
            parser = RAN3Asn1Parser(self.docx_paths, spec_number=spec_num)
            return parser.parse(progress_callback=progress_callback)

        # 3. Route PFCP specifications (TS 29.244)
        if "29.244" in spec_num:
            parser = PFCPDocxParser(self.docx_paths, spec_number=spec_num)
            return parser.parse(progress_callback=progress_callback)

        # 4. Route GTP-U specifications (TS 29.281)
        if "29.281" in spec_num:
            parser = GTPUDocxParser(self.docx_paths, spec_number=spec_num)
            return parser.parse(progress_callback=progress_callback)

        # 5. Route PDU Session UP specifications (TS 38.415)
        if "38.415" in spec_num:
            parser = PDUSessionUPDocxParser(self.docx_paths, spec_number=spec_num)
            return parser.parse(progress_callback=progress_callback)

        # 6. Route standard NAS specifications (24.501, 24.301, 24.008)
        parser = NASDocxParser(self.docx_paths)
        messages, ie_definitions = parser.parse(progress_callback=progress_callback)

        if progress_callback:
            progress_callback(f"Extracted {len(messages)} messages and {len(ie_definitions)} definitions.", 95)

        return messages, ie_definitions