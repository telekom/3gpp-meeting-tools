from enum import Enum
from typing import NamedTuple


class SpecType(Enum):
    Unknown = 1
    TS = 2
    TR = 3


class SpecReleases(NamedTuple):
    folder: str
    release: str
    base_url: str
    release_url: str


class SpecSeries(NamedTuple):
    folder: str
    series: str
    release: str
    base_url: str
    series_url: str


class SpecFile(NamedTuple):
    """
        Contains the data from a specification file as retrieved from the 3GPP FTP server and parsed from the file
        name (latest specs folder), e.g. from https://www.3gpp.org/ftp/Specs/latest/Rel-17/23_series.
    """
    file: str
    spec: str
    version: str
    series: str
    release: str
    base_url: str
    spec_url: str


class SpecVersionMapping(NamedTuple):
    """
            Contains the data from a specification file as retrieved from the 3GPP FTP server and parsed from the file
            name (latest specs folder), e.g. from https://www.3gpp.org/ftp/Specs/latest/Rel-17/23_series.
            version_mapping: 16.0.0->g00, version_mapping_inv: g00->16.0.0
    """
    spec: str
    title: str
    version_mapping: dict
    version_mapping_inv: dict
    responsible_group: str
    type: SpecType


def get_spec_full_name(spec_id: str, spec_type: SpecType) -> str:
    """
    Returns the full name of a 3GPP specification, e.g. 23.501 -> TS 23.501
    Args:
        spec_id: The specification ID, e.g. 23.501
        spec_type: The specification type, e.g. TS, TR

    Returns:
        str: The full name of this specification including the type, e.g. TS 23.501

    """
    spec_name = spec_id
    if spec_type == SpecType.TS:
        spec_name = 'TS ' + spec_name
    elif spec_type == SpecType.TR:
        spec_name = 'TR ' + spec_name

    return spec_name