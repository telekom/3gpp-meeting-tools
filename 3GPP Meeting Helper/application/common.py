from enum import Enum
from typing import NamedTuple


class ExportType(Enum):
    NONE = 0
    PDF = 1
    HTML = 2
    DOCX = 3


class ActionAfter(Enum):
    NOTHING = 0
    CLOSE_FILE = 1
    CLOSE_AND_DELETE_FILE = 2


class DocumentMetadata(NamedTuple):
    title: str | None  # Title of this TDoc
    source: str | None # Source companies of this TDoc
    path: str | None   # Filepath
