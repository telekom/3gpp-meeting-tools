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

def rgb_to_hex(rgb):
    # s.Cells(1, i).Interior.color uses bgr in hex
    bgr = (rgb[2], rgb[1], rgb[0])
    strValue = '%02x%02x%02x' % bgr
    # print(strValue)
    iValue = int(strValue, 16)
    return iValue

# Some colors also used for conditional formatting
class RgbColor(NamedTuple):
    red:int
    green:int
    blue:int

    # Used for setting Excel colors
    @property
    def hex(self)->int:
        return rgb_to_hex(self)

color_magenta = RgbColor(234, 10, 142)
color_black = RgbColor(0, 0, 0)
color_white = RgbColor(255, 255, 255)
color_green = RgbColor(0, 97, 0)
color_light_green = RgbColor(198, 239, 206)
color_dark_red = RgbColor(156, 0, 6)
color_light_red = RgbColor(255, 199, 206)
color_dark_grey = RgbColor(128, 128, 128)
color_light_grey = RgbColor(217, 217, 217)
color_dark_yellow = RgbColor(156, 87, 0)
color_light_yellow = RgbColor(255, 235, 156)



