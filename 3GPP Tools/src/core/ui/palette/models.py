"""Data model, item roles, and token scoring for the Ctrl+K Command Palette."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable, Optional, Tuple

from PyQt5.QtCore import Qt

_ROLE_KIND = Qt.UserRole + 20
_ROLE_CMD = Qt.UserRole + 21
_KIND_HEADER = "header"
_KIND_ITEM = "item"

_TDOC_PATTERN = re.compile(r"^[A-Z0-9]{2,4}-\d{6,8}$", re.IGNORECASE)


@dataclass
class CommandItem:
    """Single actionable entry rendered in the Ctrl+K Command Palette."""

    id: str
    category: str
    title: str
    subtitle: str = ""
    shortcut: str = ""
    badge: str = ""
    icon_name: str = "compass"
    callback: Optional[Callable[[], None]] = None
    keywords: str = ""


def _score_command(item: CommandItem, query: str) -> Optional[Tuple[int, str]]:
    """Return a sort rank (lower is better) if item matches query tokens."""
    tokens = [t for t in query.lower().split() if t]
    if not tokens:
        return (0, item.title.lower())

    title_l = item.title.lower()
    sub_l = item.subtitle.lower()
    kw_l = item.keywords.lower()
    cat_l = item.category.lower()
    haystack = f"{title_l} {sub_l} {kw_l} {cat_l}"

    if not all(tok in haystack for tok in tokens):
        return None

    first = tokens[0]
    if title_l == query.lower():
        tier = 0
    elif title_l.startswith(first):
        tier = 1
    elif first in title_l:
        tier = 2
    elif first in kw_l:
        tier = 3
    else:
        tier = 4
    return (tier, title_l)
