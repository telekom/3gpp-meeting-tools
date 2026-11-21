"""Cached natural sort key helper for TDocs tables and filters."""

from functools import lru_cache
import re


@lru_cache(maxsize=32768)
def natural_sort_key(s):
    """Parses text into naturally sortable chunks with caching for sub-millisecond comparisons."""
    return tuple(int(text) if text.isdigit() else text.lower() for text in re.split(r'([0-9]+)', str(s)))
