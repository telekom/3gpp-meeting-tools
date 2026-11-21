# --- File: src/modules/meetings/core/exporters/topic_digest/constants.py ---
"""Shared constants and compiled regexes for the Topic Digest Builder package."""

import re

STAGE1_DELTA_START = "<!-- STAGE1_DELTA_START -->"
STAGE1_DELTA_END = "<!-- STAGE1_DELTA_END -->"
STAGE2_SYNTHESIS_FILENAME = "company_positions_e2e_synthesis.md"
STAGE2_REPORT_FILENAME = "COMPANY_VIEWS_AND_PROGRESS.md"

DEFAULT_MAX_DIGEST_CHARS = 140_000

_SUFFIX_IN_NUM_RE = re.compile(r'(?:bis|ahe|adhoc|ah\d*|-e\d*)$', re.IGNORECASE)

_TREATED_STATUS_TOKENS = (
    "agreed",
    "approved",
    "endorsed",
    "noted",
    "revised",
    "merged",
    "handled",
    "postponed",
    "withdrawn",
)
