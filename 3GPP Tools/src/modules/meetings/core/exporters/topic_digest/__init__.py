# --- File: src/modules/meetings/core/exporters/topic_digest/__init__.py ---
"""Deterministic Per-Meeting Briefing & Token-Efficient LLM Digest Builder package."""

from .company_positions import write_company_positions_input
from .constants import (
    DEFAULT_MAX_DIGEST_CHARS,
    STAGE1_DELTA_END,
    STAGE1_DELTA_START,
    STAGE2_REPORT_FILENAME,
    STAGE2_SYNTHESIS_FILENAME,
    _SUFFIX_IN_NUM_RE,
    _TREATED_STATUS_TOKENS,
)
from .extraction import (
    _rel_to_meeting_dir,
    _strip_markdown_frontmatter,
    build_contributor_matrix,
    extract_tdoc_position_digest,
    resolve_meeting_folder_id,
    rewrite_media_links,
)
from .meeting_brief import (
    build_delta_section_markdown,
    sync_meeting_brief_analysis_state,
    write_meeting_brief,
)
from .root_reports import (
    write_consolidated_briefing_markdown,
    write_e2e_synthesis_prompt,
    write_master_index_markdown,
)
from .status_and_outcomes import (
    _format_analysis_footprint_line,
    _format_tdoc_digest_card,
    _format_tdoc_outcome_lines,
    _resolve_effective_synthesis_state,
    _resolve_rec_analysis_status,
    build_tdoc_outcome_index,
    classify_meeting_lifecycle,
    format_analysis_badge,
    format_lifecycle_badge,
    format_stage2_summary_badge,
)

__all__ = [
    "DEFAULT_MAX_DIGEST_CHARS",
    "STAGE1_DELTA_END",
    "STAGE1_DELTA_START",
    "STAGE2_REPORT_FILENAME",
    "STAGE2_SYNTHESIS_FILENAME",
    "_SUFFIX_IN_NUM_RE",
    "_TREATED_STATUS_TOKENS",
    "_format_analysis_footprint_line",
    "_format_tdoc_digest_card",
    "_format_tdoc_outcome_lines",
    "_rel_to_meeting_dir",
    "_resolve_effective_synthesis_state",
    "_resolve_rec_analysis_status",
    "_strip_markdown_frontmatter",
    "build_contributor_matrix",
    "build_delta_section_markdown",
    "build_tdoc_outcome_index",
    "classify_meeting_lifecycle",
    "extract_tdoc_position_digest",
    "format_analysis_badge",
    "format_lifecycle_badge",
    "format_stage2_summary_badge",
    "resolve_meeting_folder_id",
    "rewrite_media_links",
    "sync_meeting_brief_analysis_state",
    "write_company_positions_input",
    "write_consolidated_briefing_markdown",
    "write_e2e_synthesis_prompt",
    "write_master_index_markdown",
    "write_meeting_brief",
]
