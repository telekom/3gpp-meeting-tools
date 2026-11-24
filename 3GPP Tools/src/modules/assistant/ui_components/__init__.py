"""Modular UI components, prompt editor, markdown renderer, and cockpit widgets for 3GPP Assistant."""

from modules.assistant.ui_components.cockpit_widgets import (
    ActivityTimelineWidget,
    DiagnosticsTraceWidget,
    EvidenceInspectorWidget,
)
from modules.assistant.ui_components.prompt_and_markdown import (
    MultilinePromptEdit,
    StarterChipsWidget,
    render_assistant_message_card,
    render_markdown_to_html,
    render_system_message_card,
    render_user_message_card,
)
from modules.assistant.ui_components.styles import (
    BADGE_CATALOGUE,
    BADGE_PROTOCOL,
    BADGE_SPEC_TEXT,
    CARD_FRAME_STYLE,
    CARD_HEADER_STYLE,
    PROMPT_INPUT_STYLE,
    STARTER_CHIP_STYLE,
    STARTER_PROMPTS,
    STATUS_PILL_NEUTRAL,
    STATUS_PILL_OFFLINE,
    STATUS_PILL_ONLINE,
    TOOL_ICONS,
)

__all__ = [
    "ActivityTimelineWidget",
    "BADGE_CATALOGUE",
    "BADGE_PROTOCOL",
    "BADGE_SPEC_TEXT",
    "CARD_FRAME_STYLE",
    "CARD_HEADER_STYLE",
    "DiagnosticsTraceWidget",
    "EvidenceInspectorWidget",
    "MultilinePromptEdit",
    "PROMPT_INPUT_STYLE",
    "STARTER_CHIP_STYLE",
    "STARTER_PROMPTS",
    "STATUS_PILL_NEUTRAL",
    "STATUS_PILL_OFFLINE",
    "STATUS_PILL_ONLINE",
    "StarterChipsWidget",
    "TOOL_ICONS",
    "render_assistant_message_card",
    "render_markdown_to_html",
    "render_system_message_card",
    "render_user_message_card",
]
