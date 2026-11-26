# AI Email Summary Shared UI (`modules/emails/ui/ai_summary/`)

Reusable Qt widgets shared across `MeetingEmailInboxDialog`, `TrackedThreadDialog`, `TDocEmailsDialog`, and `EmailBriefingsDialog`:

- `model_picker.py`: `populate_summary_model_menu()` and `SummaryModelDefaultsDialog` for uniform model/effort selection across Ollama, Codex CLI, Claude CLI, and Gemini CLI, plus M365 Copilot launch target preference.
- `m365_copilot_dialog.py`: Modeless `M365CopilotHandoffDialog` for 2-step Send & Paste-Back M365 Copilot summaries.
- `summary_card.py`: Collapsible `InlineEmailSummaryCard` with live streaming, incremental `+N` update badge, undo history, copy/notes actions, M365 Copilot handoff, and clickable `jump://msg/` & `tdoc:` links.
