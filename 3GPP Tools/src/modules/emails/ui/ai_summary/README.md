# AI Email Summary Shared UI (`modules/emails/ui/ai_summary/`)

Reusable Qt widgets shared across `MeetingEmailInboxDialog`, `TrackedThreadDialog`, `TDocEmailsDialog`, and `EmailBriefingsDialog`:

- `model_picker.py`: `populate_summary_model_menu()` and `SummaryModelDefaultsDialog` for uniform model/effort selection across Ollama, Codex CLI, Claude CLI, and Gemini CLI.
- `summary_card.py`: Collapsible `InlineEmailSummaryCard` with live streaming, incremental `+N` update badge, undo history, copy/notes actions, and clickable `jump://msg/` & `tdoc:` links.
