# Email Briefings Hub (`modules/emails/ui/briefings/`)

Modeless `EmailBriefingsDialog` (`Qt.Window`, `Qt.NonModal`) providing Per-AI 2-stage Map-Reduce cross-thinking briefings and ad-hoc Saved Selection summaries.

- `buckets.py`: `build_meeting_ai_buckets()` aggregating TDoc families and Tracked Threads by Agenda Item.
- `dialog.py`: `EmailBriefingsDialog` shell and lifecycle.
- `mixins/ui_setup.py` & `mixins/actions.py`: Navigation tree (`👀 Followed AIs` pinned first), Focus Prompt, streaming reader with clickable `tdoc:` / `thread:` drill-down links, and collapsible underlying leaf summaries drawer.
