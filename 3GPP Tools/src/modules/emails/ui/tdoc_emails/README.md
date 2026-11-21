# `tdoc_emails` Package

Modular implementation of `TDocEmailsDialog` and `StandaloneEmailReaderWindow`, decomposed cleanly from `src/modules/emails/ui/tdoc_emails_dialog.py`.

## Structure & Mixin Composition (MRO)

- `standalone_reader.py` (`StandaloneEmailReaderWindow`): Independent top-level HTML email reader window.
- `dialog.py` (`TDocEmailsDialog`): Shell inheriting from three focused mixins ahead of `QDialog`:
  1. `TDocEmailsUiSetupMixin` (`mixins/ui_setup.py`): Builds the TDoc context header, inline note strip, `FlowLayout` action bar, company/state/scope filters, and `QTextBrowser` reader.
  2. `TDocEmailsRenderingMixin` (`mixins/rendering.py`): Loads and filters TDoc family messages, renders Option A Markdown cards, navigates unread blocks, and manages status/note persistence.
  3. `TDocEmailsActionsMixin` (`mixins/actions.py`): Handles Outlook reload/read/flag/ignore links, attachment & Drafts caching, Ollama summary (`EmailSummaryDialog`), Markdown export, and LLM CLI launching (`_call_launch_cli_in_terminal` honors monkeypatching on `modules.emails.ui.tdoc_emails_dialog.launch_cli_in_terminal`).

## Backward Compatibility

`src/modules/emails/ui/tdoc_emails_dialog.py` acts as a thin re-export shim exposing `StandaloneEmailReaderWindow`, `TDocEmailsDialog`, `launch_cli_in_terminal`, and `logger`.
