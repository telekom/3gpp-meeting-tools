"""Composite triage & reader-dispatch mixin for MeetingEmailInboxDialog."""

from __future__ import annotations

from .reader_dispatch_and_export import InboxReaderDispatchMixin
from .triage_mutations import InboxTriageMutationsMixin


class InboxTriageActionsMixin(InboxTriageMutationsMixin, InboxReaderDispatchMixin):
    """Composite mixin combining triage mutations and Outlook/reader dispatch."""


__all__ = [
    "InboxTriageActionsMixin",
    "InboxTriageMutationsMixin",
    "InboxReaderDispatchMixin",
]
