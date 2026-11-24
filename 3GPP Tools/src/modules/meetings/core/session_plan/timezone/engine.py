"""SessionPlanTimezone orchestrator combining resolver and persistence/formatting mixins."""

from __future__ import annotations

from modules.meetings.core.session_plan.timezone.cities import MEETING_TZ_FILENAME
from modules.meetings.core.session_plan.timezone.persistence_and_format import (
    TimezonePersistenceAndFormatMixin,
)
from modules.meetings.core.session_plan.timezone.resolver import TimezoneResolverMixin


class SessionPlanTimezone(
    TimezoneResolverMixin,
    TimezonePersistenceAndFormatMixin,
):
    """Manages timezone calculations, dual clocks, per-meeting persistence, and UTC conversions."""

    MEETING_TZ_FILENAME = MEETING_TZ_FILENAME
