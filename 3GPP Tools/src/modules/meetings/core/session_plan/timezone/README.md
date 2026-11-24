# `modules.meetings.core.session_plan.timezone`

Modular timezone engine for 3GPP Meeting Session Plans and World Clock, anchoring UTC conversions and translating cleanly between Venue Local Time and Laptop Local Time.

## Package Architecture (Acyclic)

- `cities.py`: `KNOWN_3GPP_CITIES`, `CITY_TO_IANA_TZ`, `_SORTED_CITY_KEYS`, `MEETING_TZ_FILENAME`, and `format_utc_offset()`.
- `resolver.py`: `TimezoneResolverMixin` (laptop UTC offset, Agenda deadline regex detection, city/country lookup, IANA `zoneinfo` DST resolution).
- `persistence_and_format.py`: `TimezonePersistenceAndFormatMixin` (`meeting_timezone.json` & `session_plans.db` sync, dual-time formatting, RFC 5545 UTC conversion).
- `engine.py`: `SessionPlanTimezone` orchestrator combining both mixins.
- `__init__.py`: Public facade exporting all constants, helpers, `ZoneInfo` fallbacks, and `SessionPlanTimezone`.

## Compatibility

`src/modules/meetings/core/session_plan/session_plan_timezone.py` remains a thin backward-compatible re-export shim.
