# `core.ui.world_clock_engine`

Modular World Clock & 3GPP Multi-Venue Timezone Service package decomposed from the former `world_clock_service.py` monolith:

- [`models.py`](models.py): `VenueClockState`, `WorldClockSnapshot`, `_COUNTRY_CITY_FALLBACKS`, city query normalization, and freshness helpers.
- [`sync_worker.py`](sync_worker.py): `WebTimezoneSyncWorker` (`QThread`) for Open-Meteo geocoding, TimeAPI.io / WorldTimeAPI lookup, and Geo-IP detection.
- [`formatting.py`](formatting.py): `WorldClockFormattingMixin` for compact status-bar clock pill text and rich multi-line tooltips.
- [`service.py`](service.py): `WorldClockService` handling DB meeting caching, 3-tier venue timezone resolution, and snapshot construction.
