# `core.ui.world_clock`

Modular World Clock & 3GPP Venue Timezone dialog package decomposed from the former `world_clock_dialog.py` monolith:

- [`venue_card.py`](venue_card.py): `_VenueClockCard` widget plus UTC offset list, source badge, and laptop-delta formatting helpers.
- [`mixins/ui_setup.py`](mixins/ui_setup.py): Reference clocks card, scrollable venue list, Working Group filter preferences, and card lifecycle updates.
- [`mixins/sync_actions.py`](mixins/sync_actions.py): Tick refresh, manual offset/DST overrides, async Geo-IP & city web-sync workers, and clean shutdown.
- [`dialog.py`](dialog.py): `WorldClockDialog` composition and timer initialization.
