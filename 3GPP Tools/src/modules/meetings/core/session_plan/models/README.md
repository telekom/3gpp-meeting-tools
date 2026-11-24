# `modules.meetings.core.session_plan.models`

Modular SA2 Session Plan domain package decomposed from the former `session_plan_models.py` monolith:

- [`dataclasses.py`](dataclasses.py): Core grid, snapshot, annotation, palette, and delegation dataclasses.
- [`delegation_resolver.py`](delegation_resolver.py): Agenda Item wildcard/range matching, boundary-aware text lookup, effective delegate resolution, unassigned/not-followed detection, and unique initials calculation.
- [`display_and_rooms.py`](display_and_rooms.py): Slot display normalization, break/time range labels, expanded HTML bullet formatting, and room/location lookup helpers.
