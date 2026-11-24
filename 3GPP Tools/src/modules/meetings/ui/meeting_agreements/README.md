# `modules.meetings.ui.meeting_agreements`

Modular Chairman's Notes viewer and export console for a single 3GPP meeting's agreements, working assumptions, pCRs, CRs, LSs, and FFS items.

## Package Architecture (Acyclic)

- `model.py`: `_clean_agreement_preview()` and `AgreementsTableModel` (filtering, styling, tooltips).
- `mixins/ui_setup.py`: `MeetingAgreementsUiSetupMixin` (header ribbon, KPI chips, search/kind/AI filters, master-detail splitter, bottom actions; 100% `core.ui.theme.C` design tokens).
- `mixins/actions_and_export.py`: `MeetingAgreementsActionsMixin` (live filtering, HTML detail viewer rendering, parent TDoc localization, clipboard citation copy, GFM Markdown export, `focus_tdoc`).
- `dialog.py`: `MeetingAgreementsDialog` composing the mixins and initializing the table model.
- `__init__.py`: Public facade exporting `MeetingAgreementsDialog`, `AgreementsTableModel`, `_clean_agreement_preview`, and `logger`.

## Compatibility

`src/modules/meetings/ui/meeting_agreements_dialog.py` remains a thin backward-compatible re-export shim.
