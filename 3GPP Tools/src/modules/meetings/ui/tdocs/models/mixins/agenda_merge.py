"""Live TDocsByAgenda & Chairman's Notes merging mixin for TDocsTableModel."""

import logging
import re

from modules.meetings.core.tdocs.chair_notes_agreements_parser import parse_chair_result

logger = logging.getLogger("modules.meetings.ui.tdocs_models")


class TDocsAgendaMergeMixin:
    """Merges agenda updates, injects on-the-fly revisions, and reconciles statuses."""

    def merge_agenda_data(self, agenda_data: dict, ui_logger=None):
        self.beginResetModel()

        existing_tdocs = {row.get('TDoc', ''): idx for idx, row in enumerate(self._data)}
        tdoc_dict = {row.get('TDoc', ''): row for row in self._data}
        new_rows = []
        rev_links_to_apply = []  # list of (source_tdoc_id, [target_tdoc_ids])

        unsettled_statuses = {'unknown', '', '-', 'not handled', 'not treated', 'unhandled'}
        pending_statuses = unsettled_statuses | {'available', 'reserved', 'open', 'opened', 'handled'}
        settled_outcomes = {
            'approved', 'agreed', 'parallel agreed', 'endorsed', 'noted', 'postponed',
            'withdrawn', 'revised', 'merged', 'replied to', 'not agreed', 'not approved',
            'not noted', 'not endorsed', 'not pursued', 'rejected', 'for e-mail approval',
        }

        def _append_revised_to(row_obj: dict, targets: list):
            if not targets:
                return
            curr_revs = str(row_obj.get('Revised to', '') or '').strip()
            rev_list = [r.strip() for r in curr_revs.split(',') if r.strip() and r.strip().lower() != 'none']
            seen_upper = {r.upper() for r in rev_list}
            for tgt in targets:
                if tgt.upper() not in seen_upper:
                    rev_list.append(tgt)
                    seen_upper.add(tgt.upper())
            row_obj['Revised to'] = ', '.join(rev_list)

        for tdoc_id, info in agenda_data.items():
            comments = info.get('Comments', '')
            email = info.get('e-mail_Discussion', '')
            has_result_key = 'Result' in info
            raw_res = info.get('Result', '')
            clean_status, res_remark, res_rev_targets = parse_chair_result(raw_res)

            base_remarks = ""
            if comments and email:
                base_remarks = f"{comments}\n\n[e-mail_Discussion]: {email}"
            elif comments:
                base_remarks = comments
            elif email:
                base_remarks = email

            remarks = base_remarks
            if res_remark and res_remark.lower() not in remarks.lower():
                remarks = f"{remarks}\n{res_remark}" if remarks else res_remark

            if tdoc_id in existing_tdocs:
                idx = existing_tdocs[tdoc_id]
                existing_remarks = str(self._data[idx].get('Secretary Remarks', '') or '').strip()
                if base_remarks:
                    final_remarks = base_remarks
                else:
                    final_remarks = existing_remarks
                if res_remark and res_remark.lower() not in final_remarks.lower():
                    final_remarks = f"{final_remarks}\n{res_remark}" if final_remarks else res_remark
                if final_remarks:
                    self._data[idx]['Secretary Remarks'] = final_remarks

                if info.get('Sub-Item'):
                    self._data[idx]['Sub-Item'] = info.get('Sub-Item')
                if info.get('_agenda_order'):
                    self._data[idx]['_agenda_order'] = info.get('_agenda_order')
                if info.get('_sub_item_order'):
                    self._data[idx]['_sub_item_order'] = info.get('_sub_item_order')

                if info.get('Agenda Item') and self._data[idx].get('Agenda Item', '') in ('', 'N/A', 'None', '-'):
                    self._data[idx]['Agenda Item'] = info.get('Agenda Item')

                if info.get('Related WIs') and not self._data[idx].get('Related WIs'):
                    self._data[idx]['Related WIs'] = info.get('Related WIs')

                if info.get('Title') and not self._data[idx].get('Title'):
                    self._data[idx]['Title'] = info.get('Title')

                if info.get('Source') and not self._data[idx].get('Source'):
                    self._data[idx]['Source'] = info.get('Source')
                    self._apply_company_sanitization([self._data[idx]])

                if info.get('For') and not self._data[idx].get('For'):
                    self._data[idx]['For'] = info.get('For')

                if res_rev_targets:
                    _append_revised_to(self._data[idx], res_rev_targets)
                    rev_links_to_apply.append((tdoc_id, res_rev_targets))

                if clean_status:
                    curr_st = str(self._data[idx].get('TDoc Status', 'Unknown') or '').strip()
                    curr_low = curr_st.lower()
                    if (
                        any(sym in clean_status for sym in ('🟢', '🟡', '🔵'))
                        or 'agreement' in clean_status.lower()
                        or 'working assumption' in clean_status.lower()
                    ):
                        self._data[idx]['TDoc Status'] = clean_status
                    elif curr_low in unsettled_statuses or (
                        curr_low in pending_statuses and clean_status.lower() in settled_outcomes
                    ):
                        self._data[idx]['TDoc Status'] = clean_status

            else:
                logger.info(f"✨ Injecting new on-the-fly TDoc: {tdoc_id}")

                agenda_item = info.get('Agenda Item') or 'N/A'
                doc_type = 'Revision'
                predecessor = None
                base_tdoc = None

                abstract = info.get('Abstract', '')
                source = info.get('Source', '')
                title = info.get('Title', '')
                related_wis = (
                    info.get('Related WIs') or
                    info.get('Related WI') or
                    info.get('Related WI(s)') or
                    info.get('Work Item') or
                    info.get('WI') or
                    ''
                )

                comment_match = re.search(
                    r'(?:revision of|rev of)\s*([A-Z][A-Z0-9]{1,3}-\d{5,8}(?:r\d{1,2}[a-zA-Z]?)?)',
                    remarks,
                    re.IGNORECASE
                )
                id_match = re.search(r'^(.*?)-?(?:r|rev)(\d{1,2}[a-zA-Z]?)$', tdoc_id, re.IGNORECASE)

                if comment_match:
                    predecessor = comment_match.group(1).upper()
                    base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', predecessor, re.IGNORECASE)
                    base_tdoc = base_match.group(1).upper() if base_match else predecessor
                elif id_match:
                    base_tdoc = id_match.group(1).upper()
                    predecessor = base_tdoc
                    try:
                        rev_num = int(re.match(r'\d+', id_match.group(2)).group())
                        if rev_num > 1:
                            prev_rev = f"{base_tdoc}r{rev_num - 1:02d}"
                            if prev_rev in tdoc_dict:
                                predecessor = prev_rev
                    except Exception:
                        pass

                if not predecessor:
                    doc_type = 'TDoc'

                base_row = tdoc_dict.get(base_tdoc, {}) if base_tdoc else {}

                def fuzzy_get(key):
                    for k, v in base_row.items():
                        if str(k).strip().lower() == key.lower():
                            return str(v).strip() if v is not None else ''
                    return ''

                if base_row:
                    agenda_item = fuzzy_get('Agenda Item') or agenda_item
                    doc_type = fuzzy_get('Type') or 'TDoc'
                    abstract = fuzzy_get('Abstract')
                    if not source:
                        source = fuzzy_get('Source')
                    if not title:
                        title = fuzzy_get('Title')
                    if not related_wis:
                        related_wis = fuzzy_get('Related WIs')

                if not has_result_key:
                    initial_status = 'Unknown'
                elif clean_status:
                    initial_status = clean_status
                elif res_remark:
                    initial_status = 'Not Handled'
                else:
                    initial_status = ''

                new_row = {
                    'TDoc': tdoc_id,
                    'Title': title if title else 'Unknown (Parsed from Agenda)',
                    'Source': source,
                    'Secretary Remarks': remarks,
                    'Agenda Item': agenda_item,
                    'Sub-Item': info.get('Sub-Item') or (base_row.get('Sub-Item', '') if base_row else ''),
                    'Type': info.get('Type') or doc_type,
                    'For': info.get('For', ''),
                    'Abstract': abstract,
                    'TDoc Status': initial_status,
                    'Is revision of': predecessor if predecessor else '',
                    'Revised to': ', '.join(res_rev_targets),
                    'Original LS': '',
                    'Reply in': '',
                    'Related WIs': related_wis,
                    '_agenda_order': info.get('_agenda_order'),
                    '_sub_item_order': info.get('_sub_item_order', 0),
                }
                if res_rev_targets:
                    rev_links_to_apply.append((tdoc_id, res_rev_targets))
                new_rows.append(new_row)

        if new_rows:
            self._apply_company_sanitization(new_rows)
            self._precompute_row_wi_tokens(new_rows)
            self._data.extend(new_rows)
            tdoc_dict = {row.get('TDoc', ''): row for row in self._data}

            for new_row in new_rows:
                tdoc_id = new_row['TDoc']
                predecessor = new_row['Is revision of']

                if predecessor:
                    if predecessor not in tdoc_dict:
                        base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', predecessor, re.IGNORECASE)
                        if base_match:
                            base_tdoc = base_match.group(1).upper()
                            if base_tdoc in tdoc_dict:
                                predecessor = base_tdoc

                    if predecessor in tdoc_dict:
                        _append_revised_to(tdoc_dict[predecessor], [tdoc_id])

        for src_id, targets in rev_links_to_apply:
            for tgt in targets:
                tgt_row = tdoc_dict.get(tgt.upper()) or tdoc_dict.get(tgt)
                if tgt_row and not str(tgt_row.get('Is revision of', '') or '').strip():
                    tgt_row['Is revision of'] = src_id

        self.valid_tdocs = {str(r.get("TDoc", "")) for r in self._data if r.get("TDoc")}
        self._apply_user_data_logic()
        self._rebuild_family_cache()
        self._recompute_cached_email_stats()
        self._precompute_agenda_sort_tuples(self._data)
        self._detect_and_tag_agreements(self._data, getattr(self, "chair_agreements", None))
        self.endResetModel()
