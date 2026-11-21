# --- File: src/modules/meetings/core/session_plan/db/mixins/delegation.py ---
"""Delegation roster, default/per-meeting topic ownership, session appointments, and JSON import/export."""

from typing import Dict

from modules.meetings.core.session_plan.session_plan_models import (
    DELEGATE_PALETTE,
    DelegationMember,
    DelegationPlan,
    SessionAppointment,
    TopicOwnership,
)


class SessionPlanDelegationMixin:
    """Handles delegation roster, topic owners, session appointments, and DelegationPlan JSON."""

    def get_delegation_members(self, meeting_number: str) -> Dict[str, DelegationMember]:
        """
        Returns map of name -> DelegationMember.
        Combines persistent cross-meeting default roster with any meeting-specific updates.
        """
        self._sync_from_global_defaults()
        with self._get_connection() as conn:
            def_rows = conn.execute(
                "SELECT name, email, color, is_self, notes FROM default_delegation_roster ORDER BY id ASC"
            ).fetchall()
            mtg_rows = conn.execute(
                "SELECT name, email, color, is_self, notes FROM delegation_roster WHERE meeting_number = ? ORDER BY id ASC",
                (str(meeting_number),),
            ).fetchall()

            result: Dict[str, DelegationMember] = {}
            for r in def_rows:
                result[r["name"]] = DelegationMember(
                    name=r["name"],
                    email=r["email"] or "",
                    color=r["color"] or DELEGATE_PALETTE[len(result) % len(DELEGATE_PALETTE)],
                    is_self=bool(r["is_self"]),
                    notes=r["notes"] or "",
                )
            for r in mtg_rows:
                result[r["name"]] = DelegationMember(
                    name=r["name"],
                    email=r["email"] or "",
                    color=r["color"] or DELEGATE_PALETTE[len(result) % len(DELEGATE_PALETTE)],
                    is_self=bool(r["is_self"]),
                    notes=r["notes"] or "",
                )
            return result

    def save_delegation_member(self, meeting_number: str, member: DelegationMember, update_default: bool = True):
        """Adds or updates a delegation member and persists across meetings by default."""
        with self._get_connection() as conn:
            clean_name = member.name.strip()
            clean_email = member.email.strip()
            clean_color = member.color.strip() or "#2563EB"
            is_self_int = 1 if member.is_self else 0
            clean_notes = member.notes.strip()

            # If marking as self, unmark any previous self in this meeting and defaults
            if member.is_self:
                conn.execute(
                    "UPDATE delegation_roster SET is_self = 0 WHERE meeting_number = ?",
                    (str(meeting_number),),
                )
                if update_default:
                    conn.execute("UPDATE default_delegation_roster SET is_self = 0")

            conn.execute(
                """
                INSERT INTO delegation_roster (meeting_number, name, email, color, is_self, notes)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(meeting_number, name) DO UPDATE SET
                    email = excluded.email,
                    color = excluded.color,
                    is_self = excluded.is_self,
                    notes = excluded.notes
                """,
                (str(meeting_number), clean_name, clean_email, clean_color, is_self_int, clean_notes),
            )
            if update_default:
                conn.execute(
                    """
                    INSERT INTO default_delegation_roster (name, email, color, is_self, notes)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(name) DO UPDATE SET
                        email = excluded.email,
                        color = excluded.color,
                        is_self = excluded.is_self,
                        notes = excluded.notes
                    """,
                    (clean_name, clean_email, clean_color, is_self_int, clean_notes),
                )
                self._flush_global_defaults_conn(conn)

    def delete_delegation_member(self, meeting_number: str, name: str, update_default: bool = True):
        """Deletes a delegation member and unassigns any of their topics/appointments."""
        with self._get_connection() as conn:
            clean_name = name.strip()
            conn.execute(
                "DELETE FROM delegation_roster WHERE meeting_number = ? AND name = ?",
                (str(meeting_number), clean_name),
            )
            conn.execute(
                "DELETE FROM topic_owners WHERE meeting_number = ? AND owner_name = ?",
                (str(meeting_number), clean_name),
            )
            conn.execute(
                "DELETE FROM session_appointments WHERE meeting_number = ? AND appointed_delegate = ?",
                (str(meeting_number), clean_name),
            )
            if update_default:
                conn.execute("DELETE FROM default_delegation_roster WHERE name = ?", (clean_name,))
                conn.execute("DELETE FROM default_topic_owners WHERE owner_name = ?", (clean_name,))
                self._flush_global_defaults_conn(conn)

    def set_self_delegate(self, meeting_number: str, name: str, update_default: bool = True):
        """Designates the given member name as the local user ('is_self')."""
        with self._get_connection() as conn:
            clean_name = name.strip()
            conn.execute(
                "UPDATE delegation_roster SET is_self = 0 WHERE meeting_number = ?",
                (str(meeting_number),),
            )
            conn.execute(
                "UPDATE delegation_roster SET is_self = 1 WHERE meeting_number = ? AND name = ?",
                (str(meeting_number), clean_name),
            )
            if update_default:
                conn.execute("UPDATE default_delegation_roster SET is_self = 0")
                conn.execute("UPDATE default_delegation_roster SET is_self = 1 WHERE name = ?", (clean_name,))
                self._flush_global_defaults_conn(conn)

    # --- Topic Ownership CRUD (Default AI Assignment Persistent Across Meetings) ---

    def get_default_topic_owners(self) -> Dict[str, TopicOwnership]:
        """Returns the cross-meeting persistent default AI assignments."""
        self._sync_from_global_defaults()
        with self._get_connection() as conn:
            rows = conn.execute(
                "SELECT ai_number, owner_name, notes FROM default_topic_owners ORDER BY id ASC"
            ).fetchall()
            return {
                r["ai_number"]: TopicOwnership(
                    meeting_number="__DEFAULT__",
                    ai_number=r["ai_number"],
                    owner_name=r["owner_name"],
                    notes=r["notes"] or "",
                )
                for r in rows
            }

    def get_topic_owners(self, meeting_number: str) -> Dict[str, TopicOwnership]:
        """
        Returns map of ai_number -> TopicOwnership for the given meeting.
        Starts from persistent cross-meeting Default AI Assignments and overlays
        any meeting-level updates, ensuring assignments persist across meetings
        automatically while allowing updates between meetings.
        """
        self._sync_from_global_defaults()
        with self._get_connection() as conn:
            def_rows = conn.execute(
                "SELECT ai_number, owner_name, notes FROM default_topic_owners ORDER BY id ASC"
            ).fetchall()
            mtg_rows = conn.execute(
                "SELECT ai_number, owner_name, notes FROM topic_owners WHERE meeting_number = ? ORDER BY id ASC",
                (str(meeting_number),),
            ).fetchall()

            result: Dict[str, TopicOwnership] = {}
            for r in def_rows:
                result[r["ai_number"]] = TopicOwnership(
                    meeting_number=str(meeting_number),
                    ai_number=r["ai_number"],
                    owner_name=r["owner_name"],
                    notes=r["notes"] or "",
                )
            for r in mtg_rows:
                result[r["ai_number"]] = TopicOwnership(
                    meeting_number=str(meeting_number),
                    ai_number=r["ai_number"],
                    owner_name=r["owner_name"],
                    notes=r["notes"] or "",
                )
            return result

    def set_topic_owner(
        self,
        meeting_number: str,
        ai_number: str,
        owner_name: str,
        notes: str = "",
        update_default: bool = True,
    ):
        """
        Sets the Default AI Assignment for `ai_number`.
        Persists both to the meeting and (when `update_default=True`) to the
        cross-meeting default table and global JSON store so subsequent meetings
        automatically inherit the latest AI assignment.
        """
        clean_ai = ai_number.strip()
        clean_owner = owner_name.strip()
        clean_notes = notes.strip()
        with self._get_connection() as conn:
            if not clean_owner:
                conn.execute(
                    "DELETE FROM topic_owners WHERE meeting_number = ? AND ai_number = ?",
                    (str(meeting_number), clean_ai),
                )
                if update_default:
                    conn.execute("DELETE FROM default_topic_owners WHERE ai_number = ?", (clean_ai,))
                    # Also remove stale per-meeting entries in this DB so clearing unassigns cleanly
                    conn.execute("DELETE FROM topic_owners WHERE ai_number = ?", (clean_ai,))
            else:
                conn.execute(
                    """
                    INSERT INTO topic_owners (meeting_number, ai_number, owner_name, notes)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT(meeting_number, ai_number) DO UPDATE SET
                        owner_name = excluded.owner_name,
                        notes = excluded.notes
                    """,
                    (str(meeting_number), clean_ai, clean_owner, clean_notes),
                )
                if update_default:
                    conn.execute(
                        """
                        INSERT INTO default_topic_owners (ai_number, owner_name, notes)
                        VALUES (?, ?, ?)
                        ON CONFLICT(ai_number) DO UPDATE SET
                            owner_name = excluded.owner_name,
                            notes = excluded.notes
                        """,
                        (clean_ai, clean_owner, clean_notes),
                    )
            if update_default:
                self._flush_global_defaults_conn(conn)

    def delete_topic_owner(self, meeting_number: str, ai_number: str, update_default: bool = True):
        clean_ai = ai_number.strip()
        with self._get_connection() as conn:
            conn.execute(
                "DELETE FROM topic_owners WHERE meeting_number = ? AND ai_number = ?",
                (str(meeting_number), clean_ai),
            )
            if update_default:
                conn.execute("DELETE FROM default_topic_owners WHERE ai_number = ?", (clean_ai,))
                conn.execute("DELETE FROM topic_owners WHERE ai_number = ?", (clean_ai,))
                self._flush_global_defaults_conn(conn)

    # --- Session Appointments CRUD (Session-Level Collision / Ad-Hoc Overrides) ---

    def get_session_appointments(self, meeting_number: str) -> Dict[str, SessionAppointment]:
        """Returns map of slot_key -> SessionAppointment (scoped strictly to this meeting)."""
        with self._get_connection() as conn:
            rows = conn.execute(
                "SELECT slot_key, appointed_delegate, original_owner, reason FROM session_appointments WHERE meeting_number = ?",
                (str(meeting_number),),
            ).fetchall()
            return {
                r["slot_key"]: SessionAppointment(
                    meeting_number=str(meeting_number),
                    slot_key=r["slot_key"],
                    appointed_delegate=r["appointed_delegate"],
                    original_owner=r["original_owner"] or "",
                    reason=r["reason"] or "",
                )
                for r in rows
            }

    def set_session_appointment(
        self,
        meeting_number: str,
        slot_key: str,
        appointed_delegate: str,
        original_owner: str = "",
        reason: str = "",
    ):
        clean_delegate = appointed_delegate.strip()
        with self._get_connection() as conn:
            if not clean_delegate:
                conn.execute(
                    "DELETE FROM session_appointments WHERE meeting_number = ? AND slot_key = ?",
                    (str(meeting_number), slot_key.strip()),
                )
            else:
                conn.execute(
                    """
                    INSERT INTO session_appointments (meeting_number, slot_key, appointed_delegate, original_owner, reason)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(meeting_number, slot_key) DO UPDATE SET
                        appointed_delegate = excluded.appointed_delegate,
                        original_owner = excluded.original_owner,
                        reason = excluded.reason
                    """,
                    (str(meeting_number), slot_key.strip(), clean_delegate, original_owner.strip(), reason.strip()),
                )

    def delete_session_appointment(self, meeting_number: str, slot_key: str):
        with self._get_connection() as conn:
            conn.execute(
                "DELETE FROM session_appointments WHERE meeting_number = ? AND slot_key = ?",
                (str(meeting_number), slot_key.strip()),
            )

    def clear_session_appointments(self, meeting_number: str) -> int:
        """Clears all session-level overrides for a meeting while preserving persistent Default AI Assignments."""
        with self._get_connection() as conn:
            cur = conn.execute(
                "DELETE FROM session_appointments WHERE meeting_number = ?",
                (str(meeting_number),),
            )
            return cur.rowcount

    # --- Aggregate Delegation Plan ---

    def get_delegation_plan(self, meeting_number: str) -> DelegationPlan:
        return DelegationPlan(
            meeting_number=str(meeting_number),
            members=self.get_delegation_members(meeting_number),
            topic_owners=self.get_topic_owners(meeting_number),
            appointments=self.get_session_appointments(meeting_number),
        )

    def export_delegation_plan_dict(self, meeting_number: str) -> dict:
        """Serializes delegation state to JSON-serializable dictionary."""
        plan = self.get_delegation_plan(meeting_number)
        return {
            "3gpp_delegation_plan_version": 1,
            "meeting_number": str(meeting_number),
            "members": [
                {
                    "name": m.name,
                    "email": m.email,
                    "color": m.color,
                    "is_self": m.is_self,
                    "notes": m.notes,
                }
                for m in plan.members.values()
            ],
            "topic_owners": {
                ai: {"owner": to.owner_name, "notes": to.notes}
                for ai, to in plan.topic_owners.items()
            },
            "appointments": {
                sk: {
                    "appointed_delegate": ap.appointed_delegate,
                    "original_owner": ap.original_owner,
                    "reason": ap.reason,
                }
                for sk, ap in plan.appointments.items()
            },
        }

    def import_delegation_plan_dict(self, meeting_number: str, data: dict, merge: bool = False) -> int:
        """
        Imports delegation state from dictionary.
        If merge=False, clears existing roster, topic owners, and appointments for this meeting.
        Also updates persistent cross-meeting defaults.
        Returns count of imported/updated items.
        """
        imported_count = 0
        with self._get_connection() as conn:
            m_num = str(meeting_number)
            if not merge:
                conn.execute("DELETE FROM delegation_roster WHERE meeting_number = ?", (m_num,))
                conn.execute("DELETE FROM topic_owners WHERE meeting_number = ?", (m_num,))
                conn.execute("DELETE FROM session_appointments WHERE meeting_number = ?", (m_num,))
                conn.execute("DELETE FROM default_delegation_roster")
                conn.execute("DELETE FROM default_topic_owners")

            # 1. Members
            for m_data in data.get("members", []):
                name = m_data.get("name", "").strip()
                if not name:
                    continue
                email = m_data.get("email", "").strip()
                color = m_data.get("color", "").strip() or "#2563EB"
                is_self = 1 if m_data.get("is_self") else 0
                notes = m_data.get("notes", "").strip()
                conn.execute(
                    """
                    INSERT INTO delegation_roster (meeting_number, name, email, color, is_self, notes)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ON CONFLICT(meeting_number, name) DO UPDATE SET
                        email = excluded.email,
                        color = excluded.color,
                        is_self = excluded.is_self,
                        notes = excluded.notes
                    """,
                    (m_num, name, email, color, is_self, notes),
                )
                conn.execute(
                    """
                    INSERT INTO default_delegation_roster (name, email, color, is_self, notes)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(name) DO UPDATE SET
                        email = excluded.email,
                        color = excluded.color,
                        is_self = excluded.is_self,
                        notes = excluded.notes
                    """,
                    (name, email, color, is_self, notes),
                )
                imported_count += 1

            # 2. Topic Owners
            for ai, to_data in data.get("topic_owners", {}).items():
                owner = to_data.get("owner", "") if isinstance(to_data, dict) else str(to_data)
                notes = to_data.get("notes", "") if isinstance(to_data, dict) else ""
                clean_ai = str(ai).strip()
                clean_owner = str(owner).strip()
                clean_notes = str(notes).strip()
                if clean_ai and clean_owner:
                    conn.execute(
                        """
                        INSERT INTO topic_owners (meeting_number, ai_number, owner_name, notes)
                        VALUES (?, ?, ?, ?)
                        ON CONFLICT(meeting_number, ai_number) DO UPDATE SET
                            owner_name = excluded.owner_name,
                            notes = excluded.notes
                        """,
                        (m_num, clean_ai, clean_owner, clean_notes),
                    )
                    conn.execute(
                        """
                        INSERT INTO default_topic_owners (ai_number, owner_name, notes)
                        VALUES (?, ?, ?)
                        ON CONFLICT(ai_number) DO UPDATE SET
                            owner_name = excluded.owner_name,
                            notes = excluded.notes
                        """,
                        (clean_ai, clean_owner, clean_notes),
                    )
                    imported_count += 1

            # 3. Appointments (session-level overrides for this meeting only)
            for sk, ap_data in data.get("appointments", {}).items():
                if isinstance(ap_data, dict):
                    app_del = ap_data.get("appointed_delegate", "").strip()
                    orig_ow = ap_data.get("original_owner", "").strip()
                    reason = ap_data.get("reason", "").strip()
                else:
                    app_del = str(ap_data).strip()
                    orig_ow = ""
                    reason = ""
                if sk and app_del:
                    conn.execute(
                        """
                        INSERT INTO session_appointments (meeting_number, slot_key, appointed_delegate, original_owner, reason)
                        VALUES (?, ?, ?, ?, ?)
                        ON CONFLICT(meeting_number, slot_key) DO UPDATE SET
                            appointed_delegate = excluded.appointed_delegate,
                            original_owner = excluded.original_owner,
                            reason = excluded.reason
                        """,
                        (m_num, str(sk).strip(), app_del, orig_ow, reason),
                    )
                    imported_count += 1

            self._flush_global_defaults_conn(conn)

        return imported_count
