# --- File: src/modules/nas/core/db/ingestion.py ---
"""Atomic ingestion of parsed NAS/ASN.1 specification versions into 3gpp_protocol_data.db."""

from typing import Any, Dict, List


class NasDbIngestionMixin:
    """Mixin inserting parsed protocol messages, IEs, and Clause 9/ASN.1 definitions."""

    def insert_parsed_spec(
        self,
        spec_number: str,
        version: str,
        messages: List[Dict[str, Any]],
        ie_defs: List[Dict[str, Any]],
        spec_type: str = "NAS",
    ) -> bool:
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "DELETE FROM spec_versions WHERE spec_number = ? AND version = ?",
                    (spec_number, version),
                )

                cursor.execute(
                    "INSERT INTO spec_versions (spec_number, version, spec_type) VALUES (?, ?, ?)",
                    (spec_number, version, spec_type),
                )
                version_id = cursor.lastrowid

                for msg in messages:
                    cursor.execute(
                        """
                        INSERT INTO nas_messages (version_id, clause, message_name, table_caption, description, direction)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """,
                        (
                            version_id,
                            msg.get("clause", ""),
                            msg.get("message_name", ""),
                            msg.get("table_caption", ""),
                            msg.get("description", ""),
                            msg.get("direction", ""),
                        ),
                    )
                    message_id = cursor.lastrowid

                    ie_rows = []
                    for idx, ie in enumerate(msg.get("ies", [])):
                        ie_rows.append((
                            message_id,
                            ie.get("iei", ""),
                            ie.get("information_element", ""),
                            ie.get("field_path", ie.get("information_element", "")),
                            ie.get("depth", 0),
                            ie.get("type_reference", ""),
                            ie.get("presence", ""),
                            ie.get("format", ""),
                            ie.get("length", ""),
                            idx,
                            ie.get("applicability", ""),
                            ie.get("field_description", ""),
                        ))

                    cursor.executemany(
                        """
                        INSERT INTO message_ies (message_id, iei, ie_name, field_path, depth, type_reference, presence, format, length, order_index, applicability, field_description)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                        ie_rows,
                    )

                def_rows = [
                    (
                        version_id,
                        d.get("clause", ""),
                        d.get("ie_name", ""),
                        d.get("raw_description", ""),
                        d.get("structure_table", ""),
                    )
                    for d in ie_defs
                ]
                cursor.executemany(
                    """
                    INSERT OR REPLACE INTO ie_definitions (version_id, clause, ie_name, raw_description, structure_table)
                    VALUES (?, ?, ?, ?, ?)
                """,
                    def_rows,
                )
            return True
        except Exception as e:
            self.logger.error(f"Failed to insert parsed spec TS {spec_number} v{version}: {e}")
            return False
