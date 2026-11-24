"""User annotations (tags and comments) persistence for specifications."""

from __future__ import annotations

from datetime import datetime
from typing import Dict, List, Set, Union


class SpecsAnnotationsMixin:
    """CRUD operations for per-specification user tags and comments."""

    def get_spec_annotations(self, spec_number: str) -> dict:
        """Returns user tags and comments for a specification."""
        clean_spec = (spec_number or "").strip()
        if not clean_spec:
            return {"tags": [], "comments": "", "updated_at": ""}
        query = "SELECT tags, comments, updated_at FROM spec_user_annotations WHERE spec_number = ?"
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, (clean_spec,))
                row = cursor.fetchone()
                if not row:
                    return {"tags": [], "comments": "", "updated_at": ""}
                tags_raw = row[0] or ""
                tags_list = [t.strip() for t in tags_raw.split(",") if t.strip()]
                return {
                    "tags": tags_list,
                    "comments": row[1] or "",
                    "updated_at": row[2] or "",
                }
        except Exception as e:
            self.logger.error(f"Error fetching annotations for {clean_spec}: {e}")
            return {"tags": [], "comments": "", "updated_at": ""}

    def get_all_annotations(self) -> Dict[str, dict]:
        """Returns all annotations keyed by spec_number."""
        query = "SELECT spec_number, tags, comments, updated_at FROM spec_user_annotations"
        result = {}
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query)
                for s_num, tags_raw, comments, updated_at in cursor.fetchall():
                    tags_list = [t.strip() for t in (tags_raw or "").split(",") if t.strip()]
                    result[s_num] = {
                        "tags": tags_list,
                        "comments": comments or "",
                        "updated_at": updated_at or "",
                    }
        except Exception as e:
            self.logger.error(f"Error fetching all annotations: {e}")
        return result

    def get_all_unique_tags(self) -> List[str]:
        """Returns a sorted list of all unique tags used across specifications."""
        query = "SELECT tags FROM spec_user_annotations WHERE tags IS NOT NULL AND tags != ''"
        all_tags: Set[str] = set()
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query)
                for (tags_str,) in cursor.fetchall():
                    for t in (tags_str or "").split(","):
                        t_clean = t.strip()
                        if t_clean:
                            all_tags.add(t_clean)
            return sorted(list(all_tags), key=lambda x: x.lower())
        except Exception as e:
            self.logger.error(f"Error fetching unique tags: {e}")
            return []

    def set_spec_annotations(self, spec_number: str, tags: Union[List[str], str], comments: str = "") -> bool:
        """Upserts user tags and comments for a specification."""
        spec_clean = (spec_number or "").strip()
        if not spec_clean:
            return False

        if isinstance(tags, list):
            clean_tags_list = sorted(list({t.strip() for t in tags if t.strip()}))
        else:
            clean_tags_list = sorted(list({t.strip() for t in (tags or "").split(",") if t.strip()}))
        clean_tags = ", ".join(clean_tags_list)

        clean_comments = (comments or "").strip()
        now_iso = datetime.now().isoformat(timespec="seconds")

        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                if not clean_tags and not clean_comments:
                    cursor.execute("DELETE FROM spec_user_annotations WHERE spec_number = ?", (spec_clean,))
                else:
                    cursor.execute("""
                        INSERT INTO spec_user_annotations (spec_number, tags, comments, updated_at)
                        VALUES (?, ?, ?, ?)
                        ON CONFLICT(spec_number) DO UPDATE SET
                            tags = excluded.tags,
                            comments = excluded.comments,
                            updated_at = excluded.updated_at
                    """, (spec_clean, clean_tags, clean_comments, now_iso))
            return True
        except Exception as e:
            self.logger.error(f"Failed to set annotations for {spec_clean}: {e}")
            return False

    def add_tag_to_specs(self, spec_numbers: List[str], tag: str) -> int:
        """Adds a tag to multiple specifications, preserving existing tags and comments."""
        clean_tag = (tag or "").strip()
        if not clean_tag:
            return 0

        count = 0
        for s_num in spec_numbers:
            ann = self.get_spec_annotations(s_num)
            tags = set(ann.get("tags", []))
            if clean_tag not in tags:
                tags.add(clean_tag)
                if self.set_spec_annotations(s_num, list(tags), ann.get("comments", "")):
                    count += 1
        return count

    def remove_tag_from_specs(self, spec_numbers: List[str], tag: str) -> int:
        """Removes a tag from multiple specifications."""
        clean_tag = (tag or "").strip().lower()
        if not clean_tag:
            return 0

        count = 0
        for s_num in spec_numbers:
            ann = self.get_spec_annotations(s_num)
            current_tags = ann.get("tags", [])
            filtered_tags = [t for t in current_tags if t.lower() != clean_tag]
            if len(filtered_tags) != len(current_tags):
                if self.set_spec_annotations(s_num, filtered_tags, ann.get("comments", "")):
                    count += 1
        return count

    def delete_spec_annotations(self, spec_number: str) -> bool:
        """Deletes user annotations for a specification."""
        spec_clean = (spec_number or "").strip()
        if not spec_clean:
            return False
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM spec_user_annotations WHERE spec_number = ?", (spec_clean,))
            return True
        except Exception as e:
            self.logger.error(f"Failed to delete annotations for {spec_clean}: {e}")
            return False
