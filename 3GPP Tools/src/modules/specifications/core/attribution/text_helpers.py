# --- File: src/modules/specifications/core/attribution/text_helpers.py ---
"""Stopword constants, diff keyword salience extraction, and ancestor context resolution."""

import collections
import logging
import re
from typing import List, Set, Tuple

logger = logging.getLogger(__name__)

# Stopwords for telecommunications title and diff keyword salience
BOILERPLATE_STOPWORDS: Set[str] = {
    "a", "about", "above", "after", "again", "against", "all", "am", "an", "and", "any", "are",
    "as", "at", "be", "because", "been", "before", "being", "below", "between", "both", "but",
    "by", "can", "clause", "clauses", "cr", "did", "do", "does", "doing", "down", "during",
    "each", "editor", "editors", "few", "first", "for", "from", "further", "had", "has", "have",
    "having", "he", "her", "here", "hers", "herself", "him", "himself", "his", "how", "i", "if",
    "in", "into", "is", "it", "its", "itself", "just", "key", "may", "me", "more", "most", "my",
    "myself", "next", "no", "nor", "not", "note", "notes", "now", "of", "off", "on", "once",
    "only", "or", "other", "our", "ours", "ourselves", "out", "over", "own", "pcr", "s", "same",
    "shall", "she", "should", "so", "solution", "solutions", "some", "such", "t", "table", "than",
    "that", "the", "their", "theirs", "them", "themselves", "then", "there", "these", "they",
    "this", "those", "through", "to", "too", "tr", "ts", "under", "until", "up", "update",
    "updates", "very", "was", "we", "were", "what", "when", "where", "which", "while", "who",
    "whom", "why", "will", "with", "would", "you", "your", "yours", "yourself", "yourselves",
}


class AttributionTextMixin:
    """Mixin providing token cleaning, Key Issue extraction, and ancestor lookup."""

    @staticmethod
    def clean_tokens(text: str) -> List[str]:
        """Extracts lowercase alphabetic/alphanumeric tokens, discarding common stopwords."""
        raw_words = re.findall(r"\b[A-Za-z0-9_-]{2,}\b", text.lower())
        return [w for w in raw_words if w not in BOILERPLATE_STOPWORDS and not w.isdigit()]

    @staticmethod
    def extract_key_issue_references(text: str) -> List[str]:
        """Detects Key Issue references (e.g. 'Key Issue #4', 'KI #4', 'KI5') in clause text."""
        matches = re.findall(r"(?:Key\s+Issue|KI)\s*#?\s*(\d+)", text, re.IGNORECASE)
        unique_nums = sorted(list(set(matches)), key=lambda x: int(x))
        return [f"KI #{num}" for num in unique_nums]

    @staticmethod
    def extract_salient_diff_keywords(diff_text: str, max_keywords: int = 15) -> List[str]:
        """Extracts high-entropy technical terms and acronyms from added (+) diff lines."""
        added_lines = [
            line[1:].strip() for line in diff_text.splitlines()
            if line.startswith("+") and not line.startswith("+++")
        ]
        joined = " ".join(added_lines)
        if not joined:
            return []

        # Find telecom tokens: hyphenated identifiers, all-caps acronyms, camelCase, and words >= 4 chars
        tokens = re.findall(r"\b[A-Za-z0-9]+(?:-[A-Za-z0-9]+)+\b|\b[A-Z]{2,}\b|\b[a-zA-Z0-9_]{4,}\b", joined)
        filtered = [
            t for t in tokens
            if t.lower() not in BOILERPLATE_STOPWORDS and not t.isdigit()
        ]
        # Count frequencies and return most frequent distinctive tokens
        counter = collections.Counter(filtered)
        return [tok for tok, _ in counter.most_common(max_keywords)]

    def resolve_ancestor_context(
        self,
        spec_number: str,
        version: str,
        clause_number: str,
    ) -> Tuple[str, List[str]]:
        """
        Walks up the clause tree using 3gpp_spec_search.db to resolve ancestor titles.
        Returns (parent_root_title, list_of_ancestor_titles).
        """
        if not self.search_db or not hasattr(self.search_db, "get_connection"):
            return "", []

        ancestor_titles: List[str] = []
        curr_clause = clause_number.strip()

        try:
            with self.search_db.get_connection() as conn:
                while "." in curr_clause:
                    parts = curr_clause.rsplit(".", 1)
                    parent_clause = parts[0]
                    row = conn.execute(
                        "SELECT title, text FROM spec_clauses WHERE spec = ? AND version = ? AND clause = ? LIMIT 1",
                        (spec_number, version, parent_clause),
                    ).fetchone()
                    if row and row[0]:
                        t_str = row[0].strip()
                        if t_str and t_str not in ancestor_titles:
                            ancestor_titles.append(t_str)
                    curr_clause = parent_clause
        except Exception as e:
            logger.debug(f"Could not resolve ancestor context for {clause_number}: {e}")

        root_title = ancestor_titles[-1] if ancestor_titles else ""
        return root_title, ancestor_titles
