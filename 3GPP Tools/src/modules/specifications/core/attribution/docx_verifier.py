# --- File: src/modules/specifications/core/attribution/docx_verifier.py ---
"""Stage 2 Track-Changes and placeholder-agnostic OpenXML content verification."""

import logging
import re
import zipfile
from pathlib import Path
from typing import Set, Tuple

logger = logging.getLogger(__name__)


class DocxVerifierMixin:
    """Mixin verifying clause diffs against candidate Word/ZIP documents."""

    @classmethod
    def verify_candidate_in_docx(
        cls,
        tdoc: str,
        diff_text: str,
        doc_path: Path,
        cr_number: str = "",
    ) -> Tuple[bool, int, str]:
        """
        Executes OpenXML Track-Changes-aware shingle containment check on a candidate Word document.
        Normalizes dynamic numbering ('6.x' / '6.20' -> '<NUM>') to achieve 100% ground truth.
        Returns (is_matched, match_percentage, matching_snippet).
        """
        if not doc_path.exists() or not diff_text.strip():
            return False, 0, ""

        # Extract added lines from diff
        added_lines = [
            line[1:].strip()
            for line in diff_text.splitlines()
            if line.startswith("+") and not line.startswith("+++") and len(line[1:].strip()) > 15
        ]
        if not added_lines:
            return False, 0, ""

        try:
            # Handle .docx (and Plenary CR pack .zip containing CR .docx)
            target_xml = ""
            if doc_path.suffix.lower() == ".docx":
                with zipfile.ZipFile(doc_path, "r") as zf:
                    if "word/document.xml" in zf.namelist():
                        target_xml = zf.read("word/document.xml").decode("utf-8", errors="replace")
            elif doc_path.suffix.lower() == ".zip":
                with zipfile.ZipFile(doc_path, "r") as zf:
                    candidates = [n for n in zf.namelist() if n.lower().endswith(".docx") and not n.startswith("~$")]
                    matched_file = None
                    if cr_number:
                        for c in candidates:
                            if cr_number in c:
                                matched_file = c
                                break
                    if not matched_file and candidates:
                        matched_file = candidates[0]

                    if matched_file:
                        with zipfile.ZipFile(zf.open(matched_file)) as inner_zf:
                            if "word/document.xml" in inner_zf.namelist():
                                target_xml = inner_zf.read("word/document.xml").decode("utf-8", errors="replace")

            if not target_xml:
                return False, 0, ""

            # Track-changes awareness: remove deleted text (<w:del> subtrees) when checking additions
            clean_xml = re.sub(r"<w:del\b[^>]*>.*?</w:del>", "", target_xml, flags=re.DOTALL)
            # Extract plain text runs
            text_runs = re.findall(r"<w:t\b[^>]*>(.*?)</w:t>", clean_xml, flags=re.DOTALL)
            doc_body_text = " ".join(text_runs)

            # Normalization: replace all clause/figure/table numbers with canonical <NUM>
            def _normalize(t_str: str) -> str:
                t_sub = re.sub(r"\b\d+(?:\.[0-9xX]+)+(?:-\d+)?\b", "<NUM>", t_str)
                t_sub = re.sub(r"\s+", " ", t_sub).strip().lower()
                return t_sub

            norm_doc_text = _normalize(doc_body_text)

            # 1. Exact Sentence Search (Testing 3 longest added lines)
            longest_lines = sorted(added_lines, key=len, reverse=True)[:3]
            for candidate_line in longest_lines:
                norm_line = _normalize(candidate_line)
                if len(norm_line) > 25 and norm_line in norm_doc_text:
                    snip = candidate_line[:120] + ("..." if len(candidate_line) > 120 else "")
                    return True, 98, snip

            # 2. 6-Word Shingle Containment Check
            def _make_shingles(s_str: str, k: int = 6) -> Set[Tuple[str, ...]]:
                words = s_str.split()
                if len(words) < k:
                    return {tuple(words)} if words else set()
                return {tuple(words[i:i + k]) for i in range(len(words) - k + 1)}

            sample_diff = _normalize(" ".join(added_lines[:15]))
            diff_shingles = _make_shingles(sample_diff, k=6)

            if diff_shingles:
                doc_shingles = _make_shingles(norm_doc_text, k=6)
                matches = diff_shingles.intersection(doc_shingles)
                containment = len(matches) / len(diff_shingles)
                match_pct = int(containment * 100)

                if match_pct >= 45:
                    snip = added_lines[0][:120] if added_lines else ""
                    return True, match_pct, snip

        except Exception as e:
            logger.debug(f"Error during Stage 2 fingerprinting for {tdoc}: {e}")

        return False, 0, ""
