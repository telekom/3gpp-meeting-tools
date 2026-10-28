# --- File: tests/test_dossier_exporter.py ---
"""
test_dossier_exporter.py - Test suite for the Unified Navigable Dossier Exporter (F-04).
Tests HTML synthesis, vector PDF rendering with hierarchical outlines, and companion AI prompts.
"""

import os
import sys
import tempfile
from pathlib import Path

import pypdf
from PyQt5.QtWidgets import QApplication

# Ensure QApplication exists for QTextDocument / QPrinter
app = QApplication.instance() or QApplication(sys.argv)

from modules.meetings.core.dossier_exporter import DossierExporterThread
from modules.meetings.ui.dossier_export_dialog import DossierExportDialog


def test_dossier_export_dialog_config():
    """Tests DossierExportDialog initialization and config extraction."""
    with tempfile.TemporaryDirectory() as tmpdir:
        meeting_dir = Path(tmpdir)
        mtg_info = {"wg_name": "SA2", "meeting_number": "162"}
        dlg = DossierExportDialog(None, meeting_dir, tdocs_count=5, mtg_info=mtg_info, scope_label="selected")

        cfg = dlg.get_config()
        assert cfg["export_format"] == "both"
        assert cfg["include_prompt"] is True
        assert cfg["embed_base64_images"] is True
        assert cfg["download_missing"] is True
        assert "Dossier_SA2_162" in str(cfg["output_dir"])


def test_dossier_exporter_html_synthesis():
    """Tests HTML synthesis with diff highlighting, change markers, and metadata badges."""
    with tempfile.TemporaryDirectory() as tmpdir:
        meeting_dir = Path(tmpdir)
        media_export_dir = meeting_dir / "media"
        media_export_dir.mkdir(parents=True, exist_ok=True)

        docs = [
            {
                "tdoc_id": "S2-240101",
                "title": "Clarification on UPF Selection",
                "source": "Deutsche Telekom",
                "type": "pCR",
                "status": "Agreed",
                "ai": "6.2",
                "md_text": (
                    "# S2-240101\n\n"
                    "## 1. Reason for Change\n"
                    "Clarify UPF selection procedures.\n\n"
                    "* * * * Start of changes * * * *\n\n"
                    "The SMF shall select [INSERTED: an appropriate] UPF based on [DELETED: old] criteria.\n\n"
                    "* * * * End of changes * * * *\n"
                ),
                "media_dir": media_export_dir,
                "tdoc_url": "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/Docs/S2-240101.zip",
            },
            {
                "tdoc_id": "S2-240102",
                "title": "Redirection Flow Enhancements",
                "source": "Qualcomm Incorporated",
                "type": "CR",
                "status": "Agreed",
                "ai": "6.2",
                "md_text": (
                    "# S2-240102\n\n"
                    "## 1. Discussion\n"
                    "Proposed redirection flow.\n"
                ),
                "media_dir": media_export_dir,
                "tdoc_url": "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/Docs/S2-240102.zip",
            },
        ]

        exporter = DossierExporterThread(
            meeting_dir=meeting_dir,
            tdocs_list=docs,
            docs_ftp_url="https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/Docs/",
            revisions_url="",
            mtg_info={"wg_name": "SA2", "meeting_number": "162"},
            output_dir=meeting_dir / "out",
        )

        html_content, tdoc_map = exporter._build_html_dossier(docs, media_export_dir)

        # Check essential structural tags
        assert "<title>3GPP Dossier - SA2 #162</title>" in html_content
        assert "floating-nav" in html_content
        assert "master-index" in html_content
        assert "Master Table of Contents" in html_content
        assert "S2-240101" in html_content
        assert "S2-240102" in html_content
        assert "Clarification on UPF Selection" in html_content
        assert "Deutsche Telekom" in html_content
        assert "Qualcomm Incorporated" in html_content

        # Check badges
        assert "badge-pcr" in html_content
        assert "badge-cr" in html_content
        assert "badge-agreed" in html_content

        # Check diff highlighting
        assert '<ins class="diff-ins">an appropriate</ins>' in html_content
        assert '<del class="diff-del">old</del>' in html_content

        # Check 3GPP change markers
        assert '<div class="change-marker">' in html_content


def test_dossier_exporter_pdf_and_bookmarks():
    """Tests vector PDF generation and pypdf hierarchical bookmark outline injection."""
    with tempfile.TemporaryDirectory() as tmpdir:
        out_dir = Path(tmpdir)
        pdf_path = out_dir / "test_dossier.pdf"
        media_export_dir = out_dir / "media"
        media_export_dir.mkdir(parents=True, exist_ok=True)

        docs = [
            {
                "tdoc_id": "S2-240101",
                "title": "UPF Selection Enhancements",
                "source": "Deutsche Telekom",
                "type": "pCR",
                "status": "Agreed",
                "ai": "6.2",
                "md_text": "# S2-240101\n\nContent for UPF selection...\n" + ("Paragraph text.\n\n" * 5),
                "media_dir": media_export_dir,
                "tdoc_url": "",
            },
            {
                "tdoc_id": "S2-240102",
                "title": "Redirection Flow",
                "source": "Qualcomm",
                "type": "CR",
                "status": "Agreed",
                "ai": "6.2",
                "md_text": "# S2-240102\n\nContent for redirection flow...\n" + ("Paragraph text.\n\n" * 5),
                "media_dir": media_export_dir,
                "tdoc_url": "",
            },
            {
                "tdoc_id": "S2-240201",
                "title": "RedCap Paging Optimizations",
                "source": "Ericsson",
                "type": "Discussion",
                "status": "Noted",
                "ai": "8.1",
                "md_text": "# S2-240201\n\nContent for RedCap paging...\n" + ("Paragraph text.\n\n" * 5),
                "media_dir": media_export_dir,
                "tdoc_url": "",
            },
        ]

        exporter = DossierExporterThread(
            meeting_dir=out_dir,
            tdocs_list=docs,
            docs_ftp_url="",
            revisions_url="",
            mtg_info={"wg_name": "SA2", "meeting_number": "162"},
            output_dir=out_dir,
        )

        html_content, _ = exporter._build_html_dossier(docs, media_export_dir)
        page_spans = exporter._render_bookmarked_pdf(html_content, pdf_path, docs)

        assert pdf_path.exists()
        assert pdf_path.stat().st_size > 0

        # Verify page spans
        assert "S2-240101" in page_spans
        assert "S2-240102" in page_spans
        assert "S2-240201" in page_spans
        for tid, (st, en) in page_spans.items():
            assert 1 <= st <= en

        # Verify pypdf outline bookmarks
        reader = pypdf.PdfReader(str(pdf_path))
        assert len(reader.pages) >= 2
        assert len(reader.outline) >= 2  # Master index + Agenda item groups

        # Inspect outlines
        outline_titles = []
        for item in reader.outline:
            if isinstance(item, list):
                for sub in item:
                    outline_titles.append(sub.title)
            else:
                outline_titles.append(item.title)

        combined_titles = " ".join(outline_titles)
        assert "Master Index" in combined_titles
        assert "Agenda Item 6.2" in combined_titles
        assert "Agenda Item 8.1" in combined_titles
        assert "S2-240101" in combined_titles
        assert "S2-240102" in combined_titles
        assert "S2-240201" in combined_titles


def test_dossier_exporter_ai_prompt():
    """Tests companion AI prompt generation with exact page spans and 3GPP evaluation directives."""
    with tempfile.TemporaryDirectory() as tmpdir:
        out_dir = Path(tmpdir)
        prompt_path = out_dir / "dossier_prompt.txt"

        docs = [
            {
                "tdoc_id": "S2-240101",
                "title": "Clarification on UPF Selection",
                "source": "Deutsche Telekom",
                "type": "pCR",
                "status": "Agreed",
                "ai": "6.2",
            },
            {
                "tdoc_id": "S2-240102",
                "title": "Redirection Flow Enhancements",
                "source": "Qualcomm",
                "type": "CR",
                "status": "Agreed",
                "ai": "6.2",
            },
        ]
        page_spans = {
            "S2-240101": (2, 5),
            "S2-240102": (6, 9),
        }

        exporter = DossierExporterThread(
            meeting_dir=out_dir,
            tdocs_list=docs,
            docs_ftp_url="",
            revisions_url="",
            mtg_info={"wg_name": "SA2", "meeting_number": "162"},
            output_dir=out_dir,
        )

        exporter._generate_ai_prompt(docs, page_spans, prompt_path)

        assert prompt_path.exists()
        prompt_text = prompt_path.read_text(encoding="utf-8")

        # Verify page map
        assert "S2-240101" in prompt_text
        assert "pages 2 to 5" in prompt_text
        assert "Deutsche Telekom" in prompt_text
        assert "S2-240102" in prompt_text
        assert "pages 6 to 9" in prompt_text

        # Verify 3GPP evaluation rules
        assert "Change Requests (CRs)" in prompt_text
        assert "Reason for change" in prompt_text
        assert "Pseudo-CRs (pCRs)" in prompt_text
        assert "Discussion Papers" in prompt_text
        assert "Liaison Statements" in prompt_text
        assert "notebooklm" in prompt_text.lower() or "claude" in prompt_text.lower()


def test_dossier_exporter_full_run():
    """Tests full synchronous run() of DossierExporterThread creating all deliverables."""
    with tempfile.TemporaryDirectory() as tmpdir:
        meeting_dir = Path(tmpdir)
        out_dir = meeting_dir / "Dossier_Export"

        docs = [
            {
                "TDoc": "S2-240101",
                "Title": "UPF Selection",
                "Source": "Deutsche Telekom",
                "Type": "pCR",
                "TDoc Status": "Agreed",
                "Agenda Item": "6.2",
            },
            {
                "TDoc": "S2-240102",
                "Title": "Redirection Flow",
                "Source": "Qualcomm",
                "Type": "CR",
                "TDoc Status": "Agreed",
                "Agenda Item": "6.2",
            },
        ]

        # Pre-seed local markdown file for S2-240101
        tdoc_folder = meeting_dir / "S2-240101"
        tdoc_folder.mkdir(parents=True, exist_ok=True)
        (tdoc_folder / "S2-240101.md").write_text("# S2-240101\n\nPre-seeded markdown content.", encoding="utf-8")

        exporter = DossierExporterThread(
            meeting_dir=meeting_dir,
            tdocs_list=docs,
            docs_ftp_url="",
            revisions_url="",
            mtg_info={"wg_name": "SA2", "meeting_number": "162"},
            export_format="both",
            include_prompt=True,
            embed_base64_images=True,
            download_missing=False,
            output_dir=out_dir,
        )

        results = {}
        def on_finished(success, msg, files):
            results["success"] = success
            results["msg"] = msg
            results["files"] = files

        exporter.finished.connect(on_finished)
        exporter.run()

        assert results.get("success") is True
        assert (out_dir / "dossier.html").exists()
        assert (out_dir / "dossier.pdf").exists()
        assert (out_dir / "dossier_prompt.txt").exists()
        assert (out_dir / "media").is_dir()
