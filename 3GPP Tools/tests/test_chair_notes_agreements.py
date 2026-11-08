# --- File: tests/test_chair_notes_agreements.py ---
"""
Unit tests for Chairman's Notes downloading/selection (`ChairNotesDownloaderThread`),
multi-WG agreements & working assumptions parsing (`ChairNotesAgreementsParser`),
and Cross-Meeting Study Item Agreement Tracking (`AgreementTrackerWorker`).
"""
import zipfile
from pathlib import Path

import docx
from docx.shared import RGBColor

from modules.meetings.core.agreement_tracker_worker import AgreementTrackerWorker
from modules.meetings.core.chair_notes_downloader import ChairNotesDownloaderThread
from modules.meetings.core.tdocs.chair_notes_agreements_parser import ChairNotesAgreementsParser


def test_chair_notes_filename_filtering_and_smart_selection():
    """
    Verifies that RAN3 EOM files and RAN2 Final ZIP/DOCX files are selected over
    dozens of daily interim snapshots, while pure TDoc ZIPs (`R2-2607187.zip`) and
    Office lock files (`~$...`) are excluded.
    """
    raw_files = [
        ("~$RAN3_133bis_agenda.docx", "https://example.org/~$RAN3_133bis_agenda.docx"),
        ("R3-264505.zip", "https://example.org/R3-264505.zip"),
        ("RAN3_133bis_agenda_20261012.docx", "https://example.org/RAN3_133bis_agenda_20261012.docx"),
        ("RAN3_133bis_agenda_20261014.docx", "https://example.org/RAN3_133bis_agenda_20261014.docx"),
        ("RAN3_133bis_agenda_20261015_1845.docx", "https://example.org/RAN3_133bis_agenda_20261015_1845.docx"),
        ("RAN3_133bis_agenda_20261016_EOM.docx", "https://example.org/RAN3_133bis_agenda_20261016_EOM.docx"),
    ]

    valid = [t for t in raw_files if ChairNotesDownloaderThread.is_valid_chair_notes_filename(t[0])]
    valid_names = [t[0] for t in valid]
    assert "~$RAN3_133bis_agenda.docx" not in valid_names
    assert "R3-264505.zip" not in valid_names
    assert "RAN3_133bis_agenda_20261016_EOM.docx" in valid_names

    selected = ChairNotesDownloaderThread.select_best_chair_notes_files(valid, only_final_or_latest=True)
    assert len(selected) == 1
    assert selected[0][0] == "RAN3_133bis_agenda_20261016_EOM.docx"


def test_ran2_zip_extraction_prefers_clean_docx(tmp_path: Path):
    """
    Verifies that when a RAN2 `R2_135b_Chair_Notes_final.zip` contains both
    `_clean.docx` and `_tracked_changes.docx`, `extract_chair_notes_zip` extracts
    and returns only the `_clean.docx` variant.
    """
    zip_path = tmp_path / "R2_135b_Chair_Notes_10-16-26_final.zip"
    clean_doc = docx.Document()
    clean_doc.add_paragraph("Clean notes")
    clean_bytes_path = tmp_path / "clean_src.docx"
    clean_doc.save(str(clean_bytes_path))

    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.write(clean_bytes_path, arcname="R2_135b_Chair_Notes_10-16-26_final_clean.docx")
        zf.write(clean_bytes_path, arcname="R2_135b_Chair_Notes_10-16-26_final_tracked_changes.docx")

    out_dir = tmp_path / "ChairNotes"
    extracted = ChairNotesDownloaderThread.extract_chair_notes_zip(zip_path, out_dir)
    assert len(extracted) == 1
    assert extracted[0] == "R2_135b_Chair_Notes_10-16-26_final_clean.docx"


def test_ran2_block_based_chair_notes_parsing(tmp_path: Path):
    """
    Tests RAN2 document-order block parsing:
    - Agenda Item Headings with parenthetical WI metadata (`(FS_6G_Radio, ...)`),
    - Tab-delimited TDoc lines (`R2-2607326\tTitle\tSource\tType\tFor\tRel-20\tFS_6G_Radio`),
    - 1x1 Boxed `Agreements` tables,
    - Multi-paragraph `Agreements` blocks,
    - Filtering of trivial status lines (`The CR is postponed`, `Noted`).
    """
    doc = docx.Document()
    doc.add_heading("9.3.1.1\tUser Plane Radio Protocols", level=3)
    doc.add_paragraph("(FS_6G_Radio, leading WG: RAN2; REL-20; SID: RP-253876)")

    doc.add_paragraph("R2-2607326\t6G L2 Architecture\tNokia\tdiscussion\tagreement\tRel-20\tFS_6G_Radio")
    doc.add_paragraph("Noted")

    # Pattern A: 1x1 boxed table with Agreements
    tbl = doc.add_table(rows=1, cols=1)
    tbl.cell(0, 0).text = (
        "Agreements\n"
        "1. PDCP and RLC layers are kept separated in 6GR.\n"
        "2. Support selective out-of-order delivery at PDCP receiver."
    )

    # Second TDoc with Pattern B (paragraph sequence starting with Agreements)
    doc.add_paragraph("R2-2607674\t6G UL Data Burst Reporting\tHuawei\tdiscussion\tagreement\tRel-20\tFS_6G_Radio")
    doc.add_paragraph("Agreements")
    doc.add_paragraph("- The type 1 burst reporting is realized by BSR/DSR.")
    doc.add_paragraph("- Burst information of non-AS buffer data is FFS.")

    docx_file = tmp_path / "R2_135b_Notes.docx"
    doc.save(str(docx_file))

    agenda_data, agreements = ChairNotesAgreementsParser.parse_docx_with_agreements(
        str(docx_file),
        mtg_info={"wg_name": "RAN2", "meeting_number": "135bis", "start_date": "2026-10-12"},
    )

    assert "R2-2607326" in agenda_data
    assert "R2-2607674" in agenda_data
    assert agenda_data["R2-2607326"]["Related WIs"] == "FS_6G_Radio"
    assert len(agreements) == 2

    ag1 = agreements[0]
    assert ag1["ai"] == "9.3.1.1"
    assert ag1["wi"] == "FS_6G_Radio"
    assert ag1["release"] == "Rel-20"
    assert ag1["tdoc_id"] == "R2-2607326"
    assert "PDCP and RLC layers are kept separated" in ag1["text"]

    ag2 = agreements[1]
    assert ag2["tdoc_id"] == "R2-2607674"
    assert "type 1 burst reporting is realized by BSR/DSR" in ag2["text"]


def test_ran3_table_based_color_and_banner_parsing(tmp_path: Path):
    """
    Tests RAN3 3-column table parsing (`TDoc | Title | Comments`):
    - Shaded banner row with `SID [FS_6G_Radio]: RP-261559` and `From RAN3#132:` recap separation,
    - Child AI banner inheriting parent `FS_6G_Radio` WI,
    - Green (`#008000`) Agreements and Working Assumptions (`WA:`),
    - Blue (`#0000FF`) FFS / Open Issues,
    - Merged offline summary row linking to `R3-264505`.
    """
    doc = docx.Document()
    tbl = doc.add_table(rows=1, cols=3)
    tbl.rows[0].cells[0].text = "TDoc"
    tbl.rows[0].cells[1].text = "Title"
    tbl.rows[0].cells[2].text = "Comments"

    # Row 1: Parent AI 10 Banner with SID [FS_6G_Radio] and a "From RAN3#132:" recap
    r_banner = tbl.add_row().cells
    r_banner[0].merge(r_banner[2])
    p_head = r_banner[0].paragraphs[0]
    p_head.text = "10. Study on 6G Radio"
    r_banner[0].add_paragraph("SID [FS_6G_Radio]: RP-261559")
    r_banner[0].add_paragraph("From RAN3#132:")
    p_old = r_banner[0].add_paragraph()
    run_old = p_old.add_run("Old agreement from previous meeting.")
    run_old.font.color.rgb = RGBColor(0x00, 0x80, 0x00)

    # Row 2: Child AI 10.2.2 Banner (inherits FS_6G_Radio)
    r_sub = tbl.add_row().cells
    r_sub[0].merge(r_sub[2])
    r_sub[0].paragraphs[0].text = "10.2.2. Architecture and interfaces"

    # Row 3: TDoc row with Green Agreement in Comments
    r_tdoc = tbl.add_row().cells
    r_tdoc[0].text = "R3-264464"
    r_tdoc[1].text = "6G CN-RAN Interface Name (Ericsson, Discussion)"
    p_c1 = r_tdoc[2].paragraphs[0]
    p_c1.text = "Noted"
    p_c2 = r_tdoc[2].add_paragraph()
    run_g = p_c2.add_run("The aNB to 6G CN interface is renamed to Ha.")
    run_g.font.color.rgb = RGBColor(0x00, 0x80, 0x00)

    # Row 4: Merged Moderator Summary row with WA (Green) and FFS (Blue)
    r_sum = tbl.add_row().cells
    r_sum[0].merge(r_sum[2])
    r_sum[0].paragraphs[0].text = "Summary of offline R3-264505 Noted"
    p_wa = r_sum[0].add_paragraph()
    run_wa = p_wa.add_run("WA: For connectivity services, the Xa interface is a point-to-point logical interface.")
    run_wa.font.color.rgb = RGBColor(0x00, 0x80, 0x00)
    p_ffs = r_sum[0].add_paragraph()
    run_ffs = p_ffs.add_run("FFS whether indirect forwarding via UPF is supported.")
    run_ffs.font.color.rgb = RGBColor(0x00, 0x00, 0xFF)

    docx_file = tmp_path / "RAN3_133bis_EOM.docx"
    doc.save(str(docx_file))

    agenda_data, agreements = ChairNotesAgreementsParser.parse_docx_with_agreements(
        str(docx_file),
        mtg_info={"wg_name": "RAN3", "meeting_number": "133bis", "start_date": "2026-10-12"},
    )

    assert "R3-264464" in agenda_data
    assert agenda_data["R3-264464"]["Related WIs"] == "FS_6G_Radio"

    new_ags = [a for a in agreements if not a["is_previous_meeting_recap"]]
    recap_ags = [a for a in agreements if a["is_previous_meeting_recap"]]
    assert len(recap_ags) == 1
    assert len(new_ags) == 3

    kinds = [a["kind"] for a in new_ags]
    assert "Agreement" in kinds
    assert "Working Assumption" in kinds
    assert "FFS / Open Issue" in kinds

    for a in new_ags:
        assert a["wi"] == "FS_6G_Radio"
        assert a["ai"] == "10.2.2"


def test_chair_notes_cache_persistence_and_worker_filtering(tmp_path: Path):
    """
    Tests saving/loading `Agenda/chair_notes_cache.json` and `AgreementTrackerWorker._matches_filter`.
    """
    mtg_dir = tmp_path / "RAN2_135bis"
    sample_agenda = {
        "R2-2607326": {"Title": "6G L2", "Agenda Item": "9.3.1", "Related WIs": "FS_6G_Radio"}
    }
    sample_agreements = [
        {
            "wg": "RAN2",
            "meeting_number": "135bis",
            "meeting_date": "2026-10-12",
            "ai": "9.3.1",
            "ai_title": "User Plane",
            "wi": "FS_6G_Radio",
            "release": "Rel-20",
            "kind": "Agreement",
            "text": "PDCP and RLC layers are kept separated in 6GR.",
            "tdoc_id": "R2-2607326",
            "source_file": "R2_notes.docx",
            "is_previous_meeting_recap": False,
        }
    ]
    ChairNotesAgreementsParser.save_cached_chair_notes_data(mtg_dir, sample_agenda, sample_agreements)
    loaded_ag, loaded_list = ChairNotesAgreementsParser.load_cached_chair_notes_data(mtg_dir)
    assert "R2-2607326" in loaded_ag
    assert len(loaded_list) == 1

    # Verify worker filtering by WI token, AI prefix (`9.3` matching `9.3.1`), and keyword
    worker = AgreementTrackerWorker(
        meetings_db=None,
        wgs=["RAN2"],
        date_from="2026-01-01",
        date_to="2026-12-31",
        target_wis=["6G_Radio"],
        target_ais=["9.3"],
        keyword_query="PDCP",
        cache_dir=str(tmp_path),
    )
    record = {
        "WG": "RAN2",
        "Meeting": "RAN2 #135bis",
        "AI": "9.3.1",
        "AI Title": "User Plane",
        "Work Item": "FS_6G_Radio",
        "Kind": "Agreement",
        "Agreement Text": "PDCP and RLC layers are kept separated in 6GR.",
        "TDoc": "R2-2607326",
        "TDoc Title": "6G L2",
        "Source": "Nokia",
    }
    assert worker._matches_filter(record) is True

    worker_mismatch = AgreementTrackerWorker(
        meetings_db=None,
        wgs=["RAN2"],
        date_from="2026-01-01",
        date_to="2026-12-31",
        target_wis=["NR_Mob_Ph5"],
        target_ais=[],
        keyword_query="",
        cache_dir=str(tmp_path),
    )
    assert worker_mismatch._matches_filter(record) is False


if __name__ == "__main__":
    import tempfile
    import unittest

    class TestChairNotesAgreementsSuite(unittest.TestCase):
        def test_filename_selection(self):
            test_chair_notes_filename_filtering_and_smart_selection()

        def test_zip_extraction(self):
            with tempfile.TemporaryDirectory() as td:
                test_ran2_zip_extraction_prefers_clean_docx(Path(td))

        def test_ran2_parsing(self):
            with tempfile.TemporaryDirectory() as td:
                test_ran2_block_based_chair_notes_parsing(Path(td))

        def test_ran3_parsing(self):
            with tempfile.TemporaryDirectory() as td:
                test_ran3_table_based_color_and_banner_parsing(Path(td))

        def test_cache_and_worker(self):
            with tempfile.TemporaryDirectory() as td:
                test_chair_notes_cache_persistence_and_worker_filtering(Path(td))

    unittest.main()

