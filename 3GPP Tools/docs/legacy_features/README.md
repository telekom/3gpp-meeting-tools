# Legacy Features & Gap Analysis Catalog

This directory contains in-depth architectural and technical specifications for each functionality originally implemented in the legacy **3GPP Meeting Helper** (`..\3GPP Meeting Helper`) that is not yet fully ported to the modernized **3GPP Tools** desktop environment.

Each document provides a comprehensive analysis covering:
1. **Executive Summary & Functional Objective**
2. **Deep Code Analysis of Legacy Implementation** (Exact files, functions, lines of code, Win32 COM APIs, regexes, data structures)
3. **Current State & Gap in Modern Codebase**
4. **Target Modern Architecture & Design Specification** (PyQt5 classes, threading models, headless OpenXML / python-docx engines, UI integration)
5. **Architectural & Sequence / Workflow Diagrams** (Mermaid)
6. **Step-by-Step Implementation Roadmap & Acceptance Criteria**

---

## Master Catalog & Status Matrix

| ID | Specification Document | Feature Domain | Complexity | Priority | Status |
| :---: | :--- | :--- | :---: | :---: | :---: |
| **F-01** | [F01: Windows `tdoc://` URL Protocol Handler](F01_windows_tdoc_url_handler.md) | OS Integration | Low | 🟢 High | Specification Ready |
| **F-02** | [F02: Batch CR Cover Sheet Parser to Excel](F02_batch_cr_cover_sheet_parser.md) | Word Automation | Medium | 🟢 High | Specification Ready |
| **F-03** | [F03: PowerPoint Status Report & Presentation Merger](F03_powerpoint_presentation_merger.md) | Office Automation | Low | 🟡 Medium | Specification Ready |
| **F-04** | [F04: Bookmarked Multi-TDoc PDF Dossier & LLM Prompt](F04_bookmarked_pdf_dossier_llm_prompt.md) | PDF & AI Tools | Medium | 🟡 Medium | Specification Ready |
| **F-05** | [F05: "Active Word vs. TDoc" Direct Comparison](F05_active_word_vs_tdoc_comparison.md) | Word Automation | Low | 🟢 High | Specification Ready |
| **F-06** | [F06: MIP / Purview Sensitivity Label Setter](F06_mip_sensitivity_label_setter.md) | Security & DLP | Medium | 🟡 Medium | Specification Ready |
| **F-07** | [F07: Track Changes Revision Author Replacement](F07_track_changes_author_renamer.md) | Word Automation | Low | ⚪ Low | Specification Ready |
| **F-08** | [F08: Word Process Watchdog & Force Killer](F08_word_process_watchdog_killer.md) | System Watchdog | Low | ⚪ Low | Specification Ready |
| **F-09** | [F09: Word Meeting Contribution Summary Report (.docx)](F09_word_meeting_contribution_report.md) | Word Automation | High | ⚪ Low | Specification Ready |
| **F-10** | [F10: "Accept All Changes" Clean Word Export Option](F10_word_accept_all_changes_export.md) | Word Automation | Low | ⚪ Low | Specification Ready |
| **F-11** | [F11: Bi-directional Excel AutoFilter Synchronization](F11_excel_bidirectional_autofilter_sync.md) | Excel Bridge | Medium | ⚪ Low | Specification Ready |
| **F-12** | [F12: Multi-Company Review Comments Matrix Importer](F12_company_comments_matrix_importer.md) | Review Collaboration | Medium | 🟡 Medium | Specification Ready |
| **F-13** | [F13: Native Excel COM Pivot Tables & Charts](F13_native_excel_pivot_tables_charts.md) | Excel Analytics | Medium | ⚪ Low | Specification Ready |
| **F-14** | [F14: Reflector `Attachment.txt` Redirect Link Fetcher](F14_reflector_attachment_link_fetcher.md) | Email Reflector | Low | 🟡 Medium | Specification Ready |
| **F-15** | [F15: Chairman Email Discussion Notes Parser (`<<START>>...<<END>>`)](F15_chairman_email_notes_parser.md) | Email Reflector | Low | ⚪ Low | Specification Ready |
| **F-16** | [F16: Email Approval Excel Summary](F16_email_approval_excel_export.md) | Email Reflector | Medium | ⚪ Low | Specification Ready |
| **F-17** | [F17: 3GPP Guru Web Viewer Integration](F17_3gpp_guru_web_viewer_integration.md) | Web Services | Very Low | 🟢 High | Specification Ready |
| **F-18** | [F18: 3GPP Forge APIs Direct Launcher](F18_3gpp_forge_apis_launcher.md) | Web Services | Very Low | 🟢 High | Specification Ready |
| **F-19** | [F19: Chairman's Notes Word Table Parser (Experimental / WIP)](F19_chairmans_notes_word_table_parser.md) | Meeting Live Notes | High | 🟡 Medium | Specification Ready |

---

## Recommended Staged Implementation Roadmap

```mermaid
flowchart TD
    subgraph Phase1["Phase 1: High-Impact Ergonomics & Quick Wins (Sprint 1)"]
        P1_1["F-01: Windows tdoc:// Protocol Handler"]
        P1_2["F-17: 3GPP Guru Deep Links"]
        P1_3["F-18: 3GPP Forge APIs Launcher"]
        P1_4["F-05: Compare vs. Active Word Document"]
    end

    subgraph Phase2["Phase 2: Office Batch Automation & Exporters (Sprint 2)"]
        P2_1["F-02: Batch CR Cover Sheet Parser to Excel"]
        P2_2["F-03: PowerPoint Presentation Merger"]
        P2_3["F-04: Bookmarked PDF Dossier & LLM Prompt"]
        P2_4["F-07: Track Changes Author Consolidator"]
        P2_5["F-10: Accept All Changes Clean PDF Export"]
    end

    subgraph Phase3["Phase 3: Specialized Workflows & Enterprise Compliance (Sprint 3)"]
        P3_1["F-06: MIP / Purview Sensitivity Label Setter"]
        P3_2["F-12: Multi-Company Comment Matrix Importer"]
        P3_3["F-14: Reflector Attachment.txt Link Fetcher"]
        P3_4["F-15: Chairman <<START>> Delimiter Parser"]
        P3_5["F-19: Chairman's Notes Word Table Parser"]
    end

    Phase1 --> Phase2 --> Phase3
```
