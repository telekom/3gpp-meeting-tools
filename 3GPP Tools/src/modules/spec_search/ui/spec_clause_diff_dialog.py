# --- File: src/modules/spec_search/ui/spec_clause_diff_dialog.py ---
import html
import logging
from pathlib import Path
from typing import Optional, Any, Tuple

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (
    QApplication,
    QButtonGroup,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from modules.spec_search.core.spec_clause_diff import build_llm_clause_prompt, generate_unified_diff
from modules.spec_search.core.spec_search_db import SpecSearchDatabase

logger = logging.getLogger(__name__)


class SpecClauseDiffDialog(QDialog):
    """
    Unified Clause Comparison & Analysis Dialog.
    Features:
      1. Tab 1 - Visual Diff: In-app color-coded text diff (+ green, - red, @@ headers)
         with change statistics and 1-click clipboard copy.
      2. Tab 2 - AI Analysis Prompt: Multi-tier hierarchical context prompt generator
         tailored for LLM standards/patent/signalling analysis.
      3. Footer - Word Track Changes: 1-click extraction and launch into Microsoft Word
         COM automation with native Track Changes (strikethrough & underline).
    """

    def __init__(
        self,
        db: SpecSearchDatabase,
        spec_number: str,
        current_version: str,
        clause_number: str,
        clause_title: str,
        cache_dir: Optional[Path] = None,
        specs_db: Optional[Any] = None,
        base_version: Optional[str] = None,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.db = db
        self.spec_number = str(spec_number).strip()
        self.current_version = str(current_version).lstrip("v").strip()
        self.clause_number = str(clause_number).strip()
        self.clause_title = str(clause_title).strip()
        self.cache_dir = cache_dir
        self.specs_db = specs_db
        self.initial_base_version = str(base_version).lstrip("v").strip() if base_version else None

        self._raw_diff_text = ""
        self._generated_prompt = ""
        self._word_cmp_worker = None

        self.setWindowTitle(f"🔍 Compare Clause {self.clause_number} — 3GPP TS {self.spec_number}")
        self.resize(980, 720)
        self._setup_ui()
        self._load_available_versions()
        self._generate_diff_preview()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(10)

        # ---------------------------------------------------------------------
        # Top Controls: Base and Target Version Selection
        # ---------------------------------------------------------------------
        config_group = QGroupBox("Release Comparison")
        config_layout = QHBoxLayout(config_group)
        config_layout.setContentsMargins(12, 8, 12, 8)
        config_layout.setSpacing(10)

        config_layout.addWidget(QLabel("<b>Base Release (Old):</b>"))
        self.base_ver_combo = QComboBox()
        self.base_ver_combo.currentIndexChanged.connect(self._generate_diff_preview)
        config_layout.addWidget(self.base_ver_combo, stretch=1)

        arrow_lbl = QLabel("➔")
        arrow_lbl.setStyleSheet("font-weight: bold; color: #64748B; font-size: 14px;")
        config_layout.addWidget(arrow_lbl)

        config_layout.addWidget(QLabel("<b>Target Release (New):</b>"))
        self.target_ver_combo = QComboBox()
        self.target_ver_combo.currentIndexChanged.connect(self._generate_diff_preview)
        config_layout.addWidget(self.target_ver_combo, stretch=1)

        clause_lbl = QLabel(f"<b>Clause:</b> {self.clause_number} ({self.clause_title})")
        clause_lbl.setStyleSheet("color: #475569; font-size: 11px;")
        config_layout.addWidget(clause_lbl, stretch=1)

        main_layout.addWidget(config_group)

        # ---------------------------------------------------------------------
        # Main Tabbed Area: [1] Visual Diff  |  [2] AI Analysis Prompt
        # ---------------------------------------------------------------------
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("""
            QTabBar::tab {
                font-weight: bold;
                padding: 6px 16px;
            }
        """)

        # Tab 1: Visual Diff
        self.tab_diff = QWidget()
        diff_layout = QVBoxLayout(self.tab_diff)
        diff_layout.setContentsMargins(8, 8, 8, 8)
        diff_layout.setSpacing(8)

        diff_top_bar = QHBoxLayout()
        self.diff_stats_lbl = QLabel("<b>Diff Summary:</b> Calculating...")
        self.diff_stats_lbl.setStyleSheet("font-size: 12px; color: #1E293B;")
        diff_top_bar.addWidget(self.diff_stats_lbl, stretch=1)

        self.btn_copy_diff = QPushButton("📋 Copy Diff")
        self.btn_copy_diff.setToolTip("Copy raw unified diff text to clipboard.")
        self.btn_copy_diff.setStyleSheet("padding: 4px 12px; font-weight: bold; background-color: #F1F5F9; border: 1px solid #CBD5E1; border-radius: 4px;")
        self.btn_copy_diff.clicked.connect(self._copy_diff)
        diff_top_bar.addWidget(self.btn_copy_diff)

        diff_layout.addLayout(diff_top_bar)

        self.diff_browser = QTextEdit()
        self.diff_browser.setReadOnly(True)
        self.diff_browser.setStyleSheet("""
            QTextEdit {
                background-color: #FAFAFA;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 11px;
                line-height: 1.45;
                color: #0F172A;
                padding: 8px;
            }
        """)
        diff_layout.addWidget(self.diff_browser)
        self.tab_widget.addTab(self.tab_diff, "📊 Visual Diff")

        # Tab 2: AI Analysis Prompt
        self.tab_prompt = QWidget()
        prompt_layout = QVBoxLayout(self.tab_prompt)
        prompt_layout.setContentsMargins(8, 8, 8, 8)
        prompt_layout.setSpacing(8)

        prompt_opts_card = QFrame()
        prompt_opts_card.setStyleSheet("background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 4px;")
        opts_vlayout = QVBoxLayout(prompt_opts_card)
        opts_vlayout.setContentsMargins(8, 8, 8, 8)
        opts_vlayout.setSpacing(6)

        row_focus = QHBoxLayout()
        row_focus.addWidget(QLabel("<b>Analysis Focus:</b>"))
        self.focus_combo = QComboBox()
        self.focus_combo.addItem("📋 General Standards & Functional Impact", "standards")
        self.focus_combo.addItem("⚖️ Patent & Prior Art Evaluation", "patent")
        self.focus_combo.addItem("📡 Signalling & Protocol Encoding", "signalling")
        self.focus_combo.currentIndexChanged.connect(self._generate_diff_preview)
        row_focus.addWidget(self.focus_combo, stretch=2)
        opts_vlayout.addLayout(row_focus)

        row_scope = QHBoxLayout()
        row_scope.addWidget(QLabel("<b>Context Scope:</b>"))
        self.tier_group = QButtonGroup(self)

        self.radio_tier1 = QRadioButton("Tier 1: Exact Clause Only")
        self.radio_tier1.setToolTip("Includes only the delta of the target clause.")
        self.tier_group.addButton(self.radio_tier1, 1)
        row_scope.addWidget(self.radio_tier1)

        self.radio_tier2 = QRadioButton("Tier 2: + Parent Procedure Scope (Recommended)")
        self.radio_tier2.setToolTip("Includes hierarchical breadcrumbs, parent intro/preconditions, and the clause delta.")
        self.radio_tier2.setChecked(True)
        self.tier_group.addButton(self.radio_tier2, 2)
        row_scope.addWidget(self.radio_tier2)

        self.radio_tier3 = QRadioButton("Tier 3: + Full Procedure Branch")
        self.radio_tier3.setToolTip("Includes sibling subclauses in the same procedure branch for deep architectural context.")
        self.tier_group.addButton(self.radio_tier3, 3)
        row_scope.addWidget(self.radio_tier3)

        row_scope.addStretch()
        self.tier_group.buttonClicked.connect(self._generate_diff_preview)
        opts_vlayout.addLayout(row_scope)

        prompt_layout.addWidget(prompt_opts_card)

        self.prompt_browser = QTextEdit()
        self.prompt_browser.setReadOnly(True)
        self.prompt_browser.setStyleSheet("""
            QTextEdit {
                background-color: #F8FAFC;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 11px;
                line-height: 1.4;
                color: #1E293B;
                padding: 6px;
            }
        """)
        prompt_layout.addWidget(self.prompt_browser)

        # Backward compatibility alias
        self.preview_browser = self.prompt_browser

        prompt_actions = QHBoxLayout()
        self.btn_copy_prompt = QPushButton("📋 Copy Prompt for LLM")
        self.btn_copy_prompt.setStyleSheet("font-weight: bold; background-color: #0284C7; color: white; padding: 6px 16px; border-radius: 4px;")
        self.btn_copy_prompt.clicked.connect(self._copy_prompt)
        prompt_actions.addWidget(self.btn_copy_prompt)

        self.btn_save = QPushButton("💾 Save Markdown (.md)")
        self.btn_save.setStyleSheet("font-weight: bold; background-color: #F1F5F9; color: #334155; border: 1px solid #CBD5E1; padding: 6px 14px; border-radius: 4px;")
        self.btn_save.clicked.connect(self._save_markdown)
        prompt_actions.addWidget(self.btn_save)

        prompt_actions.addStretch()
        prompt_layout.addLayout(prompt_actions)

        self.tab_widget.addTab(self.tab_prompt, "🤖 AI Analysis Prompt")
        main_layout.addWidget(self.tab_widget)

        # ---------------------------------------------------------------------
        # Bottom Bar: Word Comparison & Close
        # ---------------------------------------------------------------------
        btn_bar = QHBoxLayout()

        self.btn_word_compare = QPushButton("📝 Open in Word (Track Changes)")
        self.btn_word_compare.setToolTip(
            "Extract this clause from Base and Target releases and open Microsoft Word with native Track Changes (strikethrough deletions & underlined additions)."
        )
        self.btn_word_compare.setStyleSheet("""
            QPushButton {
                font-weight: bold;
                background-color: #E8F5E9;
                color: #2E7D32;
                border: 1px solid #A5D6A7;
                padding: 6px 16px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #C8E6C9;
                color: #1B5E20;
            }
            QPushButton:disabled {
                background-color: #F3F4F6;
                color: #9CA3AF;
                border: 1px solid #E5E7EB;
            }
        """)
        self.btn_word_compare.clicked.connect(self._start_word_comparison)
        btn_bar.addWidget(self.btn_word_compare)

        btn_bar.addStretch()
        self.btn_close = QPushButton("Close")
        self.btn_close.setStyleSheet("padding: 6px 18px;")
        self.btn_close.clicked.connect(self.accept)
        btn_bar.addWidget(self.btn_close)

        main_layout.addLayout(btn_bar)

    def _load_available_versions(self):
        """Populates version dropdowns for the specification."""
        vers = self.db.get_versions_for_spec(self.spec_number)
        if not vers:
            return

        self.base_ver_combo.blockSignals(True)
        self.target_ver_combo.blockSignals(True)

        self.base_ver_combo.clear()
        self.target_ver_combo.clear()

        for v in vers:
            ver_str = v["version"]
            date_str = f" ({v['release_date']})" if v.get("release_date") else ""
            label = f"v{ver_str}{date_str}"
            self.base_ver_combo.addItem(label, v)
            self.target_ver_combo.addItem(label, v)

        # Default Target Version to current_version
        target_idx = 0
        for i in range(self.target_ver_combo.count()):
            v_data = self.target_ver_combo.itemData(i)
            if v_data and v_data.get("version") == self.current_version:
                target_idx = i
                break
        self.target_ver_combo.setCurrentIndex(target_idx)

        # Default Base Version: use initial_base_version if provided, else preceding version
        base_idx = -1
        if self.initial_base_version:
            for i in range(self.base_ver_combo.count()):
                v_data = self.base_ver_combo.itemData(i)
                if v_data and v_data.get("version") == self.initial_base_version:
                    base_idx = i
                    break

        if base_idx < 0:
            base_idx = min(target_idx + 1, self.base_ver_combo.count() - 1)

        self.base_ver_combo.setCurrentIndex(base_idx)

        self.base_ver_combo.blockSignals(False)
        self.target_ver_combo.blockSignals(False)

    def _format_diff_to_html(
        self, diff_text: str, base_ver: str, target_ver: str
    ) -> Tuple[str, int, int]:
        """
        Converts unified diff text into syntax-highlighted HTML.
        Returns: (html_str, additions_count, deletions_count)
        """
        if not diff_text or "No text differences detected" in diff_text:
            empty_html = (
                f"<div style='font-family: Segoe UI, sans-serif; padding: 18px; color: #166534; "
                f"background-color: #F0FDF4; border: 1px solid #BBF7D0; border-radius: 6px; margin: 8px;'>"
                f"<h3 style='margin: 0 0 6px 0; font-size: 14px;'>✓ No Text Differences Detected</h3>"
                f"<p style='margin: 0; font-size: 12px; color: #15803D;'>"
                f"Clause {html.escape(self.clause_number)} is identical between "
                f"<b>3GPP TS {html.escape(self.spec_number)} v{html.escape(base_ver)}</b> and "
                f"<b>v{html.escape(target_ver)}</b>."
                f"</p></div>"
            )
            return empty_html, 0, 0

        html_lines = []
        additions = 0
        deletions = 0

        for line in diff_text.splitlines():
            escaped = html.escape(line)
            if line.startswith("+++ ") or line.startswith("--- "):
                html_lines.append(
                    f"<div style='color: #64748B; font-weight: bold; font-family: Consolas, monospace; line-height: 1.4;'>{escaped}</div>"
                )
            elif line.startswith("@@"):
                html_lines.append(
                    f"<div style='background-color: #E0E7FF; color: #3730A3; font-weight: bold; font-family: Consolas, monospace; "
                    f"padding: 3px 6px; margin: 4px 0; border-radius: 3px;'>{escaped}</div>"
                )
            elif line.startswith("+"):
                additions += 1
                html_lines.append(
                    f"<div style='background-color: #DCFCE7; color: #15803D; font-family: Consolas, monospace; padding: 1px 6px; white-space: pre-wrap;'>{escaped}</div>"
                )
            elif line.startswith("-"):
                deletions += 1
                html_lines.append(
                    f"<div style='background-color: #FEE2E2; color: #B91C1C; font-family: Consolas, monospace; padding: 1px 6px; white-space: pre-wrap;'>{escaped}</div>"
                )
            else:
                html_lines.append(
                    f"<div style='color: #1E293B; font-family: Consolas, monospace; padding: 1px 6px; white-space: pre-wrap;'>{escaped}</div>"
                )

        full_html = (
            f"<div style='font-family: Consolas, Courier New, monospace; font-size: 11px; line-height: 1.45;'>"
            + "".join(html_lines)
            + "</div>"
        )
        return full_html, additions, deletions

    def _generate_diff_preview(self):
        """Generates the unified diff and updates both the Visual Diff and the AI prompt preview."""
        base_data = self.base_ver_combo.currentData()
        target_data = self.target_ver_combo.currentData()

        if not base_data or not target_data:
            self.diff_stats_lbl.setText("<b>Diff Summary:</b> Select both Base and Target releases.")
            self.diff_browser.setText("Select both Base and Target specification releases.")
            self.prompt_browser.setText("Select both Base and Target specification releases.")
            return

        base_ver = base_data.get("version", "")
        base_date = base_data.get("release_date", "")
        target_ver = target_data.get("version", "")
        target_date = target_data.get("release_date", "")

        tier = self.tier_group.checkedId()
        focus_mode = self.focus_combo.currentData() or "standards"

        # Fetch Clause Text
        base_clause = self.db.get_clause_content_by_spec_ver(self.spec_number, base_ver, self.clause_number)
        target_clause = self.db.get_clause_content_by_spec_ver(self.spec_number, target_ver, self.clause_number)

        base_text = base_clause.get("content", "") if base_clause else ""
        target_text = target_clause.get("content", "") if target_clause else ""

        base_lbl = f"TS {self.spec_number} v{base_ver} (Clause {self.clause_number})"
        target_lbl = f"TS {self.spec_number} v{target_ver} (Clause {self.clause_number})"
        self._raw_diff_text = generate_unified_diff(base_text, target_text, base_lbl, target_lbl)

        # 1. Update Visual Diff Tab
        diff_html, additions, deletions = self._format_diff_to_html(self._raw_diff_text, base_ver, target_ver)
        self.diff_browser.setHtml(diff_html)

        if additions == 0 and deletions == 0:
            self.diff_stats_lbl.setText(
                f"<b>Diff Summary:</b> <span style='color: #166534; font-weight: bold;'>✓ Identical</span> "
                f"<span style='color: #64748B;'>(TS {self.spec_number} v{base_ver} ➔ v{target_ver})</span>"
            )
        else:
            self.diff_stats_lbl.setText(
                f"<b>Diff Summary:</b> "
                f"<span style='color: #15803D; font-weight: bold;'>+{additions} addition(s)</span> &nbsp;|&nbsp; "
                f"<span style='color: #B91C1C; font-weight: bold;'>-{deletions} deletion(s)</span> &nbsp;&nbsp; "
                f"<span style='color: #64748B;'>(TS {self.spec_number} v{base_ver} ➔ v{target_ver})</span>"
            )

        # 2. Update AI Analysis Prompt Tab
        hierarchy = None
        if tier >= 2:
            hierarchy = self.db.get_clause_hierarchy(self.spec_number, target_ver, self.clause_number)

        branch_clauses = None
        if tier >= 3:
            branch_clauses = self.db.get_branch_clauses(self.spec_number, target_ver, self.clause_number)

        self._generated_prompt = build_llm_clause_prompt(
            spec_number=self.spec_number,
            clause_number=self.clause_number,
            clause_title=self.clause_title,
            base_version=base_ver,
            base_date=base_date,
            target_version=target_ver,
            target_date=target_date,
            diff_text=self._raw_diff_text,
            tier=tier,
            hierarchy=hierarchy,
            branch_clauses=branch_clauses,
            focus_mode=focus_mode,
        )

        self.prompt_browser.setPlainText(self._generated_prompt)

    def _copy_diff(self):
        """Copies the raw unified diff text to clipboard."""
        if self._raw_diff_text:
            QApplication.clipboard().setText(self._raw_diff_text)
            self.btn_copy_diff.setText("✅ Copied Diff!")
            QTimer.singleShot(1500, lambda: self.btn_copy_diff.setText("📋 Copy Diff"))

    def _copy_prompt(self):
        """Copies the formatted LLM prompt to clipboard."""
        if self._generated_prompt:
            QApplication.clipboard().setText(self._generated_prompt)
            self.btn_copy_prompt.setText("✅ Copied to Clipboard!")
            QTimer.singleShot(1500, lambda: self.btn_copy_prompt.setText("📋 Copy Prompt for LLM"))

    def _save_markdown(self):
        """Saves the LLM prompt to a Markdown file."""
        if not self._generated_prompt:
            return
        default_name = f"TS_{self.spec_number}_Clause_{self.clause_number}_Diff.md".replace(".", "_")
        path, _ = QFileDialog.getSaveFileName(self, "Save Clause Diff Markdown", default_name, "Markdown Files (*.md)")
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(self._generated_prompt)
                QMessageBox.information(self, "Saved", f"Prompt saved successfully to:\n{path}")
            except Exception as e:
                QMessageBox.critical(self, "Save Error", f"Failed to write file:\n{e}")

    def _start_word_comparison(self):
        """Extracts this clause from Base and Target releases and runs native Word comparison with Track Changes."""
        base_data = self.base_ver_combo.currentData()
        target_data = self.target_ver_combo.currentData()
        if not base_data or not target_data:
            QMessageBox.warning(self, "Selection Required", "Please select both Base and Target specification releases.")
            return

        base_ver = base_data.get("version", "")
        target_ver = target_data.get("version", "")

        from modules.spec_search.core.spec_clause_word_comparator import SpecClauseWordCompareThread
        from PyQt5.QtWidgets import QProgressDialog

        prog = QProgressDialog(f"Comparing Clause {self.clause_number} in Microsoft Word (Track Changes)...", "Cancel", 0, 100, self)
        prog.setWindowTitle("Word Comparison (Track Changes)")
        prog.setWindowModality(Qt.WindowModal)
        prog.setValue(5)

        self._word_cmp_worker = SpecClauseWordCompareThread(
            spec_number=self.spec_number,
            base_version=base_ver,
            target_version=target_ver,
            clause_number=self.clause_number,
            clause_title=self.clause_title,
            cache_dir=self.cache_dir,
            specs_db=self.specs_db,
        )

        def on_progress(val, msg):
            prog.setValue(val)
            prog.setLabelText(msg)

        def on_success(msg):
            prog.close()
            logger.info(f"✅ {msg}")

        def on_error(err):
            prog.close()
            QMessageBox.critical(
                self,
                "Word Comparison Failed",
                f"Could not generate Microsoft Word comparison with Track Changes:\n{err}\n\n"
                f"Tip: You can still inspect the full text differences directly in the 'Visual Diff' tab."
            )

        self._word_cmp_worker.progress.connect(on_progress)
        self._word_cmp_worker.finished_success.connect(on_success)
        self._word_cmp_worker.error.connect(on_error)
        prog.canceled.connect(self._word_cmp_worker.requestInterruption)

        self._word_cmp_worker.start()
        prog.exec_()
