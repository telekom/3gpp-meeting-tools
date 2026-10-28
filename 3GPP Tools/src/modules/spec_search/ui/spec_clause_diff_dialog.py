# --- File: src/modules/spec_search/ui/spec_clause_diff_dialog.py ---
import html
import logging
from pathlib import Path
from typing import Optional, Any, Tuple, List, Dict

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QButtonGroup,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextBrowser,
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
      3. Tab 3 - Modifying TDoc Attribution: Zero-download metadata ranking and
         Track-Changes OpenXML text fingerprinting identifying originating CR/p-CR.
      4. Footer - Word Track Changes: 1-click extraction and launch into Microsoft Word
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
        focus_attribution: bool = False,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(None, Qt.Window)
        self.db = db
        self.spec_number = str(spec_number).strip()
        self.current_version = str(current_version).lstrip("v").strip()
        self.clause_number = str(clause_number).strip()
        self.clause_title = str(clause_title).strip()
        self.cache_dir = cache_dir
        self.specs_db = specs_db
        self.initial_base_version = str(base_version).lstrip("v").strip() if base_version else None
        self.focus_attribution = focus_attribution
        self._parent_ref = parent

        self._raw_diff_text = ""
        self._generated_prompt = ""
        self._word_cmp_worker = None

        self._attr_worker = None
        self._deep_verify_worker = None
        self._cached_attribution_results: Optional[List[Dict[str, Any]]] = None
        self._attribution_query_key = None
        self._candidate_list: List[Dict[str, Any]] = []

        self.setWindowTitle(f"🔍 Compare Clause {self.clause_number} — 3GPP TS {self.spec_number}")
        self.resize(1060, 750)
        self._setup_ui()
        self._load_available_versions()
        self._generate_diff_preview()

        if self.focus_attribution:
            self.tab_widget.setCurrentIndex(2)
            QTimer.singleShot(150, lambda: self._run_attribution(force=True))

    @property
    def version_b(self) -> str:
        if hasattr(self, "target_ver_combo") and self.target_ver_combo.currentData():
            return self.target_ver_combo.currentData().get("version", "")
        return getattr(self, "current_version", "")

    @property
    def version_a(self) -> str:
        if hasattr(self, "base_ver_combo") and self.base_ver_combo.currentData():
            return self.base_ver_combo.currentData().get("version", "")
        return getattr(self, "initial_base_version", "")

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

        self.btn_view_changes = QPushButton(f"📜 v{self.version_b} Changes / TDocs")
        self.btn_view_changes.setToolTip(f"View all integrated CRs, p-CRs, and TDocs that created v{self.version_b}, and directly open the source contribution documents.")
        self.btn_view_changes.setStyleSheet("padding: 4px 12px; font-weight: bold; background-color: #F0FDF4; color: #166534; border: 1px solid #86EFAC; border-radius: 4px;")
        self.btn_view_changes.setCursor(Qt.PointingHandCursor)
        self.btn_view_changes.clicked.connect(self._open_version_changes)
        diff_top_bar.addWidget(self.btn_view_changes)

        self.btn_find_tdoc = QPushButton("🎯 Identify Modifying TDoc")
        self.btn_find_tdoc.setToolTip("Identify which specific CR or p-CR introduced or modified this clause across releases.")
        self.btn_find_tdoc.setStyleSheet("padding: 4px 12px; font-weight: bold; background-color: #F0FDF4; color: #166534; border: 1px solid #86EFAC; border-radius: 4px;")
        self.btn_find_tdoc.setCursor(Qt.PointingHandCursor)
        self.btn_find_tdoc.clicked.connect(self._focus_and_run_attribution)
        diff_top_bar.addWidget(self.btn_find_tdoc)

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

        # Tab 3: Modifying TDoc Attribution
        self._setup_attribution_tab()
        self.tab_widget.currentChanged.connect(self._on_tab_changed)

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

        if hasattr(self, "btn_view_changes") and target_ver:
            self.btn_view_changes.setText(f"📜 v{target_ver} Changes / TDocs")
            self.btn_view_changes.setToolTip(f"View all integrated CRs, p-CRs, and TDocs that created v{target_ver}, and directly open the source contribution documents.")

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

        # Invalidate cached attribution when version selection changes
        self._cached_attribution_results = None
        if hasattr(self, "tab_widget") and self.tab_widget.currentIndex() == 2:
            self._run_attribution()

    def _copy_diff(self):
        """Copies the raw unified diff text to clipboard."""
        if self._raw_diff_text:
            QApplication.clipboard().setText(self._raw_diff_text)
            self.btn_copy_diff.setText("✅ Copied Diff!")
            QTimer.singleShot(1500, lambda: self.btn_copy_diff.setText("📋 Copy Diff"))

    def _open_version_changes(self):
        """Opens Change History dialog modelessly pre-filtered to version B with directly openable TDocs."""
        from modules.specifications.ui.spec_history_dialog import SpecVersionHistoryDialog
        logger.info(f"📜 [Change History] Opening Change History & TDocs for TS {self.spec_number} v{self.version_b}...")
        dlg = SpecVersionHistoryDialog(
            spec_number=self.spec_number,
            initial_version=self.version_b,
            parent=self.window() or self,
            specs_db=self.specs_db,
            initial_clause=self.clause_number,
        )
        self._history_dialog = dlg
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()

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

    # -------------------------------------------------------------------------
    # Tab 3: Modifying TDoc Attribution & OpenXML Fingerprinting
    # -------------------------------------------------------------------------

    def _setup_attribution_tab(self):
        """Sets up Tab 3: Modifying TDoc Attribution (Stage 1 Metadata & Stage 2 Document Fingerprinting)."""
        self.tab_attribution = QWidget()
        layout = QVBoxLayout(self.tab_attribution)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        # 1. Header Control Bar
        header_bar = QHBoxLayout()
        header_bar.setSpacing(8)

        self.attr_status_lbl = QLabel(f"<b>Clause Attribution:</b> Identifying originating CR/p-CR for Clause {self.clause_number}...")
        self.attr_status_lbl.setStyleSheet("font-size: 12px; color: #1E293B;")
        header_bar.addWidget(self.attr_status_lbl, stretch=1)

        self.btn_attr_run = QPushButton("🔄 Re-Analyze")
        self.btn_attr_run.setToolTip("Re-run multi-stage clause attribution for current releases.")
        self.btn_attr_run.setStyleSheet("padding: 4px 12px; font-weight: bold; background-color: #F1F5F9; border: 1px solid #CBD5E1; border-radius: 4px;")
        self.btn_attr_run.setCursor(Qt.PointingHandCursor)
        self.btn_attr_run.clicked.connect(lambda: self._run_attribution(force=True))
        header_bar.addWidget(self.btn_attr_run)

        self.btn_attr_deep_verify = QPushButton("🔬 Deep Verify Top 3 (DOCX)")
        self.btn_attr_deep_verify.setToolTip("Download and inspect Word OpenXML track changes in top 3 candidate documents for 100% deterministic text match.")
        self.btn_attr_deep_verify.setStyleSheet("padding: 4px 12px; font-weight: bold; background-color: #EFF6FF; color: #1D4ED8; border: 1px solid #93C5FD; border-radius: 4px;")
        self.btn_attr_deep_verify.setCursor(Qt.PointingHandCursor)
        self.btn_attr_deep_verify.setEnabled(False)
        self.btn_attr_deep_verify.clicked.connect(self._run_deep_verify)
        header_bar.addWidget(self.btn_attr_deep_verify)

        layout.addLayout(header_bar)

        # Progress Bar
        self.attr_progress_bar = QProgressBar()
        self.attr_progress_bar.setRange(0, 100)
        self.attr_progress_bar.setFixedHeight(8)
        self.attr_progress_bar.setTextVisible(False)
        self.attr_progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #CBD5E1;
                border-radius: 3px;
                background-color: #F1F5F9;
            }
            QProgressBar::chunk {
                background-color: #0284C7;
                border-radius: 2px;
            }
        """)
        self.attr_progress_bar.setVisible(False)
        layout.addWidget(self.attr_progress_bar)

        # 2. Hero Banner Card (Top Match)
        self.attr_hero_banner = QFrame()
        self.attr_hero_banner.setStyleSheet("""
            QFrame {
                background-color: #F0FDF4;
                border: 1px solid #BBF7D0;
                border-radius: 6px;
                padding: 6px;
            }
        """)
        self.attr_hero_banner.setVisible(False)
        hero_layout = QVBoxLayout(self.attr_hero_banner)
        hero_layout.setContentsMargins(8, 8, 8, 8)
        hero_layout.setSpacing(6)

        # Hero Top Row: Title + Action Buttons
        hero_row1 = QHBoxLayout()
        self.hero_match_title_lbl = QLabel("🎯 <b>Top Candidate:</b> Loading...")
        self.hero_match_title_lbl.setStyleSheet("font-size: 13px; color: #15803D;")
        hero_row1.addWidget(self.hero_match_title_lbl, stretch=1)

        self.btn_hero_open = QPushButton("📄 Open TDoc")
        self.btn_hero_open.setToolTip("Open contribution document (.docx) in Microsoft Word.")
        self.btn_hero_open.setStyleSheet("padding: 3px 10px; font-weight: bold; background-color: #166534; color: white; border-radius: 4px;")
        self.btn_hero_open.setCursor(Qt.PointingHandCursor)
        hero_row1.addWidget(self.btn_hero_open)

        self.btn_hero_jump = QPushButton("🚀 Jump to Meetings")
        self.btn_hero_jump.setToolTip("Navigate to Meetings tab with this TDoc focused.")
        self.btn_hero_jump.setStyleSheet("padding: 3px 10px; font-weight: bold; background-color: #F1F5F9; color: #334155; border: 1px solid #CBD5E1; border-radius: 4px;")
        self.btn_hero_jump.setCursor(Qt.PointingHandCursor)
        hero_row1.addWidget(self.btn_hero_jump)

        self.btn_hero_copy = QPushButton("📋 Copy Citation")
        self.btn_hero_copy.setToolTip("Copy plain text and rich HTML citation.")
        self.btn_hero_copy.setStyleSheet("padding: 3px 10px; font-weight: bold; background-color: #F1F5F9; color: #334155; border: 1px solid #CBD5E1; border-radius: 4px;")
        self.btn_hero_copy.setCursor(Qt.PointingHandCursor)
        hero_row1.addWidget(self.btn_hero_copy)

        hero_layout.addLayout(hero_row1)

        self.hero_doc_title_lbl = QLabel("")
        self.hero_doc_title_lbl.setStyleSheet("font-size: 11px; color: #334155; font-style: italic;")
        self.hero_doc_title_lbl.setWordWrap(True)
        hero_layout.addWidget(self.hero_doc_title_lbl)

        self.hero_badges_layout = QHBoxLayout()
        self.hero_badges_layout.setSpacing(6)
        hero_layout.addLayout(self.hero_badges_layout)

        layout.addWidget(self.attr_hero_banner)

        # 3. Middle & Bottom Splitter: Candidates Table & Details Browser
        splitter = QSplitter(Qt.Vertical)

        # Table of Candidates
        self.attr_table = QTableWidget(0, 6)
        self.attr_table.setHorizontalHeaderLabels(["Confidence", "TDoc / CR #", "Source", "Title / Subject", "Evidence Signals", "Actions"])
        self.attr_table.horizontalHeader().setStretchLastSection(False)
        self.attr_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.attr_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.attr_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.attr_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.attr_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.attr_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.attr_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.attr_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.attr_table.verticalHeader().setVisible(False)
        self.attr_table.itemSelectionChanged.connect(self._on_candidate_table_selection_changed)
        self.attr_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                background-color: #FFFFFF;
                font-size: 11px;
                selection-background-color: #EBF3FC;
                selection-color: #0F172A;
            }
            QHeaderView::section {
                background-color: #F8FAFC;
                font-weight: bold;
                color: #475569;
                padding: 4px 6px;
                border: none;
                border-bottom: 1px solid #CBD5E1;
            }
        """)
        splitter.addWidget(self.attr_table)

        # Details Browser
        self.attr_details_browser = QTextBrowser()
        self.attr_details_browser.setOpenExternalLinks(True)
        self.attr_details_browser.setPlaceholderText("Select a candidate contribution above to inspect full metadata, meeting agenda context, and verified text excerpts...")
        self.attr_details_browser.setStyleSheet("""
            QTextBrowser {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                background-color: #F8FAFC;
                padding: 6px;
                font-size: 11px;
                line-height: 1.4;
            }
        """)
        splitter.addWidget(self.attr_details_browser)
        splitter.setSizes([300, 160])

        layout.addWidget(splitter)
        self.tab_widget.addTab(self.tab_attribution, "🎯 Modifying TDoc")

    def _focus_and_run_attribution(self):
        """Switches to the Modifying TDoc tab and triggers attribution."""
        self.tab_widget.setCurrentIndex(2)
        self._run_attribution()

    def _on_tab_changed(self, index: int):
        """Trigger attribution lazily when user navigates to the Attribution tab."""
        if index == 2:
            query_key = (self.spec_number, self.clause_number, self.version_b, self.version_a)
            if self._attribution_query_key != query_key or self._cached_attribution_results is None:
                self._run_attribution()

    def _run_attribution(self, force: bool = False):
        """Executes multi-stage clause attribution via background thread."""
        query_key = (self.spec_number, self.clause_number, self.version_b, self.version_a)
        if not force and self._attribution_query_key == query_key and self._cached_attribution_results is not None:
            return

        if not self.version_b:
            self.attr_status_lbl.setText("Select target specification release first.")
            return

        self._attribution_query_key = query_key
        self.attr_progress_bar.setVisible(True)
        self.attr_progress_bar.setValue(10)
        self.attr_status_lbl.setText(f"Analyzing candidates for TS {self.spec_number} Clause {self.clause_number} in v{self.version_b}...")
        self.btn_attr_run.setEnabled(False)
        self.btn_attr_deep_verify.setEnabled(False)

        target_clause = self.db.get_clause_content_by_spec_ver(self.spec_number, self.version_b, self.clause_number)
        target_content = target_clause.get("content", "") if target_clause else ""

        if not hasattr(self, "attr_engine"):
            from modules.specifications.core.spec_clause_attribution import SpecClauseAttributionEngine
            self.attr_engine = SpecClauseAttributionEngine(search_db=self.db)

        from modules.specifications.core.spec_clause_attribution import SpecClauseAttributionThread
        self._attr_worker = SpecClauseAttributionThread(
            engine=self.attr_engine,
            spec_number=self.spec_number,
            clause_number=self.clause_number,
            target_version=self.version_b,
            base_version=self.version_a,
            clause_title=self.clause_title,
            diff_text=self._raw_diff_text,
            clause_content=target_content,
            parent=self,
        )
        self._attr_worker.progress.connect(self._on_attribution_progress)
        self._attr_worker.finished.connect(self._on_attribution_finished)
        self._attr_worker.error.connect(self._on_attribution_error)
        self._attr_worker.start()

    def _on_attribution_progress(self, val: int, msg: str):
        self.attr_progress_bar.setValue(val)
        self.attr_status_lbl.setText(f"<b>Attribution:</b> {html.escape(msg)}")

    def _on_attribution_finished(self, candidates: List[Dict[str, Any]]):
        self.attr_progress_bar.setVisible(False)
        self.btn_attr_run.setEnabled(True)
        self._cached_attribution_results = candidates
        self._populate_attribution_view(candidates)

    def _on_attribution_error(self, err_msg: str):
        self.attr_progress_bar.setVisible(False)
        self.btn_attr_run.setEnabled(True)
        self.attr_status_lbl.setText(f"<span style='color: #B91C1C;'><b>Attribution Error:</b> {html.escape(err_msg)}</span>")

    def _populate_attribution_view(self, candidates: List[Dict[str, Any]]):
        self._candidate_list = list(candidates)
        if not candidates:
            self.attr_hero_banner.setVisible(False)
            self.attr_table.setRowCount(0)
            self.attr_status_lbl.setText(f"No candidate CRs or p-CRs found for TS {self.spec_number} v{self.version_b}.")
            self.attr_details_browser.setHtml(
                f"<div style='color: #64748B; padding: 12px; font-family: Segoe UI, sans-serif;'>"
                f"No change history records or candidate TDocs could be resolved for "
                f"<b>3GPP TS {html.escape(self.spec_number)} v{html.escape(self.version_b)}</b>.<br/><br/>"
                f"Tip: Ensure specification change history is indexed in <code>3gpp_spec_history.db</code>."
                f"</div>"
            )
            self.btn_attr_deep_verify.setEnabled(False)
            return

        self.btn_attr_deep_verify.setEnabled(True)
        self.attr_status_lbl.setText(
            f"<b>Found {len(candidates)} candidate contribution(s)</b> for Clause {self.clause_number} in v{self.version_b}:"
        )

        # 1. Hero banner for top candidate
        top = candidates[0]
        top_tdoc = top.get("tdoc", "")
        top_source = top.get("source", "")
        top_score = top.get("score", 0)
        top_conf = top.get("confidence", "High")
        top_verified = top.get("is_verified", False)

        if top_verified:
            self.attr_hero_banner.setStyleSheet("QFrame { background-color: #F0FDF4; border: 1px solid #86EFAC; border-radius: 6px; padding: 6px; }")
            self.hero_match_title_lbl.setStyleSheet("font-size: 13px; color: #15803D;")
            self.hero_match_title_lbl.setText(f"🎯 <b>Top Match (Verified):</b> {html.escape(top_tdoc)} {f'({html.escape(top_source)})' if top_source else ''} · {top_score}% Match")
        else:
            self.attr_hero_banner.setStyleSheet("QFrame { background-color: #F8FAFC; border: 1px solid #CBD5E1; border-radius: 6px; padding: 6px; }")
            self.hero_match_title_lbl.setStyleSheet("font-size: 13px; color: #0369A1;")
            self.hero_match_title_lbl.setText(f"🎯 <b>Top Candidate:</b> {html.escape(top_tdoc)} {f'({html.escape(top_source)})' if top_source else ''} · {top_conf} Confidence ({top_score}%)")

        self.hero_doc_title_lbl.setText(top.get("title", "") or "No title available")

        # Clear existing badges
        while self.hero_badges_layout.count():
            item = self.hero_badges_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        for b_text in top.get("evidence_badges", []):
            b_lbl = QLabel(b_text)
            b_lbl.setStyleSheet("background-color: #E0F2FE; color: #0369A1; border: 1px solid #BAE6FD; padding: 2px 8px; border-radius: 10px; font-size: 10px; font-weight: bold;")
            self.hero_badges_layout.addWidget(b_lbl)
        self.hero_badges_layout.addStretch()

        # Connect Hero Action Buttons
        try:
            self.btn_hero_open.clicked.disconnect()
        except Exception:
            pass
        self.btn_hero_open.clicked.connect(lambda: self._open_candidate_tdoc(top_tdoc, top.get("meeting", "")))

        try:
            self.btn_hero_jump.clicked.disconnect()
        except Exception:
            pass
        self.btn_hero_jump.clicked.connect(lambda: self._jump_candidate_to_meetings(top_tdoc))

        try:
            self.btn_hero_copy.clicked.disconnect()
        except Exception:
            pass
        self.btn_hero_copy.clicked.connect(lambda: self._copy_candidate_citation(top_tdoc, top.get("meeting", ""), top.get("title", "")))

        self.attr_hero_banner.setVisible(True)

        # 2. Populate Candidate Table
        self.attr_table.setRowCount(len(candidates))
        for row_idx, cand in enumerate(candidates):
            tdoc_tok = cand.get("tdoc", "")
            cr_num = cand.get("cr_number", "")
            tdoc_display = tdoc_tok if not cr_num else f"{tdoc_tok} (CR {cr_num})"
            src = cand.get("source", "")
            title = cand.get("title", "")
            score = cand.get("score", 0)
            conf = cand.get("confidence", "Low")
            verified = cand.get("is_verified", False)
            badges = cand.get("evidence_badges", [])

            # Col 0: Confidence Pill
            if verified:
                conf_text = f"✅ {score}% Verified"
            elif score >= 80:
                conf_text = f"🟢 {score}% {conf}"
            elif score >= 55:
                conf_text = f"🔵 {score}% {conf}"
            elif score >= 35:
                conf_text = f"🟠 {score}% {conf}"
            else:
                conf_text = f"⚪ {score}% {conf}"
            item_conf = QTableWidgetItem(conf_text)
            item_conf.setFlags(item_conf.flags() & ~Qt.ItemIsEditable)
            self.attr_table.setItem(row_idx, 0, item_conf)

            # Col 1: TDoc
            item_tdoc = QTableWidgetItem(tdoc_display)
            item_tdoc.setFlags(item_tdoc.flags() & ~Qt.ItemIsEditable)
            item_tdoc.setFont(QFont("Consolas", 9, QFont.Bold))
            self.attr_table.setItem(row_idx, 1, item_tdoc)

            # Col 2: Source
            item_src = QTableWidgetItem(src)
            item_src.setFlags(item_src.flags() & ~Qt.ItemIsEditable)
            self.attr_table.setItem(row_idx, 2, item_src)

            # Col 3: Title
            item_title = QTableWidgetItem(title)
            item_title.setFlags(item_title.flags() & ~Qt.ItemIsEditable)
            item_title.setToolTip(title)
            self.attr_table.setItem(row_idx, 3, item_title)

            # Col 4: Evidence Signals
            evidence_str = " • ".join(badges)
            item_ev = QTableWidgetItem(evidence_str)
            item_ev.setFlags(item_ev.flags() & ~Qt.ItemIsEditable)
            item_ev.setToolTip(evidence_str)
            self.attr_table.setItem(row_idx, 4, item_ev)

            # Col 5: Actions
            action_widget = QWidget()
            act_layout = QHBoxLayout(action_widget)
            act_layout.setContentsMargins(2, 2, 2, 2)
            act_layout.setSpacing(4)

            btn_open = QPushButton("📄")
            btn_open.setToolTip(f"Open {tdoc_tok} in Microsoft Word")
            btn_open.setStyleSheet("padding: 2px 6px; font-size: 10px; font-weight: bold; background-color: #F1F5F9; border: 1px solid #CBD5E1; border-radius: 3px;")
            btn_open.clicked.connect(lambda _, t=tdoc_tok, m=cand.get("meeting", ""): self._open_candidate_tdoc(t, m))
            act_layout.addWidget(btn_open)

            btn_jump = QPushButton("🚀")
            btn_jump.setToolTip(f"Jump to Meetings tab for {tdoc_tok}")
            btn_jump.setStyleSheet("padding: 2px 6px; font-size: 10px; font-weight: bold; background-color: #F1F5F9; border: 1px solid #CBD5E1; border-radius: 3px;")
            btn_jump.clicked.connect(lambda _, t=tdoc_tok: self._jump_candidate_to_meetings(t))
            act_layout.addWidget(btn_jump)

            btn_cp = QPushButton("📋")
            btn_cp.setToolTip(f"Copy citation for {tdoc_tok}")
            btn_cp.setStyleSheet("padding: 2px 6px; font-size: 10px; font-weight: bold; background-color: #F1F5F9; border: 1px solid #CBD5E1; border-radius: 3px;")
            btn_cp.clicked.connect(lambda _, t=tdoc_tok, m=cand.get("meeting", ""), s=title: self._copy_candidate_citation(t, m, s))
            act_layout.addWidget(btn_cp)

            self.attr_table.setCellWidget(row_idx, 5, action_widget)

        self.attr_table.selectRow(0)
        self._on_candidate_table_selection_changed()

    def _on_candidate_table_selection_changed(self):
        """Displays rich metadata and snippet for currently selected candidate."""
        selected_rows = self.attr_table.selectedIndexes()
        if not selected_rows:
            return
        row = selected_rows[0].row()
        if row < 0 or row >= len(self._candidate_list):
            return
        cand = self._candidate_list[row]

        tdoc = cand.get("tdoc", "")
        cr_num = cand.get("cr_number", "")
        source = cand.get("source", "Unknown")
        meeting = cand.get("meeting", "")
        agenda_item = cand.get("agenda_item", "")
        title = cand.get("title", "")
        abstract = cand.get("abstract", "")
        remarks = cand.get("remarks", "")
        score = cand.get("score", 0)
        confidence = cand.get("confidence", "Low")
        badges = cand.get("evidence_badges", [])
        is_verified = cand.get("is_verified", False)
        match_pct = cand.get("verified_match_pct", 0)
        snippet = cand.get("verified_snippet", "")

        badges_html = " ".join([
            f"<span style='background-color: #E0F2FE; color: #0369A1; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: bold;'>{html.escape(b)}</span>"
            for b in badges
        ])

        verified_html = ""
        if is_verified:
            verified_html = f"""
            <div style='background-color: #DCFCE7; border: 1px solid #86EFAC; border-radius: 4px; padding: 6px 10px; margin: 8px 0;'>
                <span style='color: #15803D; font-weight: bold;'>✅ Deterministic OpenXML Text Match ({match_pct}%):</span><br/>
                <code style='font-family: Consolas, monospace; font-size: 11px; color: #166534;'>"{html.escape(snippet)}"</code>
            </div>
            """

        details_html = f"""
        <div style='font-family: Segoe UI, sans-serif;'>
            <div style='margin-bottom: 6px;'>
                <b style='font-size: 13px; color: #0284C7;'>{html.escape(tdoc)}</b>
                {f" <span style='color: #64748B;'>(CR #{html.escape(cr_num)})</span>" if cr_num else ""}
                &nbsp;—&nbsp; <span style='font-weight: 600; color: #334155;'>{html.escape(source)}</span>
                &nbsp;•&nbsp; <span style='color: #64748B;'>Meeting: <b>{html.escape(meeting or 'N/A')}</b></span>
                {f" &nbsp;•&nbsp; <span style='color: #64748B;'>AI: <b>{html.escape(agenda_item)}</b></span>" if agenda_item else ""}
                &nbsp;|&nbsp; <span style='color: #0369A1; font-weight: bold;'>Score: {score}% ({confidence})</span>
            </div>
            <div style='margin-bottom: 6px; font-size: 12px; font-weight: 600; color: #1E293B;'>
                {html.escape(title or 'No title available')}
            </div>
            <div style='margin-bottom: 8px;'>{badges_html}</div>
            {verified_html}
            {f"<div style='margin-top: 6px; color: #475569;'><b>Abstract / Summary:</b> {html.escape(abstract)}</div>" if abstract else ""}
            {f"<div style='margin-top: 4px; color: #64748B; font-size: 10px;'><b>Secretary Remarks:</b> {html.escape(remarks)}</div>" if remarks else ""}
        </div>
        """
        self.attr_details_browser.setHtml(details_html)

    def _run_deep_verify(self):
        """Runs targeted Stage 2 OpenXML fingerprinting on top 3 candidates via background thread."""
        if not self._candidate_list or not self._raw_diff_text:
            return

        self.btn_attr_deep_verify.setEnabled(False)
        self.attr_progress_bar.setVisible(True)
        self.attr_progress_bar.setValue(10)
        self.attr_status_lbl.setText("Running targeted Stage 2 OpenXML text fingerprinting...")

        if not hasattr(self, "attr_engine"):
            from modules.specifications.core.spec_clause_attribution import SpecClauseAttributionEngine
            self.attr_engine = SpecClauseAttributionEngine(search_db=self.db)

        from modules.specifications.core.spec_clause_attribution import SpecClauseDeepVerifyThread
        self._deep_verify_worker = SpecClauseDeepVerifyThread(
            engine=self.attr_engine,
            candidates=self._candidate_list,
            diff_text=self._raw_diff_text,
            parent=self,
        )

        def _on_dv_progress(val: int, msg: str):
            self.attr_progress_bar.setValue(val)
            self.attr_status_lbl.setText(f"<b>Deep Verify:</b> {html.escape(msg)}")

        def _on_dv_finished(updated_cands: List[Dict[str, Any]]):
            self.attr_progress_bar.setVisible(False)
            self.btn_attr_deep_verify.setEnabled(True)
            self._cached_attribution_results = updated_cands
            self._populate_attribution_view(updated_cands)

        def _on_dv_error(err: str):
            self.attr_progress_bar.setVisible(False)
            self.btn_attr_deep_verify.setEnabled(True)
            self.attr_status_lbl.setText(f"<span style='color: #B91C1C;'><b>Deep Verify Error:</b> {html.escape(err)}</span>")

        self._deep_verify_worker.progress.connect(_on_dv_progress)
        self._deep_verify_worker.finished.connect(_on_dv_finished)
        self._deep_verify_worker.error.connect(_on_dv_error)
        self._deep_verify_worker.start()

    def _open_candidate_tdoc(self, tdoc_id: str, meeting: str = ""):
        from modules.specifications.core.spec_tdoc_opener import SpecTDocOpener
        SpecTDocOpener.open_tdoc(tdoc_id, meeting_hint=meeting, parent=self)

    def _jump_candidate_to_meetings(self, tdoc_id: str):
        from modules.specifications.core.spec_tdoc_opener import SpecTDocOpener
        SpecTDocOpener.jump_to_meetings_tab(tdoc_id)

    def _copy_candidate_citation(self, tdoc_id: str, meeting: str = "", title: str = ""):
        from modules.specifications.core.spec_tdoc_opener import SpecTDocOpener
        SpecTDocOpener.copy_tdoc_citation(
            tdoc_id=tdoc_id,
            spec_number=self.spec_number,
            version=self.version_b,
            meeting=meeting,
            subject=title,
            parent=self,
        )

