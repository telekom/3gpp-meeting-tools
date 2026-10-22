# --- File: src/modules/spec_search/ui/spec_clause_diff_dialog.py ---
import logging
from pathlib import Path
from typing import Optional, Any

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import (
    QApplication,
    QButtonGroup,
    QComboBox,
    QDialog,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from modules.spec_search.core.spec_clause_diff import build_llm_clause_prompt, generate_unified_diff
from modules.spec_search.core.spec_search_db import SpecSearchDatabase

logger = logging.getLogger(__name__)

class SpecClauseDiffDialog(QDialog):
    """
    Interactive Dialog to configure clause comparison parameters, select context depth tiers,
    preview the structured Markdown diff, and export or copy prompts for LLM analysis.
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
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.db = db
        self.spec_number = spec_number
        self.current_version = current_version
        self.clause_number = clause_number
        self.clause_title = clause_title
        self.cache_dir = cache_dir
        self.specs_db = specs_db
        self._generated_prompt = ""
        self._word_cmp_worker = None

        self.setWindowTitle(f"🤖 Compare Clause {self.clause_number} for LLM Analysis (TS {self.spec_number})")
        self.resize(960, 680)
        self._setup_ui()
        self._load_available_versions()
        self._generate_diff_preview()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # ---------------------------------------------------------------------
        # Top Controls: Version Selectors & Analysis Focus
        # ---------------------------------------------------------------------
        config_group = QGroupBox("Comparison Configuration")
        config_layout = QVBoxLayout(config_group)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("<b>Base Release (Old):</b>"))
        self.base_ver_combo = QComboBox()
        self.base_ver_combo.currentIndexChanged.connect(self._generate_diff_preview)
        row1.addWidget(self.base_ver_combo, stretch=1)

        row1.addWidget(QLabel("<b>Target Release (New):</b>"))
        self.target_ver_combo = QComboBox()
        self.target_ver_combo.currentIndexChanged.connect(self._generate_diff_preview)
        row1.addWidget(self.target_ver_combo, stretch=1)

        row1.addWidget(QLabel("<b>Analysis Focus:</b>"))
        self.focus_combo = QComboBox()
        self.focus_combo.addItem("📋 General Standards & Functional Impact", "standards")
        self.focus_combo.addItem("⚖️ Patent & Prior Art Evaluation", "patent")
        self.focus_combo.addItem("📡 Signalling & Protocol Encoding", "signalling")
        self.focus_combo.currentIndexChanged.connect(self._generate_diff_preview)
        row1.addWidget(self.focus_combo, stretch=2)
        config_layout.addLayout(row1)

        # Context Depth Tier Radios
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("<b>Context Scope:</b>"))
        self.tier_group = QButtonGroup(self)

        self.radio_tier1 = QRadioButton("Tier 1: Exact Clause Only")
        self.radio_tier1.setToolTip("Includes only the delta of the target clause.")
        self.tier_group.addButton(self.radio_tier1, 1)
        row2.addWidget(self.radio_tier1)

        self.radio_tier2 = QRadioButton("Tier 2: + Parent Procedure Scope (Recommended)")
        self.radio_tier2.setToolTip("Includes hierarchical breadcrumbs, parent intro/preconditions, and the clause delta.")
        self.radio_tier2.setChecked(True)
        self.tier_group.addButton(self.radio_tier2, 2)
        row2.addWidget(self.radio_tier2)

        self.radio_tier3 = QRadioButton("Tier 3: + Full Procedure Branch")
        self.radio_tier3.setToolTip("Includes sibling subclauses in the same procedure branch for deep architectural context.")
        self.tier_group.addButton(self.radio_tier3, 3)
        row2.addWidget(self.radio_tier3)

        row2.addStretch()
        self.tier_group.buttonClicked.connect(self._generate_diff_preview)
        config_layout.addLayout(row2)

        layout.addWidget(config_group)

        # ---------------------------------------------------------------------
        # Markdown Preview Area
        # ---------------------------------------------------------------------
        preview_header = QHBoxLayout()
        preview_header.addWidget(QLabel("<b>Generated Prompt Preview for LLM:</b>"))
        preview_header.addStretch()
        layout.addLayout(preview_header)

        self.preview_browser = QTextEdit()
        self.preview_browser.setReadOnly(True)
        self.preview_browser.setStyleSheet("""
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
        layout.addWidget(self.preview_browser)

        # ---------------------------------------------------------------------
        # Action Buttons
        # ---------------------------------------------------------------------
        btn_bar = QHBoxLayout()

        self.btn_copy = QPushButton("📋 Copy Prompt for LLM")
        self.btn_copy.setStyleSheet("font-weight: bold; background-color: #0284C7; color: white; padding: 6px 16px; border-radius: 4px;")
        self.btn_copy.clicked.connect(self._copy_prompt)
        btn_bar.addWidget(self.btn_copy)

        self.btn_save = QPushButton("💾 Save Markdown (.md)")
        self.btn_save.setStyleSheet("font-weight: bold; background-color: #F1F5F9; color: #334155; border: 1px solid #CBD5E1; padding: 6px 14px; border-radius: 4px;")
        self.btn_save.clicked.connect(self._save_markdown)
        btn_bar.addWidget(self.btn_save)

        self.btn_word_compare = QPushButton("📝 Compare in Word (Redline)")
        self.btn_word_compare.setToolTip("Extract this clause from both specification releases and generate a native Word document with tracked redlines (strikethrough deletions & underlined insertions).")
        self.btn_word_compare.setStyleSheet("""
            QPushButton {
                font-weight: bold;
                background-color: #E8F5E9;
                color: #2E7D32;
                border: 1px solid #A5D6A7;
                padding: 6px 14px;
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
        self.btn_close.clicked.connect(self.accept)
        btn_bar.addWidget(self.btn_close)

        layout.addLayout(btn_bar)

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

        # Default Base Version to the preceding version
        base_idx = min(target_idx + 1, self.base_ver_combo.count() - 1)
        self.base_ver_combo.setCurrentIndex(base_idx)

        self.base_ver_combo.blockSignals(False)
        self.target_ver_combo.blockSignals(False)

    def _generate_diff_preview(self):
        """Generates the unified diff and updates the markdown preview."""
        base_data = self.base_ver_combo.currentData()
        target_data = self.target_ver_combo.currentData()

        if not base_data or not target_data:
            self.preview_browser.setText("Select both Base and Target specification releases.")
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
        diff_text = generate_unified_diff(base_text, target_text, base_lbl, target_lbl)

        # Fetch Hierarchy Context if requested
        hierarchy = None
        if tier >= 2:
            hierarchy = self.db.get_clause_hierarchy(self.spec_number, target_ver, self.clause_number)

        # Fetch Branch Context if requested
        branch_clauses = None
        if tier >= 3:
            branch_clauses = self.db.get_branch_clauses(self.spec_number, target_ver, self.clause_number)

        # Build Markdown Prompt
        self._generated_prompt = build_llm_clause_prompt(
            spec_number=self.spec_number,
            clause_number=self.clause_number,
            clause_title=self.clause_title,
            base_version=base_ver,
            base_date=base_date,
            target_version=target_ver,
            target_date=target_date,
            diff_text=diff_text,
            tier=tier,
            hierarchy=hierarchy,
            branch_clauses=branch_clauses,
            focus_mode=focus_mode,
        )

        self.preview_browser.setPlainText(self._generated_prompt)

    def _copy_prompt(self):
        if self._generated_prompt:
            QApplication.clipboard().setText(self._generated_prompt)
            self.btn_copy.setText("✅ Copied to Clipboard!")
            QTimer.singleShot(1500, lambda: self.btn_copy.setText("📋 Copy Prompt for LLM"))

    def _save_markdown(self):
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
        """Extracts this clause from Base and Target releases and runs native Word comparison."""
        base_data = self.base_ver_combo.currentData()
        target_data = self.target_ver_combo.currentData()
        if not base_data or not target_data:
            QMessageBox.warning(self, "Selection Required", "Please select both Base and Target specification releases.")
            return

        base_ver = base_data.get("version", "")
        target_ver = target_data.get("version", "")

        from modules.spec_search.core.spec_clause_word_comparator import SpecClauseWordCompareThread
        from PyQt5.QtWidgets import QProgressDialog

        prog = QProgressDialog(f"Comparing Clause {self.clause_number} in Microsoft Word...", "Cancel", 0, 100, self)
        prog.setWindowTitle("Word Redline Generation")
        prog.setWindowModality(1)  # Qt.WindowModal
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
            QMessageBox.critical(self, "Word Comparison Failed", f"Could not generate Word redline:\n{err}")

        self._word_cmp_worker.progress.connect(on_progress)
        self._word_cmp_worker.finished_success.connect(on_success)
        self._word_cmp_worker.error.connect(on_error)
        prog.canceled.connect(self._word_cmp_worker.requestInterruption)

        self._word_cmp_worker.start()
        prog.exec_()
